import dotenv from 'dotenv';
dotenv.config();
// import Mailer from '../includes/mailer.js';
import * as cheerio from 'cheerio';

import path from 'path';
import fs from 'fs';
// import { glob } from 'glob';
import _ from 'lodash';
import {z} from 'zod';
import https from 'https';
import Graph from '../includes/graphapi.js';
import Handlebars from 'handlebars';

import { 
  HumanMessage, 
  AIMessage, 
  SystemMessage 
} from "@langchain/core/messages";
import { Expression } from '@langchain/core/structured_query';

export default class Routes {

	constructor(_server, _llms) {
    this.server = _server;
		this._excelLock = Promise.resolve();
		this.model = _llms;

		// this.mailer = new Mailer();

		// --------------------------------------------
		// Dashboard: REST REQUESTS
		// --------------------------------------------

		this.server.app.post("/testnode", async (req, res) => {
			const token = await this.get_bearer_token(req);
			console.log('token length: ', token.length);
			if(token == process.env.ACTBOT_BEARER){

				console.log('testnode: token accepted');
				return res.send('success');
			} else {
        return res.status(403).send('unauthorized access');
      }
		});

		// OneDrive file manipulation
		// --------------------------------------------

		this.server.app.post("/onedrive/files/read", async (req, res) => {
			const token = await this.get_bearer_token(req);
			if(token != process.env.ACTBOT_BEARER){
        return res.status(403).send('unauthorized access');
			}

			const onedrive_user_id = req.body.od_user_id;
			const onedrive_folder  = req.body.od_folder_id;

			let graphapi = new Graph();
			let client = graphapi.initializeGraphForUserAuth();
			let children = [];

			try {
				let response = await client.api(`/users/${onedrive_user_id}/drive/items/${onedrive_folder}/children`).get();
				children = response.value;
			} catch(error) {
				console.error('Error: ', error);
				return res.status(500).json({ error: error.message });
			}

			try {
				for (const item of children) {
					return res.status(200).json({results: children});
				}
			} catch(error){
				console.error('File read error: ', error);
				return res.status(500).json({ error: error.message });
			}

		});

		this.server.app.post("/onedrive/files/download", async (req, res) => {

			const token = await this.get_bearer_token(req);
			if(token != process.env.ACTBOT_BEARER){
        return res.status(403).send('unauthorized access');
			}

			const onedrive_user_id = req.body.od_user_id;
			const onedrive_folder  = req.body.od_folder_id;

			let graphapi = new Graph();
			let client = graphapi.initializeGraphForUserAuth();
			let children = [];

			try {
				let response = await client.api(`/users/${onedrive_user_id}/drive/items/${onedrive_folder}/children`).get();
				children = response.value;
			} catch(error) {
				console.error('Error: ', error);
				return res.status(500).json({ error: error.message });
			}

			const documentsDir = path.join(`/home/node/`, `.n8n-files`);
			if (!fs.existsSync(documentsDir)) {
				fs.mkdirSync(documentsDir, { recursive: true });
			}

			const downloaded = [];
			try {
				for (const item of children) {
					const downloadUrl = item['@microsoft.graph.downloadUrl'];
					if (!downloadUrl) continue;

					const fileResponse = await fetch(downloadUrl);

					if (!fileResponse.ok) {
						throw new Error(`Failed to download ${item.name}: ${fileResponse.statusText}`);
					}

					if(item.hasOwnProperty('folder')) 
						continue;

					const buffer = await fileResponse.arrayBuffer();
					const filePath = path.join(documentsDir, item.name);

					fs.writeFileSync(filePath, Buffer.from(buffer));
					downloaded.push({ id: item.id, name: item.name, path: filePath });
				}
			} catch(error) {
				console.error('Download error: ', error);
				return res.status(500).json({ error: error.message });
			}
			return res.status(200).json({ files: downloaded });
		});

		this.server.app.post("/onedrive/files/move", async (req, res) => {
			
			const token = await this.get_bearer_token(req);
			if(token != process.env.ACTBOT_BEARER){
        return res.status(403).send('unauthorized access');
			}

			let graphapi = new Graph();
			let client = graphapi.initializeGraphForUserAuth();
			let files = req.body.files;
			let moved = [];
			const driveItem = {
					parentReference: {
							id: req.body.od_staged_id
					}
			};
			
			try {
				for (const item of files) {

					driveItem.name = item.name;

					const response = await client
						.api(`/users/${req.body.od_user_id}/drive/items/${item.id}`)
						.update(driveItem);
						
					moved.push(response);
				}
			} catch(error) {
				console.error('Error: ', error.message);
				return res.status(500).json({ error: error.message });
			}

			return res.status(200).json({'moved': moved});
		});

		this.server.app.post("/onedrive/excel/append", async (req, res) => {

			const token = await this.get_bearer_token(req);
			if(token != process.env.ACTBOT_BEARER){
				return res.status(403).send('unauthorized access');
			}

			const od_user_id = req.body.od_user_id;
			const od_item_id = req.body.od_item_id;

			const { post_id, post_title, permalink, post_date } = req.body.published[0];

			let graphapi = new Graph();
			let client = graphapi.initializeGraphForUserAuth();

			const baseUrl = `/users/${od_user_id}/drive/items/${od_item_id}/workbook`;

			try {
				const result = await new Promise((resolve, reject) => {
					this._excelLock = this._excelLock
						.catch(() => {})
						.then(async () => {

							const usedRange = await client
								.api(`${baseUrl}/worksheets('Sheet1')/usedRange`)
								.get();

							const nextRow = usedRange.rowCount + 1;
							const range = `A${nextRow}:D${nextRow}`;

							const response = await client
								.api(`${baseUrl}/worksheets('Sheet1')/range(address='${range}')`)
								.patch({
									values: [[post_id, post_title, permalink, post_date]]
								});

							return { range, values: response.values };
						})
						.then(resolve)
						.catch(reject);
				});

				return res.status(200).json({ success: true, ...result });
			} catch(error) {
				console.error('Excel update error: ', error);
				return res.status(500).json({ error: error.message });
			}
		});

		this.server.app.post("/onedrive/excel/append/images", async (req, res) => {

			const token = await this.get_bearer_token(req);
			if(token != process.env.ACTBOT_BEARER){
				return res.status(403).send('unauthorized access');
			}

			const od_user_id = req.body.od_user_id;
			const od_item_id = req.body.od_item_id;
			const post_id 	 = req.body.post_id;
			const post_title = req.body.post_title;
			const permalink  = req.body.permalink;
			const uploaded	 = req.body.uploaded;

			let graphapi = new Graph();
			let client = graphapi.initializeGraphForUserAuth();

			// const itemId = '01MZNEP73VW6BI2XJIRVAIWIYKYBAZZYFY';
			const baseUrl = `/users/${od_user_id}/drive/items/${od_item_id}/workbook`;

			try {

				const result = await Promise.all(uploaded.map(async (upload) => {
					this._excelLock = this._excelLock
						.catch(() => {})
						.then(async () => {
							const usedRange = await client
								.api(`${baseUrl}/worksheets('Sheet2')/usedRange`)
								.get();

							const nextRow = usedRange.rowCount + 1;
							const range = `A${nextRow}:D${nextRow}`;

							const response = await client
								.api(`${baseUrl}/worksheets('Sheet2')/range(address='${range}')`)
								.patch({
									values: [[post_id, post_title, permalink, upload.url]]
								});

							return { range, values: response.values };
						});
				}));

				return res.status(200).json({ success: true, ...result });
			} catch(error) {
				console.error('Excel update error: ', error);
				return res.status(500).json({ error: error.message });
			}
		});

		this.server.app.post("/ava/render/asset-library", async (req, res) => {
			
			const token = await this.get_bearer_token(req);
			if(token != process.env.ACTBOT_BEARER){
        return res.status(403).send('unauthorized access');
			}

			const od_user_id = req.body.od_user_id;
			const od_folder_id  = req.body.od_folder_id;

			let folders = await this.fetch_onedrive(od_user_id, od_folder_id);

			const folder_categories = this.buildCategoryMap(folders);

			const folder_slugs = Object.values(folder_categories);

			const folder_names = Object.keys(folder_categories);

			const asset_library_html =
				await Promise.all(folders.map(async (folder, folderIndex) => {

					const category_list_items = await this.fetchAllFiles(od_user_id, folder.id);

					const category_list_html = await Promise.all(category_list_items.map(async (asset) => {
						asset.label = folder_names[folderIndex];
						return this.buildHubListItem(asset);
					}));

					return `<div id="${folder_slugs[folderIndex]}" class="hubList-category">
						${category_list_html}
					</div>`;
				}));
			
			// const template = fs.readFileSync('./templates/asset-library.html', 'utf8');
			// const rendered = this.renderTemplate(template, assets);
			// res.status(200).send(rendered);
			return res.status(200).send(asset_library_html);
		});

		this.server.app.post("/ava/content/focus", async (req, res) => {

			try {

				const content = req.body.content;
				const system_prompt = fs.readFileSync('./system_prompts/content-focus.txt', 'utf8');
				
				// send this to our model to focus the content
				let response = await this.model.bedrock.invoke([
					new SystemMessage(system_prompt),
					new HumanMessage(`Enrich this content:\n\n${content}Only output the enriched copy with no extra commentary.`)
				]);

				return res.status(200).json({ output: response.content.trim() });

			} catch(err){
				return res.status(500).json({'error': err});
			}

		});

		this.server.app.post("/ava/content/seo", async (req, res) => {
			
			try {
				const content = req.body.content;
				const keywords = req.body.focus_keywords.trim();
				const system_prompt = fs.readFileSync('./system_prompts/content-seo.txt', 'utf8');

				const response = await this.model.bedrock.invoke([
					new SystemMessage(system_prompt),
					new HumanMessage(`Output an SEO description using these focus keywords:\n\n${keywords}Here is the given content:\n\n${content}`)
				]);
				console.log(`ANTHROPIC RESPONSE: ${response.content.trim()}`);
				return res.status(200).json({ output: response.content.trim() });
			
			} catch(err){
				return res.status(500).json({'error' : err});
			}
		});

		this.server.app.post("/ava/image/alt", async (req, res) => {

			const { images, content } = req.body;

			if( !images.length )
				return res.status(200).json();

			try {

				// Define structured output schema
				const AltTextSchema = z.object({
					results: z.array(z.object({
						id: z.number().describe("The image ID"),
						alt_text: z.string().describe("Concise, descriptive alt text for the image"),
					}))
				});

				const modelWithStructure = this.model.openai.withStructuredOutput(AltTextSchema);

				// Build message content
				const messageContent = [
					{
						type: "text",
						text: `Generate concise, descriptive alt text for each of the following images that is relevant to the following content\n\n${content}\n\n. For each image, return its ID and the corresponding alt text.`,
					}
				];

				// Create an agent that ignores the unauthorized certificate
				const agent = new https.Agent({ rejectUnauthorized: false });

				var i = 0;
				for( i; i < images.length; i++ ){
					let image = images[i];

					image.url = image.url.replace('https://act-main.ddev.site', 'http://ddev-act-main-web');

					// Pass the agent into the fetch options
					const response = await fetch(image.url, { agent: agent });
					const arrayBuffer = await response.arrayBuffer();
					const base64Image = Buffer.from(arrayBuffer).toString('base64');
					const mimeType = response.headers.get('content-type') || 'image/jpeg';

					messageContent.push(
						// Label each image so the model can map it back to an ID
						{ type: "text", 
							text: `Image ID: ${image.id} (${image.filename})` 
						},
						{
							type: "image_url",
							image_url: {
								url: `data:${mimeType};base64,${base64Image}`,
							},
						}
					);
				}

				const response = await modelWithStructure.invoke([new HumanMessage({ content: messageContent })]);

				// response is already parsed — { results: [{ id, alt_text }, ...] }
				return res.status(200).json({ output: response.results });

			} catch(err) {
				return res.status(500).json({ error: err.message });
			}

		});

		// Competitor Analysis + Brave Search API
		// --------------------------------------------

		this.server.app.post("/research/what-changed/brave", async (req, res) => {

			const token = await this.get_bearer_token(req);
			if(token != process.env.ACTBOT_BEARER){
        return res.status(403).send('unauthorized access');
			}

			try {

				let crm = req.body.competitor;

				if( crm.toLowerCase() == 'insightly' )
					crm = 'insightly crm';

				const params = new URLSearchParams({
					q: `${crm} pricing OR product update OR announcement intitle: ${crm}`,
					freshness: 'pw',
					extra_snippets: true,
					text_decorations: false,
					count: 5,
					operators: true,
					result_filter: 'web,discussions,news',
					goggles: `! name: Blacklisted sites
										! description: Sites to ignore
										! public: true
										! author: Jacob Ross
										! avatar: #01e837

										! Boost official website community sites
										$boost=4,site=salesforce.com
										$boost=4,site=hubspot.com
										$boost=4,site=zoho.com
										$boost=4,site=monday.com
										$boost=4,site=insightly.com
										$boost=4,site=pipedrive.com
										$boost=4,site=keap.com

										! Boost Reddit generally, then boost specific subs harder
										$boost=2,site=reddit.com
										/r/hubspot/$boost=4
										/r/Zoho/$boost=4
										/r/mondaydotcom/$boost=4
										/r/InsightlyCRM/$boost=4
										/r/keap/$boost=4
										/r/pipedrive/$boost=4
										/r/CRM/$boost=3
										/r/CRMSoftware/$boost=3
										/r/sales/$boost=3
										/r/smallbusiness/$boost=3
										/r/salesforce/$boost=3

										! Boost Hacker News (often has candid competitor takes on API/dev experience)
										$boost=2,site=news.ycombinator.com

										! Discard bias sites
										$discard,site=g2.com
										$discard,site=capterra.com
										$discard,site=getapp.com
										$discard,site=softwareadvice.com
										$discard,site=trustradius.com
										$discard,site=facebook.com
										$discard,site=x.com
										$discard,site=instagram.com
										$discard,site=wikipedia.org`
				});

				const data = await fetch(`https://api.search.brave.com/res/v1/web/search?${params}`, {
					method: 'get',
					headers: {
						'X-Subscription-Token': process.env.BRAVE_API_KEY,
						'Accept-Encoding': 'gzip',
					}
				}).then(r => r.json());

				let research = [
					...(data.web?.results ?? []),
					...(data.discussions?.results ?? []),
					...(data.news?.results ?? []),
				];
				
				res.status(200).json({ research });

			} catch (err) {
				console.error('Research fetch error:', err);
				res.status(500).json({ error: err.message });
			}
		});

		this.server.app.post("/research/what-changed/row", async (req, res) => {
			
			const token = await this.get_bearer_token(req);
			if(token != process.env.ACTBOT_BEARER){
        return res.status(403).send('unauthorized access');
			}

			try {
				let crm 		= req.body.crm;
				let research = req.body.research;
				let result 	= req.body.research[0];
				let color 	= '#FF7A59';

				switch(crm){
					case 'HubSpot':
						color = '#FF7A59';
						break;
					case 'Zoho':
						color = '#E42527';
						break;
					case 'Salesforce':
						color = '#00A1E0';
						break;
					case 'Keap':
						color = '#36A635';
						break;
					case 'Pipedrive':
						color = '#1A8000';
						break;
					case 'Monday.com':
						color = '#FF3D57';
						break;
					case 'Insightly CRM':
						break;
				}

				const timestamp = result.page_age;
				const date = new Date(timestamp).toLocaleDateString('en-US', {
					month: 'long',
					day: 'numeric'
				});

				const tfile = fs.readFileSync('./templates/competitor-row.hbs', 'utf8');
				const template = Handlebars.compile(tfile);
				let html = template({ crm, result, color, date });

				html = html.replace(`\n`, '');

				return res.status(200).json({ html, research });

			} catch(err){
				console.log(err);
				return res.status(500).json({error: err});
			}
			
		});

		this.server.app.post("/research/what-changed/html", async (req, res) => {
			
			const token = await this.get_bearer_token(req);
			if(token != process.env.ACTBOT_BEARER){
        return res.status(403).send('unauthorized access');
			}
			
			let rows = req.body.rows;

			let today = new Date().toLocaleDateString('en-US', {
					month: 'long',
					day: 'numeric',
					year: 'numeric'
			});

			let next_week = new Date();
					next_week.setDate(next_week.getDate() + 7);
					next_week = next_week.toLocaleDateString('en-US', {
						month: 'long',
						day: 'numeric'
					});

			try {
				const source = fs.readFileSync('./templates/competitor-full.hbs', 'utf8');
				const template = Handlebars.compile(source);
				const html = template({ rows, today, next_week });

				return res.status(200).json({ html });
			} catch(err) {
				console.error('Formatting error: ', err);
				return res.status(500).json({ error: err.message });
			}

		});

		this.server.app.post("/research/what-changed/update", async (req, res) => {

			const token = await this.get_bearer_token(req);
			if(token != process.env.ACTBOT_BEARER){
        return res.status(403).send('unauthorized access');
			}

			try {
				
				let updated_briefing = req.body.updated_briefing;

				let domain = `https://thepoint.act.com`;

				if( process.env.NODE_ENV == 'localhost' ){
					domain = `http://ddev-actpoint-web`;
				}

				// Get the outdated version of the briefing
				let response = await fetch(`${domain}/actrest/hub/competitor_briefing`, {
					method: 'get',
					headers: {
						'Accept':'application/json',
						'Authorization': `Bearer ${process.env.ACT_REST_TOKEN}`
					}
				});

				if( !response.ok )
					return res.status(500).json({'err': 'Unable to get competitor briefing from the point'});

				let outdated_briefing = await response.json();

				// Manipulate the HTML with Cheerio
				var $ = cheerio.load( outdated_briefing.html );
				
				// replace the old card HTML with the updated briefing HTML
				$('#sec-exec #what-changed.card').html(updated_briefing);

				// Send the entire #sec-exec container with updated HTML to update endpoing
				let full_html = $('#sec-exec').parent().html();

				response = await fetch(`${domain}/actrest/hub/competitor_briefing`, {
					method: 'post',
					headers: {
						'Content-Type': 'application/json',
						'Authorization': `Bearer ${process.env.ACT_REST_TOKEN}`
					},
					body: JSON.stringify({ html: full_html }),
				});

				if( !response.ok )
					return res.status(500).json({'err': 'Unable to updated competitor briefing on the point'});


				let update_response = await response.json();

				if(update_response.data.status == 200){
					return res.status(200).json({status: 'success', updated: update_response.data});
				}

				return res.status(500).json({error: update_response.message});

			} catch(err){
				return res.status(500).json({error: err})
			}


		});

		// Competitor Analysis + Ava Knowledge
		this.server.app.post("/research/talking-points/html", async(req, res) => {

			const token = await this.get_bearer_token(req);
			if(token != process.env.ACTBOT_BEARER){
        return res.status(403).send('unauthorized access');
			}

			let crm = req.body.crm.toLowerCase();
			let talking_points = req.body.talking_points;
			try {

				let html = `<div class="bc-fullwidth" style=" display: grid; border: 1px solid #e0e4f0; border-top: none; "> 
					<div id="${crm.toLowerCase()}" class="bc-cell talking-points"> 
					<div class="bc-cell-title"><img draggable="false" role="img" class="emoji" alt="💬" src="https://s.w.org/images/core/emoji/17.0.2/svg/1f4ac.svg"> SALES STRATEGY</div>
					${talking_points} </div> </div>`;

				return res.status(200).json({crm, html});
			} catch(err){
				return res.status(500).json({'error' : err});
			}
		});

		this.server.app.post("/research/talking-points/update", async (req, res) => {

			const token = await this.get_bearer_token(req);
			if(token != process.env.ACTBOT_BEARER){
        return res.status(403).send('unauthorized access');
			}

			try {
				
				let crm = req.body.crm;
				let talking_points = req.body.talking_points;

				let domain = `https://thepoint.act.com`;

				if( process.env.NODE_ENV == 'localhost' ){
					domain = `http://ddev-actpoint-web`;
				}

				// Get the outdated version of the briefing
				let response = await fetch(`${domain}/actrest/hub/battlecards`, {
					method: 'get',
					headers: {
						'Accept':'application/json',
						'Authorization': `Bearer ${process.env.ACT_REST_TOKEN}`
					}
				});

				if( !response.ok )
					return res.status(500).json({'err': 'Unable to get competitor briefing from the point'});

				let battlecards = await response.json();

				// Manipulate the HTML with Cheerio
				var $ = cheerio.load( battlecards.html );

				// replace the old card HTML with the updated briefing HTML
				$(`#sec-battlecards .talking-points`).remove();

				for( var x = 0; x<crm.length; x++ ){
					$(`#sec-battlecards #bc-${crm[x]} .bc-fullwidth`).remove();
					$(`#sec-battlecards #bc-${crm[x]}`).append(talking_points[x]);
				}
				// Send the entire #sec-exec container with updated HTML to update endpoing
				let full_html = $('#sec-battlecards').parent().html();
				
				response = await fetch(`${domain}/actrest/hub/battlecards`, {
					method: 'post',
					headers: {
						'Content-Type': 'application/json',
						'Authorization': `Bearer ${process.env.ACT_REST_TOKEN}`
					},
					body: JSON.stringify({ html: full_html }),
				});

				if( !response.ok )
					return res.status(500).json({'err': 'Unable to updated battlecards on the point'});

				let update_response = await response.json();

				if(update_response.data.status == 200){
					return res.status(200).json({status: 'success', updated: update_response.data});
				}

				return res.status(500).json({error: update_response.message});

			} catch(err){
				return res.status(500).json({error: err})
			}

		});

		// Priorities Section
		// --------------------------------------------

		this.server.app.post("/research/priorities/ai", async (req, res) => {

			const token = await this.get_bearer_token(req);
			if(token != process.env.ACTBOT_BEARER){
        return res.status(403).send('unauthorized access');
			}

			let research = req.body.research;

			if( !research )
				return res.status(400).json({ error: 'Missing research' });

			let context = research.map( r => {
				return r.description + `\n\n` + (r.extra_snippets ? r.extra_snippets.join(`\n`) : '')
			});

			try {

				// use Ai to come up with 4 key action items
				const PrioritiesSchema = z.object({
					sales_rep: z.string().describe("What should sales reps be focusing on?"),
					product_marketing: z.string().describe("How should we be marketing our product?"),
					customer_marketing: z.string().describe("Where should we be looking for more customers?"),
					leadership: z.string().describe("How should leadership adapt their approach to meet the shifting market?"),
				});

				const modelWithStructure = this.model.openai.withStructuredOutput(PrioritiesSchema);

				const response = await modelWithStructure.invoke([
					new SystemMessage(`You are a marketing expert consulting Act! and your task is to advise Act! on competitive strategy to help Act! stand-out from the other competitors including Salesforce, HubSpot, Keap, Insightly CRM, Zoho, Pipedrive, Monday.com.`),
					new HumanMessage(`Based on the following research results, provide one concise, actionable priority for each of the four audiences below.\n\nResearch results:\n\n${context}`)
				]);

				// response is already parsed — { sales_rep, product_marketing, customer_marketing, leadership }
				return res.status(200).json({ response });

			} catch(err) {
				console.log(err);
				return res.status(500).json({ error: err.message });
			}
		});

		this.server.app.post("/research/priorities/html", async (req, res) => {

			const token = await this.get_bearer_token(req);
			if(token != process.env.ACTBOT_BEARER){
        return res.status(403).send('unauthorized access');
			}

			let response = req.body.response;

			try {

				const source = fs.readFileSync('./templates/priorities.hbs', 'utf8');
				const template = Handlebars.compile(source);
				const html = template({ response });
				return res.status(200).json({ html });

			} catch(err) {

				console.error('Formatting error: ', err);
				return res.status(500).json({ error: err.message });
			}
		});

		this.server.app.post("/research/priorities/update", async (req, res) => {

			const token = await this.get_bearer_token(req);
			if(token != process.env.ACTBOT_BEARER){
        return res.status(403).send('unauthorized access');
			}

			try {

				let updated_priorities = req.body.updated_priorities;

				let domain = `https://thepoint.act.com`;

				if( process.env.NODE_ENV == 'localhost' ){
					domain = `http://ddev-actpoint-web`;
				}

				// Get the outdated version of the briefing
				let response = await fetch(`${domain}/actrest/hub/competitor_briefing`, {
					method: 'get',
					headers: {
						'Accept': 'application/json',
						'Authorization': `Bearer ${process.env.ACT_REST_TOKEN}`
					}
				});

				let outdated_priorities = await response.json();

				// Manipulate the HTML with Cheerio
				var $ = cheerio.load( outdated_priorities.html );
				
				// replace the old card HTML with the updated priorities HTML
				$('#sec-exec #priorities').html(updated_priorities);

				// Send the entire #sec-exec container with updated HTML to update endpoing
				let full_html = $('#sec-exec').parent().html();

				response = await fetch(`${domain}/actrest/hub/competitor_briefing`, {
					method: 'post',
					headers: {
						'Content-Type': 'application/json',
						'Authorization': `Bearer ${process.env.ACT_REST_TOKEN}`
					},
					body: JSON.stringify({ html: full_html }),
				});

				let update_response = await response.json();

				if(update_response.data.status == '200'){
					return res.status(200).json({status: 'success', updated: update_response.data});
				}

				return res.status(500).json({error: update_response.message});

			} catch(err){
				return res.status(500).json({error: err})
			}

		});

		// Watching Section
		this.server.app.post("/research/watching/ai", async (req, res) => {

			const token = await this.get_bearer_token(req);
			if(token != process.env.ACTBOT_BEARER){
        return res.status(403).send('unauthorized access');
			}

			let research = req.body.research;

			if( !research )
				return res.status(400).json({ error: 'Missing research' });

			let context = research.map( r => {
				return r.description + `\n\n` + (r.extra_snippets ? r.extra_snippets.join(`\n`) : '');
			});

			try {

				// use Ai to come up with 4 key action items
				const PrioritiesSchema = z.object({
					hubspot: z.string().describe("What the latest with HubSpot?"),
					zoho: z.string().describe("What the latest with Zoho?"),
					keap: z.string().describe("What the latest with Keap?"),
					pipedrive: z.string().describe("What the latest with Pipedrive?"),
					monday: z.string().describe("What the latest with Monday.com?"),
					insightly: z.string().describe("What the latest with Insightly CRM?"),
				});

				const modelWithStructure = this.model.openai.withStructuredOutput(PrioritiesSchema);

				const response = await modelWithStructure.invoke([
					new SystemMessage(`You are a SaaS expert who specializes in keeping up with the latest news and announcements with CRM software. Your task is to provide one concise statement describing what's new with each CRM including Salesforce, HubSpot, Keap, Insightly CRM, Zoho, Pipedrive, Monday.com.`),
					new HumanMessage(`Based on the following research results, provide one concise sentence describing what's new with each CRM.\n\nResearch results:\n\n${context}`)
				]);

				// response is already parsed — { sales_rep, product_marketing, customer_marketing, leadership }
				return res.status(200).json({ response });

			} catch(err) {
				console.log(err);
				return res.status(500).json({ error: err.message });
			}
		});

		this.server.app.post("/research/watching/html", async (req, res) => {
			
			const token = await this.get_bearer_token(req);
			if(token != process.env.ACTBOT_BEARER){
        return res.status(403).send('unauthorized access');
			}

			let response = req.body.response;

			try {
				let vendors = []
				
				for (const key in response) {
					if (response.hasOwnProperty(key)) {
						vendors.push({vendor: this.ucFirst(key), update: response[key]});
					}
				}

				const source = fs.readFileSync('./templates/watching.hbs', 'utf8');
				const template = Handlebars.compile(source);
				const html = template({ vendors });
				return res.status(200).json({ html });

			} catch(err) {

				console.error('Formatting error: ', err);
				return res.status(500).json({ error: err.message });
			}
		});

		this.server.app.post("/research/watching/update", async (req, res) => {
			
			const token = await this.get_bearer_token(req);
			if(token != process.env.ACTBOT_BEARER){
        return res.status(403).send('unauthorized access');
			}

			try {

				let updated_watching = req.body.updated_watching;

				let domain = `https://thepoint.act.com`;

				if( process.env.NODE_ENV == 'localhost' ){
					domain = `http://ddev-actpoint-web`;
				}

				// Get the outdated version of the briefing
				let response = await fetch(`${domain}/actrest/hub/competitor_briefing`, {
					method: 'get',
					headers: {
						'Accept': 'application/json',
						'Authorization': `Bearer ${process.env.ACT_REST_TOKEN}`
					}
				});

				let outdated_watching = await response.json();

				// Manipulate the HTML with Cheerio
				var $ = cheerio.load( outdated_watching.html );
				
				// replace the old card HTML with the updated watching HTML
				$('#sec-exec #watching').html(updated_watching);

				// Send the entire #sec-exec container with updated HTML to update endpoing
				let full_html = $('#sec-exec').parent().html();

				response = await fetch(`${domain}/actrest/hub/competitor_briefing`, {
					method: 'post',
					headers: {
						'Accept': 'application/json',
						'Authorization': `Bearer ${process.env.ACT_REST_TOKEN}`
					},
					body: JSON.stringify({ html: full_html }),
				});

				let update_response = await response.json();

				if(update_response.data.status == '200'){
					return res.status(200).json({status: 'success', updated: update_response.data});
				}

				return res.status(500).json({error: update_response.message});

			} catch(err){
				return res.status(500).json({error: err})
			}
		});

		// Coming Up
		this.server.app.post("/research/upcoming/brave", async (req, res) => {
			
			const token = await this.get_bearer_token(req);
			if(token != process.env.ACTBOT_BEARER){
        return res.status(403).send('unauthorized access');
			}

			try {
				const params = new URLSearchParams({
					q: `${req.body.competitor} anticipated updates OR anticipated releases OR anticipated price changes intitle: ${req.body.competitor}`,
					freshness: 'pw',
					extra_snippets: true,
					text_decorations: false,
					count: 5,
					operators: true,
					result_filter: 'web,discussions,news',
					goggles: `! name: Blacklisted sites
										! description: Sites to ignore
										! public: true
										! author: Jacob Ross
										! avatar: #01e837

										! Boost official website community sites
										$boost=4,site=salesforce.com
										$boost=4,site=hubspot.com
										$boost=4,site=zoho.com
										$boost=4,site=monday.com
										$boost=4,site=insightly.com
										$boost=4,site=pipedrive.com
										$boost=4,site=keap.com

										! Boost Reddit generally, then boost specific subs harder
										$boost=2,site=reddit.com
										/r/hubspot/$boost=4
										/r/Zoho/$boost=4
										/r/mondaydotcom/$boost=4
										/r/InsightlyCRM/$boost=4
										/r/keap/$boost=4
										/r/pipedrive/$boost=4
										/r/CRM/$boost=3
										/r/CRMSoftware/$boost=3
										/r/sales/$boost=3
										/r/smallbusiness/$boost=3
										/r/salesforce/$boost=3

										! Boost Hacker News (often has candid competitor takes on API/dev experience)
										$boost=2,site=news.ycombinator.com

										! Discard bias sites
										$discard,site=g2.com
										$discard,site=capterra.com
										$discard,site=getapp.com
										$discard,site=softwareadvice.com
										$discard,site=trustradius.com
										$discard,site=facebook.com
										$discard,site=x.com
										$discard,site=instagram.com
										$discard,site=wikipedia.org`
				});

				const data = await fetch(`https://api.search.brave.com/res/v1/web/search?${params}`, {
					method: 'get',
					headers: {
						'X-Subscription-Token': process.env.BRAVE_API_KEY,
						'Accept-Encoding': 'gzip',
					}
				}).then(r => r.json());

				let research = [
					...(data.web?.results ?? []),
					...(data.discussions?.results ?? []),
					...(data.news?.results ?? []),
				];
				
				res.status(200).json({ research });

			} catch (err) {
				console.error('Research fetch error:', err);
				res.status(500).json({ error: err.message });
			}
		});

		this.server.app.post("/research/upcoming/row", async (req, res) => {
			
			const token = await this.get_bearer_token(req);
			if(token != process.env.ACTBOT_BEARER){
        return res.status(403).send('unauthorized access');
			}
			
			try {
				let crm 		 = req.body.crm;
				let research = req.body.research;
				let result 	 = req.body.research[0];

				const tfile = fs.readFileSync('./templates/upcoming.hbs', 'utf8');
				const template = Handlebars.compile(tfile);
				let html = template({ result });

				html = html.replace(`\n`, '');

				return res.status(200).json({ html });

			} catch(err){
				console.log(err);
				return res.status(500).json({error: err});
			}
		});

		this.server.app.post("/research/upcoming/update", async (req, res) => {

			const token = await this.get_bearer_token(req);
			if(token != process.env.ACTBOT_BEARER){
        return res.status(403).send('unauthorized access');
			}

			try {

				let updated_upcoming = req.body.updated_upcoming;

				let domain = `https://thepoint.act.com`;

				if( process.env.NODE_ENV == 'localhost' ){
					domain = `http://ddev-actpoint-web`;
				}

				// Get the outdated version of the briefing
				let response = await fetch(`${domain}/actrest/hub/competitor_briefing`, {
					method: 'get',
					headers: {
						'Accept': 'application/json',
						'Authorization': `Bearer ${process.env.ACT_REST_TOKEN}`
					}
				});

				let outdated_upcoming = await response.json();

				// Manipulate the HTML with Cheerio
				var $ = cheerio.load( outdated_upcoming.html );
				
				// replace the old card HTML with the updated upcoming HTML
				$('#sec-exec #upcoming').html(updated_upcoming);

				// Send the entire #sec-exec container with updated HTML to update endpoing
				let full_html = $('#sec-exec').parent().html();

				response = await fetch(`${domain}/actrest/hub/competitor_briefing`, {
					method: 'post',
					headers: {
						'Accept': 'application/json',
						'Authorization': `Bearer ${process.env.ACT_REST_TOKEN}`
					},
					body: JSON.stringify({ html: full_html }),
				});

				let update_response = await response.json();

				if(update_response.data.status == '200'){
					return res.status(200).json({status: 'success', updated: update_response.data});
				}

				return res.status(500).json({error: update_response.message});

			} catch(err){
				return res.status(500).json({error: err})
			}
		});

		this.server.app.post("/bedrock/invoke/weekly-brief", async (req, res) => {
			
			try {
				const context = req.body.context;
				const system_prompt = fs.readFileSync('./system_prompts/weekly-brief.txt', 'utf8');
				const HTML = fs.readFileSync('./templates/weekly-briefing.html', 'utf8');

				const response = await this.model.bedrock.invoke([
					new SystemMessage(system_prompt),
					new HumanMessage(
						`Here is the HTML document:\n\n${HTML}\n\nReplace the content inside div#what-changed with the 5 most impactful updates from this research context:\n\n${JSON.stringify(context, null, 2)}`
					),
				]);

				const content = typeof response.content === 'string'
					? response.content.trim()
					: response.content;

				res.status(200).json(content);

			} catch (err) {
				console.error('Weekly brief error:', err);
				res.status(500).json({ error: err.message });
			}
		});
	}

	async fetchAllFiles(od_user_id, folder_id) {
		const items = await this.fetch_onedrive(od_user_id, folder_id);
		const results = [];
		await Promise.all(items.map(async (item) => {
			if ('folder' in item) {
				const subFiles = await this.fetchAllFiles(od_user_id, item.id);
				results.push(...subFiles);
			} else {
				results.push(item);
			}
		}));
		return results;
	}

	async fetch_onedrive(od_user_id, od_folder){
		let items = [];
		try {
			const graphapi = new Graph();
			const client = graphapi.initializeGraphForUserAuth();
			let response = await client.api(`/users/${od_user_id}/drive/items/${od_folder}/children`).get();
			items = response.value;
		} catch(error) {
			console.error('Error: ', error);
			return res.status(500).json({ error: error.message });
		}
		return items;
	}

	send_to_email(){

	}

	async get_bearer_token(req){
		if(typeof req.headers.authorization === 'undefined'
			|| !req.headers.authorization.includes('Bearer') )
				return false;
		return req.headers.authorization.replace('Bearer ', '');
	}

	check_session( req, res, next ){
		if( typeof req.session === 'undefined' || typeof req.session.authenticated === 'undefined' ){
			return res.redirect('/login');
		} else next();
	}

	slugify(name) {
		return name
			.toLowerCase()
			.trim()
			.replace(/[^a-z0-9\s-]/g, '')  // strip special chars
			.replace(/\s+/g, '-');          // spaces to hyphens
	}

	buildCategoryMap(folders) {
		return folders.reduce((accumulator, folder, folderIndex, folderArray) => {
			accumulator[folder.name] = this.slugify(folder.name);
			return accumulator;
		}, {});
	}
	
	buildHubList(folders, assets) {
		// Group assets by folder
		const grouped = assets.reduce((acc, asset) => {
			const key = folders[asset.folder] ?? 'sops';
			if (!acc[key]) acc[key] = { folder: asset.folder, items: [] };
			acc[key].items.push(asset);
			return acc;
		}, {});

		return Object.entries(grouped).map(([categoryId, { folder, items }]) => `
			<div id="${categoryId}" class="hubList-category">
				${items.map(buildHubListItem).join('\n')}
			</div>`
		).join('\n');
	}
	
	buildHubListItem(asset) {

		const isNew = this.isWithinDays(asset.lastModifiedDateTime, 30);
		const newPill = isNew ? `<span class="hubList-new">NEW</span>` : '';
		const dateLabel = this.formatDate(asset.lastModifiedDateTime);
		const categoryLabel = asset.label;
		if( 'file' in asset ){
			asset.ext = path.extname(asset.name).slice(1);
		}
		// console.log(asset);
		const hasLink = !!asset.webUrl;

		const link = hasLink
			? `<a href="${asset.webUrl}" class="hubList-link" target="_blank" rel="noopener">↗ Open</a>`
			: `<a href="#" class="hubList-link hubList-link--pending" onclick="return false;" title="OneDrive link not yet added">🔗 Link pending</a>`;

		return `
			<div class="hubList-item" data-title="${this.escapeAttr(asset.name)}" data-date="${asset.lastModifiedDateTime.slice(0,10)}" data-type="${asset.ext}">
				<div class="hubList-info">
					<div class="hubList-title">${newPill}${this.escapeHtml(asset.name)}</div>
					<div class="hubList-meta">${this.escapeHtml(categoryLabel)} · ${dateLabel}</div>
				</div>
				<span class="hubList-badge hubList-badge--${asset.ext}">${asset.ext}</span>
				${link}
			</div>`;
	}
	
	renderTemplate(html, assets) {
		const hubListHTML = this.buildHubList(assets);

		// Replace the entire hubList div contents
		html = html.replace(
			/(<div id="hubList">)[\s\S]*?(<\/div>\s*<div class="hub-empty")/, 
			`$1\n${hubListHTML}\n$2`
		);

		// Update the count label
		html = html.replace(
			/(<span class="hub-count-label" id="hubCountLabel">)[^<]*/,
			`$1${assets.length} assets`
		);

		return html;
	}

	isWithinDays(dateStr, days) {
		return (Date.now() - new Date(dateStr).getTime()) < days * 86400000;
	}

	formatDate(dateStr) {
		return new Date(dateStr).toLocaleDateString('en-US', { 
			month: 'short', day: 'numeric', year: 'numeric' 
		});
	}

	escapeHtml(str) {
		return str.replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;');
	}

	escapeAttr(str) {
		return str.replace(/&/g,'&amp;').replace(/"/g,'&quot;');
	}
	ucFirst(str) {
		if (!str) return str;
		return str.charAt(0).toUpperCase() + str.slice(1);
	}
}
