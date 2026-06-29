import dotenv from 'dotenv';
dotenv.config();
// import Mailer from '../includes/mailer.js';
import path from 'path';
import fs from 'fs';
// import { glob } from 'glob';
import _ from 'lodash';
import {z} from 'zod';
import https from 'https';
import Graph from '../includes/graphapi.js';

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

		this.server.app.post("/testnode", async(req, res) => {
			const token = await this.get_bearer_token(req);
			console.log('token length: ', token.length);
			if(token == process.env.ACTBOT_BEARER){

				console.log('testnode: token accepted');
				return res.send('success');
			} else {
        return res.status(403).send('unauthorized access');
      }
		});

		this.server.app.post("/onedrive/files/read", async(req, res) => {
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

		this.server.app.post("/onedrive/files/download", async(req, res) => {

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

		this.server.app.post("/onedrive/files/move", async(req, res) => {

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
					const response = await client
						.api(`/users/${process.env.ONEDRIVE_USER_ID}/drive/items/${item.id}`)
						.update(driveItem);
					moved.push(response);
				}
			} catch(error) {
				console.error('Error: ', error.message);
				return res.status(500).json({ error: error.message });
			}

			return res.status(200).json({'moved': moved});
		});

		this.server.app.post("/onedrive/excel/append", async(req, res) => {

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

		this.server.app.post("/onedrive/excel/append/images", async(req, res) => {

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

		this.server.app.post('/ava/render/asset-library', async (req, res) => {
			
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

		this.server.app.post('/ava/content/focus', async (req, res) => {

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

		this.server.app.post('/ava/content/seo', async (req, res) => {
			
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

		this.server.app.post('/ava/image/alt', async (req, res) => {

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

		this.server.app.post("/brave/research/competitor", async(req, res) => {

			const BLOCKLIST = ['g2.com', 'capterra.com', 'getapp.com', 'softwareadvice.com'];

			const queries = [
				'HubSpot pricing OR product update OR announcement June 2026',
				'Zoho CRM product update OR pricing June 2026',
				'Salesforce release OR Agentforce OR pricing change June 2026',
				'Pipedrive product update OR pricing OR news June 2026',
				'Monday.com CRM product update OR pricing June 2026',
			];

			try {
				const responses = await Promise.allSettled(
					queries.map(q =>
						fetch(`https://api.search.brave.com/res/v1/web/search?q=${encodeURIComponent(q)}&count=5`, {
							headers: { 'X-Subscription-Token': process.env.BRAVE_API_KEY }
						}).then(r => r.json())
					)
				);

				const flattened_results = responses.flatMap(response =>
					(response.web?.results ?? [])
						.filter(r => !BLOCKLIST.some(blocked => r.meta_url?.hostname?.includes(blocked)))
						.map(r => ({
							title: r.title,
							url: r.url,
							description: r.description,
							age: r.age,
							source: r.meta_url?.hostname,
						}))
				);

				res.status(200).json({ results: flattened_results });

			} catch (err) {
				console.error('Research fetch error:', err);
				res.status(500).json({ error: err.message });
			}
		});

		this.server.app.post("/bedrock/invoke/weekly-brief", async(req, res) => {
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
	
}