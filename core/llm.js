import dotenv from 'dotenv';
dotenv.config();

// --------------
// OPEN AI
// --------------

import { 
  ChatOpenAI, 
  OpenAIEmbeddings 
} from "@langchain/openai";

// --------------
// AWS BEDROCK
// --------------

import { ChatBedrockConverse } from "@langchain/aws";

export default class LLM {

  constructor(){

    this.bedrock = {};
    this.openai = {};

    // OPEN AI
    // -----------------------------------------------------
    this.openai = new ChatOpenAI({
      apiKey: process.env.OPENAI_API_KEY,
      model: "gpt-5.4-mini",
      service_tier: "priority"
    });
    this.openai.input_cost = 0.75,
    this.openai.output_cost = 4.50,
    this.openai.per_sum = 1000000

    // BEDROCK
    // -----------------------------------------------------

    this.bedrock = new ChatBedrockConverse({
      region: process.env.BEDROCK_AWS_REGION ?? "us-east-1",
      model: "anthropic.claude-sonnet-4-6",
      applicationInferenceProfile:
      "arn:aws:bedrock:us-east-1:301745145890:inference-profile/global.anthropic.claude-sonnet-4-6",
      credentials: {
        secretAccessKey: process.env.BEDROCK_AWS_SECRET_ACCESS_KEY,
        accessKeyId: process.env.BEDROCK_AWS_ACCESS_KEY_ID,
      },
      service_tier: "priority"
    });
  }
}