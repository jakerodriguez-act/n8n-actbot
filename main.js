import dotenv from 'dotenv';
dotenv.config();
import Routes from './core/routes.js';
import Server from './core/server.js';
import LLM from './core/llm.js';


// --------------------------------------------
// SERVER
// --------------------------------------------

const _server = new Server(import.meta.dirname);
// process.on('SIGTERM', _server.shutdown);
// process.on('SIGINT', _server.shutdown);

// --------------------------------------------
// MODEL
// --------------------------------------------

const _llms = new LLM();

// --------------------------------------------
// ROUTING
// --------------------------------------------

new Routes(_server, _llms);
