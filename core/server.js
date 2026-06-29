import dotenv from 'dotenv';
dotenv.config();
import express from 'express';
import path from 'path';
import fs from 'fs';
import { createServer as createHttpServer } from 'http';
import { createServer as createHttpsServer } from 'https';
import cors from 'cors';
import bodyParser from 'body-parser';
import { DateTime } from 'luxon';
import session from 'express-session';

export default class Server {
  // --------------------------------------------
  // SERVER SETUP 
  // --------------------------------------------

  constructor(dirname){

    this.app = express();

    // For application/x-www-form-urlencoded
    this.app.use(express.urlencoded({ extended: true }));
    this.app.use(bodyParser.json()); // Parses JSON-formatted request bodies
    this.app.use(express.static(dirname + '/public/'));
    
    const session_options = {
      secret: process.env.SESS_SECRET,
      saveUninitialized: false,
      resave: true,
      cookie: {}
    };

    // secure cookie if staging or production
    if( process.env.NODE_ENV !== 'localhost' ){
      session_options.cookie.secure = true;
    }
    
    this.app.use(session(session_options));

    // Handlebars templating
    this.app.set('view engine', 'hbs');
    this.app.set(dirname + '/views/');

    this.port = process.env.PORT;

    this.options = { 
      cors: {
        origin: '*',
        methods: ["GET", "POST"]
      }
    };

    
    // --------------------------------------------
    // RUN NODE SERVER
    // --------------------------------------------
    if( process.env.NODE_ENV === 'localhost' ){

      this.server = createHttpServer(this.options, this.app);
      this.server.listen(process.env.PORT, () => {
        console.log(`server running on PORT: ${process.env.PORT}`);
        console.log(`NOTICE: You are running in LOCALHOST mode.`);
      });

    } else {
      
      if( process.env.NODE_ENV == 'development' ){
        this.options.key = fs.readFileSync('/etc/letsencrypt/archive/testbot.actagentai.com/privkey1.pem', 'utf8');
        this.options.cert = fs.readFileSync('/etc/letsencrypt/archive/testbot.actagentai.com/fullchain1.pem', 'utf8');
      } else
      if( process.env.NODE_ENV === 'production' ){
        this.options.key = fs.readFileSync('/etc/letsencrypt/archive/actbot.actagentai.com/privkey1.pem', 'utf8');
        this.options.cert = fs.readFileSync('/etc/letsencrypt/archive/actbot.actagentai.com/fullchain1.pem', 'utf8');
      }

      // this.options.cors.origin = process.env.CORS_ORIGIN.split(',');
      // this.app.use(cors(this.options.cors));
      this.server = createHttpsServer(this.options, this.app);

      this.server.listen(process.env.PORT, () => {
        console.log(`server running on PORT: ${process.env.PORT}`);
        console.log(`NOTICE: You are running in ${process.env.NODE_ENV.toUpperCase()} mode.`);
      });

    }

  }

  // shutdown(){
  //   console.log('Starting shutdown procedure');
  //   this.server.close(() =>{
  //     console.log('Express shutting down.');
  //   });
  // }
  
}