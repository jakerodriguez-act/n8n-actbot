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
import hbs from 'express-handlebars';

export default class Server {
  // --------------------------------------------
  // SERVER SETUP 
  // --------------------------------------------

  constructor(dirname){

    this.app = express();

    // For application/x-www-form-urlencoded
    this.app.use(express.urlencoded({ extended: true, limit: '10mb' }));
    this.app.use(bodyParser.json({limit: '10mb'})); // Parses JSON-formatted request bodies
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
    this.app.engine( 'hbs', hbs({ 
      extname: 'hbs',
      partialsDir: path.join(dirname, '/includes/views/partials/'),
      helpers: {
        PST_time: function (timestamp) {
          const ms = Number(timestamp);
          if (!Number.isFinite(ms)) return '';
          return DateTime
            .fromMillis(ms)
            .setZone('America/Los_Angeles')
            .toFormat("MMM d, yyyy\nh:mm:ss a");
        },
        nl2br: function(text){
          if (!text) return '';
          const escaped = Handlebars.escapeExpression(text);
          const withBreaks = escaped.replace(/\r?\n/g, '<br>');
          return new Handlebars.SafeString(withBreaks);
        },
        Uppercase: function(string) {
          if( string )
            return string.toUpperCase();
        },
        ifCond: function (v1, operator, v2, options) {
          const operators = {
            '==': (a, b) => a == b, '===': (a, b) => a === b,
            '!=': (a, b) => a != b, '!==': (a, b) => a !== b,
            '<': (a, b) => a < b, '<=': (a, b) => a <= b,
            '>': (a, b) => a > b, '>=': (a, b) => a >= b,
            '&&': (a, b) => a && b, '||': (a, b) => a || b
          };
          return operators[operator](v1, v2) ? options.fn(this) : options.inverse(this);
        }
      }
    }));

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
      
      this.server = createHttpServer(this.options, this.app);

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