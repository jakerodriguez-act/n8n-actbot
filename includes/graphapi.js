import 'isomorphic-fetch';
import * as azure from '@azure/identity';
import * as graph from '@microsoft/microsoft-graph-client';
import * as authProviders from '@microsoft/microsoft-graph-client/authProviders/azureTokenCredentials/index.js';

export default class Graph {

  constructor(){

    this.settings = {
      'clientId': process.env.GRAPH_CLIENT_ID,
      'tenantId': process.env.GRAPH_TENANT_ID,
      'clientSecret': process.env.GRAPH_SECRET_VALUE,
      'scopes': ['https://graph.microsoft.com/.default']
    };

  }

  initializeGraphForUserAuth() {
    try{

      const credential = new azure.ClientSecretCredential(
        this.settings.tenantId,
        this.settings.clientId,
        this.settings.clientSecret
      );

      const authProvider = new authProviders.TokenCredentialAuthenticationProvider(
        credential, {
          scopes: this.settings.scopes
        });

      return graph.Client.initWithMiddleware({
        authProvider: authProvider
      });

    } catch(e){
      console.log('Error:', e);
    }
  }

}

