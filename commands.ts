import { UsernamePasswordCredential } from '@azure/identity';
import { TokenCredentialAuthenticationProvider, TokenCredentialAuthenticationProviderOptions } from "@microsoft/microsoft-graph-client/authProviders/azureTokenCredentials";
import { Client } from '@microsoft/microsoft-graph-client';
require('isomorphic-fetch')
import * as fs from 'fs';
import axios from 'axios';

// Get the command line arguments
const args: string[] = process.argv.slice(2);
const command: string = args[0];
const appId: string = args[1];
const tenantId: string = args[2];
const clientId: string = args[3];
const userName: string = args[4];
const password: string = args[5];


// Check if there are any arguments
if (args.length < 6) {
    console.log("No arguments provided. node common.js <command> <appId> <tenantId> <clientId> <userName> <password>");
} else {
    // Print the provided arguments
    console.log("Arguments provided");
}

// User to get access to App Catalog
// Requirments:
//      - User must be a Teams Service Administrator
//      - User must be a Teams Service Administrator for Publish
//      - User must be Global Administrator for Update
const credential = new UsernamePasswordCredential(
  tenantId,
  clientId,
  userName,
  password,
);

const authProvider = new TokenCredentialAuthenticationProvider(credential, {
    scopes: ['User.Read', 'AppCatalog.ReadWrite.All'],
});

async function getToken() : Promise<string> {
    const response = await credential.getToken(['User.Read', 'AppCatalog.ReadWrite.All']);
    return response.token;
}

const graphClient = Client.initWithMiddleware({ authProvider: authProvider });

async function getApps() {
    let teamsApps = await graphClient.api('/appCatalogs/teamsApps')
	.filter('distributionMethod eq \'organization\'')
	.get();
    console.log(teamsApps);
  }

async function PostData(data, url) {
    return new Promise(async (resolve) => {
        var config = {
            method: 'post',
            url: url,
            headers: {
                'Authorization': await getToken(),
                'Content-Type': 'application/zip'
            },
            data: data
        };

        axios(config);
    })
}


async function publishApp() {
    const teamsApp = fs.readFile('./package/appPackage.local.zip', async (err, data) => {
        if (err) throw err;
        await PostData(data, 'https://graph.microsoft.com/v1.0/appCatalogs/teamsApps?requiresReview=false');
        console.log('App published');
    });
}

async function updateApp(appId: string) {
    const teamsApp = fs.readFile('./package/appPackage.local.zip', async (err, data) => {
        if (err) throw err;
        await PostData(data, 'https://graph.microsoft.com/v1.0/appCatalogs/teamsApps/' + appId + '/appDefinitions');
        console.log('App updated');
    });
}


if(command === 'update') updateApp(appId);
if(command === 'publish') publishApp();
if(command === 'list') getApps();
