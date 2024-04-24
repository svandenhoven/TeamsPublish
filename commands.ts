import { UsernamePasswordCredential } from '@azure/identity';
import { TokenCredentialAuthenticationProvider, TokenCredentialAuthenticationProviderOptions } from "@microsoft/microsoft-graph-client/authProviders/azureTokenCredentials";
import { Client } from '@microsoft/microsoft-graph-client';
require('isomorphic-fetch')
import * as fs from 'fs';
import axios from 'axios';
import { BlobOptions } from 'buffer';

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
//      - The clientID need to have "Allow public client flows enabled" in the app registration
const credential = new UsernamePasswordCredential(
  tenantId,
  clientId,
  userName,
  password,
);

const authProvider = new TokenCredentialAuthenticationProvider(credential, {
    scopes: ['User.Read', 'AppCatalog.ReadWrite.All'],
});

async function getToken(print: boolean = false) : Promise<string> {
    const response = await credential.getToken(['User.Read', 'AppCatalog.ReadWrite.All', 'AppCatalog.Submit']);
    if (print) console.log(response.token);
    return response.token;

}

const graphClient = Client.initWithMiddleware({ authProvider: authProvider });

async function getAppId() {
    // let teamsApps = await graphClient.api('/appCatalogs/teamsApps')
    // .filter('distributionMethod eq \'organization\'')
	// .get();
    // console.log(teamsApps);
    const url = `https://graph.microsoft.com/v1.0/appCatalogs/teamsApps?$filter=externalId eq \'${appId}\'`;
    console.log(url);
    await getData(url);
  }

async function getApps() {
    // let teamsApps = await graphClient.api('/appCatalogs/teamsApps')
    // .filter('distributionMethod eq \'organization\'')
	// .get();
    // console.log(teamsApps);
    await getData('https://graph.microsoft.com/v1.0/appCatalogs/teamsApps?$filter=distributionMethod eq \'organization\'');
  }

  async function getData(url) {
    return new Promise(async (resolve) => {
        var config = {
            method: 'get',
            url: url,
            headers: {
                'Authorization': await getToken(),
                'Content-Type': 'application/zip'
            }
        };

        try {
            const response = await axios(config);
            console.log(response);
            console.log(JSON.stringify(response.data));
            //resolve(response.data);
        } catch (error) {
            console.log(error.response.data);
        }
    })  
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

        try {
            const response = await axios(config);
            console.log(response);
            //resolve(response.data);
        } catch (error) {
            console.log(error.response.data);
        }
    })  
}

async function publishApp() {
    const teamsApp = fs.readFile('./package/appPackage.local.zip', async (err, data) => {
        if (err) throw err;
        await PostData(data, 'https://graph.microsoft.com/v1.0/appCatalogs/teamsApps?requiresReview=true');
        console.log('App published');
    });
}

async function updateApp(appId: string) {
    const teamsApp = fs.readFile('./package/appPackage.local.zip', async (err, data) => {
        let stop = false;
        const startTimestamp = Date.now();
        let endTimestamp = Date.now();
        if (err) throw err;
        try { 
            while(!stop) {
                const teamsApp = await graphClient.api(`/appCatalogs/teamsApps`)
                .filter(`externalId  eq '${appId}'`)
                .get();
                console.log(teamsApp);
                if(teamsApp.value.length == 0) {
                    console.log(`${Date.now()} : App not found`);
                    // Sleep for 1 minute
                    showTimepassed(Date.now(), startTimestamp);
                    await new Promise(resolve => setTimeout(resolve, 60000));                    
                }
                else {
                    const internalAppId = teamsApp.value[0].id;
                    const response = await PostData(data, 'https://graph.microsoft.com/v1.0/appCatalogs/teamsApps/' + internalAppId + '/appDefinitions');
                    console.log(response);
                    stop = true;
                    let endTimestamp = Date.now();
                }
            }
            showTimepassed(endTimestamp, startTimestamp);
        } catch (error) {
            console.log(error);
        }
        console.log('App updated');
    });
}

function showTimepassed(endTimestamp: number, startTimestamp: number) {
    const timePassed = endTimestamp - startTimestamp;
    const hoursPassed = Math.floor(timePassed / 3600000); // 1 Hour = 36000 Milliseconds
    const minutesPassed = Math.floor((timePassed % 3600000) / 60000); // 1 Minute = 60000 Milliseconds
    console.log(`Time passed: ${hoursPassed} Hours, ${minutesPassed} Minutes`);
}

async function updateInternalApp(internalAppID: string) {
    console.log(`Updating app ${internalAppID}`);
    const teamsApp = fs.readFile('./package/appPackage.local.zip', async (err, data) => {
        if (err) throw err;
        try { 
            const response = await PostData(data, 'https://graph.microsoft.com/v1.0/appCatalogs/teamsApps/' + internalAppID + '/appDefinitions');
            console.log(response);
        } catch (error) {
            console.log(error);
        }
        console.log('App updated');
    });
}

async function approveApp(appId: string) {
    const teamsApp = await graphClient.api(`/appCatalogs/teamsApps`)
    .filter(`externalId  eq '${appId}'`)
	.get();
    console.log(teamsApp);
    if(teamsApp.value.length == 0) {
        console.log('App not found');
        return;
    }

    const internalAppId = teamsApp.value[0].id;

    const appDefinition = await graphClient.api(`/appCatalogs/teamsApps/${internalAppId}/appDefinitions`)
	.get();
    console.log(appDefinition);
    const etag = appDefinition.value[0]['@odata.etag'];
    const appDefinitionId = appDefinition.value[0].id;

    let newApDefinition = await patchData(`https://graph.microsoft.com/v1.0/appCatalogs/teamsApps/${internalAppId}/appDefinitions/${appDefinitionId}`, etag);
    console.log(newApDefinition);
}

async function patchData(url, etag) {
    return new Promise(async (resolve) => {
        var config = {
            method: 'patch',
            url: url,
            headers: {
                'Authorization': await getToken(),
                'Content-Type': 'application/json',
                'If-Match': etag
            },
            data: {
                publishingState: 'published'
            }
        };

        const response = await axios(config);
        console.log(response);
    })
}

if(command === 'listApp') getAppId();
if(command === 'list') getApps();
if(command === 'publish') publishApp();
if(command === 'approve') approveApp(appId);
if(command === 'update') updateApp(appId);
if(command === 'updateInternal') updateInternalApp(appId);
if(command === 'token') console.log(getToken(true));
