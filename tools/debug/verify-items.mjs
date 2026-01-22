#!/usr/bin/env node
import { ClientSecretCredential } from '@azure/identity';
import { Client } from '@microsoft/microsoft-graph-client';
import { TokenCredentialAuthenticationProvider } from '@microsoft/microsoft-graph-client/authProviders/azureTokenCredentials/index.js';
import dotenv from 'dotenv';

dotenv.config();

const tenantId = process.env.AZURE_TENANT_ID;
const clientId = process.env.AZURE_CLIENT_ID;
const clientSecret = process.env.AZURE_CLIENT_SECRET;

console.log('🔍 Verifying External Items in Graph Connector...\n');

const credential = new ClientSecretCredential(tenantId, clientId, clientSecret);
const authProvider = new TokenCredentialAuthenticationProvider(credential, {
  scopes: ['https://graph.microsoft.com/.default']
});

const betaClient = Client.initWithMiddleware({
  authProvider,
  defaultVersion: 'beta'
});

try {
  // Try to get specific items
  const itemIds = [
    'person-ingrid-johansen-a830edad9050849coep9vqp9bog-onmicrosoft-com',
    'person-lars-hansen-a830edad9050849coep9vqp9bog-onmicrosoft-com',
    'person-kari-andersen-a830edad9050849coep9vqp9bog-onmicrosoft-com'
  ];

  console.log('✅ Successfully authenticated with client credentials\n');
  console.log('Checking ingested items:\n');

  for (const itemId of itemIds) {
    try {
      const item = await betaClient
        .api(`/external/connections/m365provisionpeople/items/${itemId}`)
        .get();

      console.log(`✅ ${itemId}`);
      console.log(`   Properties:`);
      if (item.properties) {
        const props = item.properties;
        if (props.accountInformation) {
          const acct = JSON.parse(props.accountInformation);
          console.log(`     - Email: ${acct.userPrincipalName}`);
        }
        if (props.skills) {
          console.log(`     - Skills: ${props.skills.length} items`);
        }
        if (props.interests) {
          console.log(`     - Interests: ${props.interests.length} items`);
        }
        if (props.aboutMe) {
          console.log(`     - About: ${props.aboutMe.substring(0, 50)}...`);
        }
      }
      console.log('');
    } catch (error) {
      console.log(`❌ ${itemId}`);
      console.log(`   Error: ${error.statusCode} - ${error.message}\n`);
    }
  }

} catch (error) {
  console.error('❌ Error:', error.message);
}
