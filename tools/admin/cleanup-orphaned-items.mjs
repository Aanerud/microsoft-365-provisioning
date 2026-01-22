#!/usr/bin/env node
import { ClientSecretCredential } from '@azure/identity';
import { Client } from '@microsoft/microsoft-graph-client';
import { TokenCredentialAuthenticationProvider } from '@microsoft/microsoft-graph-client/authProviders/azureTokenCredentials/index.js';
import dotenv from 'dotenv';

dotenv.config();

const tenantId = process.env.AZURE_TENANT_ID;
const clientId = process.env.AZURE_CLIENT_ID;
const clientSecret = process.env.AZURE_CLIENT_SECRET;

console.log('🗑️  Cleaning Up Orphaned External Items...\n');

const credential = new ClientSecretCredential(tenantId, clientId, clientSecret);
const authProvider = new TokenCredentialAuthenticationProvider(credential, {
  scopes: ['https://graph.microsoft.com/.default']
});

const betaClient = Client.initWithMiddleware({
  authProvider,
  defaultVersion: 'beta'
});

const itemsToDelete = [
  'person-lars-hansen-a830edad9050849coep9vqp9bog-onmicrosoft-com',
  'person-kari-andersen-a830edad9050849coep9vqp9bog-onmicrosoft-com'
];

try {
  console.log('✅ Successfully authenticated with client credentials\n');
  console.log(`Deleting ${itemsToDelete.length} orphaned items...\n`);

  for (const itemId of itemsToDelete) {
    try {
      await betaClient
        .api(`/external/connections/m365provisionpeople/items/${itemId}`)
        .delete();

      console.log(`✅ Deleted: ${itemId}`);
    } catch (error) {
      if (error.statusCode === 404) {
        console.log(`⚠️  Already deleted: ${itemId}`);
      } else {
        console.log(`❌ Failed to delete: ${itemId}`);
        console.log(`   Error: ${error.statusCode} - ${error.message}`);
      }
    }
  }

  console.log('\n✅ Cleanup complete!\n');

} catch (error) {
  console.error('❌ Error:', error.message);
}
