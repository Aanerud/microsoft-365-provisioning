#!/usr/bin/env node
import { ClientSecretCredential } from '@azure/identity';
import { Client } from '@microsoft/microsoft-graph-client';
import { TokenCredentialAuthenticationProvider } from '@microsoft/microsoft-graph-client/authProviders/azureTokenCredentials/index.js';
import dotenv from 'dotenv';

dotenv.config();

const tenantId = process.env.AZURE_TENANT_ID;
const clientId = process.env.AZURE_CLIENT_ID;
const clientSecret = process.env.AZURE_CLIENT_SECRET;

console.log('🔍 Verifying All External Items in Graph Connector...\n');

const credential = new ClientSecretCredential(tenantId, clientId, clientSecret);
const authProvider = new TokenCredentialAuthenticationProvider(credential, {
  scopes: ['https://graph.microsoft.com/.default']
});

const betaClient = Client.initWithMiddleware({
  authProvider,
  defaultVersion: 'beta'
});

try {
  // Expected items (from CSV)
  const expectedItems = [
    'person-ingrid-johansen-a830edad9050849coep9vqp9bog-onmicrosoft-com',
    'person-ola-nordmann-a830edad9050849coep9vqp9bog-onmicrosoft-com'
  ];

  // Items that should be deleted
  const deletedItems = [
    'person-lars-hansen-a830edad9050849coep9vqp9bog-onmicrosoft-com',
    'person-kari-andersen-a830edad9050849coep9vqp9bog-onmicrosoft-com'
  ];

  console.log('✅ Successfully authenticated with client credentials\n');

  console.log('📋 Checking EXPECTED items (should exist):\n');

  for (const itemId of expectedItems) {
    try {
      const item = await betaClient
        .api(`/external/connections/m365provisionpeople/items/${itemId}`)
        .get();

      console.log(`✅ ${itemId}`);
      if (item.properties) {
        const props = item.properties;
        if (props.accountInformation) {
          const acct = JSON.parse(props.accountInformation);
          console.log(`   Email: ${acct.userPrincipalName}`);
        }
        if (props.skills) {
          console.log(`   Skills: ${Array.isArray(props.skills) ? props.skills.length : 0} items`);
        }
      }
      console.log('');
    } catch (error) {
      console.log(`❌ ${itemId}`);
      console.log(`   Error: ${error.statusCode} - ${error.message}\n`);
    }
  }

  console.log('\n📋 Checking DELETED users (should NOT exist or should be cleaned):\n');

  for (const itemId of deletedItems) {
    try {
      const item = await betaClient
        .api(`/external/connections/m365provisionpeople/items/${itemId}`)
        .get();

      console.log(`⚠️  ${itemId}`);
      console.log(`   Status: STILL EXISTS (should be deleted manually)`);
      if (item.properties?.accountInformation) {
        const acct = JSON.parse(item.properties.accountInformation);
        console.log(`   Email: ${acct.userPrincipalName}`);
      }
      console.log('');
    } catch (error) {
      if (error.statusCode === 404) {
        console.log(`✅ ${itemId}`);
        console.log(`   Status: Correctly deleted (404 Not Found)\n`);
      } else {
        console.log(`❌ ${itemId}`);
        console.log(`   Error: ${error.statusCode} - ${error.message}\n`);
      }
    }
  }

  console.log('\n═══════════════════════════════════════════════════════════');
  console.log('📝 Note: Option B does not automatically delete external items');
  console.log('    for users removed from CSV. This is by design.');
  console.log('    To clean up old items, delete them manually or implement');
  console.log('    a cleanup routine.');
  console.log('═══════════════════════════════════════════════════════════\n');

} catch (error) {
  console.error('❌ Error:', error.message);
}
