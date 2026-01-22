#!/usr/bin/env node
import { ClientSecretCredential } from '@azure/identity';
import { Client } from '@microsoft/microsoft-graph-client';
import { TokenCredentialAuthenticationProvider } from '@microsoft/microsoft-graph-client/authProviders/azureTokenCredentials/index.js';
import dotenv from 'dotenv';

dotenv.config();

const tenantId = process.env.AZURE_TENANT_ID;
const clientId = process.env.AZURE_CLIENT_ID;
const clientSecret = process.env.AZURE_CLIENT_SECRET;

if (!tenantId || !clientId || !clientSecret) {
  console.error('❌ Missing required environment variables');
  process.exit(1);
}

console.log('🔍 Checking Application Permissions for Graph Connectors...\n');

try {
  const credential = new ClientSecretCredential(tenantId, clientId, clientSecret);

  const authProvider = new TokenCredentialAuthenticationProvider(credential, {
    scopes: ['https://graph.microsoft.com/.default']
  });

  console.log('✅ Successfully created client credentials\n');

  // Create Graph client (v1.0)
  const client = Client.initWithMiddleware({
    authProvider,
    defaultVersion: 'v1.0'
  });

  // Try to list connections
  console.log('Testing: GET /external/connections');
  try {
    const connections = await client.api('/external/connections').get();
    console.log(`✅ Can list connections: ${connections.value?.length || 0} found\n`);
  } catch (error) {
    console.log(`❌ Cannot list connections: ${error.statusCode} - ${error.message}\n`);
  }

  // Try to get specific connection
  console.log('Testing: GET /external/connections/m365provisionpeople');
  try {
    const connection = await client.api('/external/connections/m365provisionpeople').get();
    console.log(`✅ Can read connection: ${connection.name}`);
    console.log(`   State: ${connection.state}\n`);
  } catch (error) {
    console.log(`❌ Cannot read connection: ${error.statusCode} - ${error.message}\n`);
  }

  // Try to create an item (beta endpoint)
  console.log('Testing: PUT /beta/external/connections/m365provisionpeople/items/test-item');
  const betaClient = Client.initWithMiddleware({
    authProvider,
    defaultVersion: 'beta'
  });

  const testItem = {
    id: 'test-item',
    content: {
      value: 'Test content',
      type: 'text'
    },
    properties: {
      accountInformation: JSON.stringify({
        userPrincipalName: 'test@example.com'
      })
    },
    acl: [
      {
        type: 'everyone',
        value: 'everyone',
        accessType: 'grant'
      }
    ]
  };

  try {
    await betaClient
      .api('/external/connections/m365provisionpeople/items/test-item')
      .put(testItem);
    console.log('✅ Can create external items with app-only auth!\n');

    // Clean up test item
    await betaClient
      .api('/external/connections/m365provisionpeople/items/test-item')
      .delete();
    console.log('✅ Cleaned up test item\n');
  } catch (error) {
    console.log(`❌ Cannot create external items: ${error.statusCode} - ${error.message}`);
    if (error.body) {
      console.log(`   Details: ${error.body}\n`);
    }
  }

  console.log('═══════════════════════════════════════════════════════════');
  console.log('Required Application Permissions for Graph Connectors:');
  console.log('  • ExternalConnection.ReadWrite.OwnedBy (or .All)');
  console.log('  • ExternalItem.ReadWrite.OwnedBy (or .All)');
  console.log('═══════════════════════════════════════════════════════════\n');

} catch (error) {
  console.error('❌ Error:', error.message);
}
