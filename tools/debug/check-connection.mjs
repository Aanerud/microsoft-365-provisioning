#!/usr/bin/env node
/**
 * Check Graph Connector Connection Status
 * Uses OAuth 2.0 Client Credentials Flow (Application permissions)
 */
import { Client } from '@microsoft/microsoft-graph-client';
import { TokenCredentialAuthenticationProvider } from '@microsoft/microsoft-graph-client/authProviders/azureTokenCredentials/index.js';
import { ClientSecretCredential } from '@azure/identity';
import dotenv from 'dotenv';

dotenv.config();

const CONNECTION_ID = process.argv[2] || 'm365provisionpeople';

const { AZURE_TENANT_ID, AZURE_CLIENT_ID, AZURE_CLIENT_SECRET } = process.env;

if (!AZURE_TENANT_ID || !AZURE_CLIENT_ID || !AZURE_CLIENT_SECRET) {
  console.error('❌ Missing required environment variables: AZURE_TENANT_ID, AZURE_CLIENT_ID, AZURE_CLIENT_SECRET');
  process.exit(1);
}

console.log('🔐 Authenticating with client credentials...');

const credential = new ClientSecretCredential(
  AZURE_TENANT_ID,
  AZURE_CLIENT_ID,
  AZURE_CLIENT_SECRET
);

const authProvider = new TokenCredentialAuthenticationProvider(credential, {
  scopes: ['https://graph.microsoft.com/.default']
});

// Use beta endpoint for Graph Connectors
const client = Client.initWithMiddleware({
  authProvider,
  defaultVersion: 'beta'
});

// Check connection
console.log(`\n🔍 Checking connection: ${CONNECTION_ID}...`);

try {
  const connection = await client.api(`/external/connections/${CONNECTION_ID}`).get();
  console.log('\n📊 Connection Status:');
  console.log(`  ID: ${connection.id}`);
  console.log(`  Name: ${connection.name}`);
  console.log(`  State: ${connection.state}`);
  console.log(`  Description: ${connection.description || 'N/A'}`);
  if (connection.state === 'failed') {
    console.log(`  Failure Reason: ${connection.failureReason}`);
  }

  // Try to get schema
  try {
    const schema = await client.api(`/external/connections/${CONNECTION_ID}/schema`).get();
    console.log('\n📋 Schema:');
    console.log(`  Base Type: ${schema.baseType}`);
    console.log(`  Properties: ${schema.properties?.length || 0}`);
    if (schema.properties) {
      schema.properties.forEach(p => {
        const labels = p.labels ? ` [${p.labels.join(', ')}]` : '';
        console.log(`    - ${p.name}: ${p.type}${labels}`);
      });
    }
  } catch (schemaError) {
    console.log(`\n📋 Schema: Not yet registered or unavailable`);
  }

} catch (error) {
  if (error.statusCode === 404) {
    console.log('ℹ️  Connection does not exist');
  } else {
    console.error(`❌ Error: ${error.message}`);
    console.error(`   Status: ${error.statusCode}`);
    if (error.body) {
      console.error(`   Details: ${error.body}`);
    }
  }
}
