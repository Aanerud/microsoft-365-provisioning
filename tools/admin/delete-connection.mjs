#!/usr/bin/env node
/**
 * Delete Graph Connector Connection
 * Uses OAuth 2.0 Client Credentials Flow (Application permissions)
 */
import { ClientSecretCredential } from '@azure/identity';
import { Client } from '@microsoft/microsoft-graph-client';
import { TokenCredentialAuthenticationProvider } from '@microsoft/microsoft-graph-client/authProviders/azureTokenCredentials/index.js';
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
  scopes: ['https://graph.microsoft.com/.default'],
});

// Use beta endpoint for Graph Connectors
const client = Client.initWithMiddleware({
  authProvider,
  defaultVersion: 'beta',
});

console.log(`🗑️  Deleting connection: ${CONNECTION_ID}...`);

try {
  await client.api(`/external/connections/${CONNECTION_ID}`).delete();
  console.log('✅ Connection deleted successfully');
  console.log('');
  console.log('Next steps:');
  console.log('  1. npm run enrich-profiles:setup   # Create new connection with updated schema');
  console.log('  2. npm run enrich-profiles:wait    # Wait for schema to be ready (~10 min)');
  console.log('  3. npm run enrich-profiles         # Ingest items');
} catch (error) {
  if (error.statusCode === 404) {
    console.log('ℹ️  Connection does not exist (already deleted or never created)');
  } else {
    console.log(`❌ Failed: ${error.statusCode} - ${error.message}`);
    if (error.body) {
      console.log('   Details:', error.body);
    }
  }
}
