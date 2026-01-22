#!/usr/bin/env node
import { Client } from '@microsoft/microsoft-graph-client';
import { TokenCredentialAuthenticationProvider } from '@microsoft/microsoft-graph-client/authProviders/azureTokenCredentials/index.js';
import { DeviceCodeCredential } from '@azure/identity';
import fs from 'fs';
import dotenv from 'dotenv';

dotenv.config();

// Read cached token
const cachePath = `${process.env.HOME}/.m365-provision/token-cache.json`;
const cacheData = JSON.parse(fs.readFileSync(cachePath, 'utf-8'));

const credential = new DeviceCodeCredential({
  tenantId: process.env.AZURE_TENANT_ID,
  clientId: process.env.AZURE_CLIENT_ID,
  userPromptCallback: () => {}
});

const authProvider = new TokenCredentialAuthenticationProvider(credential, {
  scopes: ['https://graph.microsoft.com/.default']
});

const client = Client.initWithMiddleware({ authProvider });

// Check connection
try {
  const connection = await client.api('/external/connections/m365provisionpeople').get();
  console.log('\n📊 Connection Status:');
  console.log(`  ID: ${connection.id}`);
  console.log(`  Name: ${connection.name}`);
  console.log(`  State: ${connection.state}`);
  if (connection.state === 'failed') {
    console.log(`  Failure Reason: ${connection.failureReason}`);
  }
} catch (error) {
  console.error('Error:', error.message);
  console.error('Status:', error.statusCode);
}
