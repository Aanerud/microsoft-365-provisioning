#!/usr/bin/env node
import { Client } from '@microsoft/microsoft-graph-client';
import { TokenCredentialAuthenticationProvider } from '@microsoft/microsoft-graph-client/authProviders/azureTokenCredentials/index.js';
import { DeviceCodeCredential } from '@azure/identity';
import dotenv from 'dotenv';
import fs from 'fs';

dotenv.config();

const cachePath = `${process.env.HOME}/.m365-provision/token-cache.json`;
const cacheData = JSON.parse(fs.readFileSync(cachePath, 'utf-8'));

// Use the cached token directly
const client = Client.init({
  authProvider: (done) => {
    done(null, cacheData.accessToken);
  }
});

console.log('\n🔍 Testing Graph Connector permissions...\n');

// Test 1: List connections
try {
  console.log('Test 1: List all external connections');
  const connections = await client.api('/external/connections').get();
  console.log(`✅ Success! Found ${connections.value?.length || 0} connections`);
  if (connections.value?.length > 0) {
    connections.value.forEach(conn => {
      console.log(`  - ${conn.id} (${conn.state})`);
    });
  }
} catch (error) {
  console.log(`❌ Failed: ${error.statusCode} - ${error.message}`);
}

// Test 2: Get specific connection
try {
  console.log('\nTest 2: Get m365provisionpeople connection');
  const connection = await client.api('/external/connections/m365provisionpeople').get();
  console.log(`✅ Success! State: ${connection.state}`);
} catch (error) {
  console.log(`❌ Failed: ${error.statusCode} - ${error.message}`);
}

// Test 3: Try to create a simple item
try {
  console.log('\nTest 3: Try to create a simple test item');
  const testItem = {
    id: 'test-item-123',
    content: { value: 'Test content', type: 'text' },
    properties: {
      accountInformation: JSON.stringify({
        userPrincipalName: 'ingrid.johansen@a830edad9050849coep9vqp9bog.onmicrosoft.com'
      })
    },
    acl: [{ type: 'everyone', value: 'everyone', accessType: 'grant' }]
  };

  const result = await client
    .api('/external/connections/m365provisionpeople/items/test-item-123')
    .put(testItem);
  console.log('✅ Success! Item created');
} catch (error) {
  console.log(`❌ Failed: ${error.statusCode} - ${error.message}`);
  if (error.body) {
    console.log('Error details:', JSON.stringify(JSON.parse(error.body), null, 2));
  }
}
