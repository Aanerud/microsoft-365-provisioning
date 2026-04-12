#!/usr/bin/env node
/**
 * Wait for schema to be ready, then ingest items
 * Use this when connection is created but schema is still provisioning
 */
import { ClientSecretCredential } from '@azure/identity';
import { Client } from '@microsoft/microsoft-graph-client';
import { TokenCredentialAuthenticationProvider } from '@microsoft/microsoft-graph-client/authProviders/azureTokenCredentials/index.js';
import fs from 'fs/promises';
import { parse } from 'csv-parse/sync';
import dotenv from 'dotenv';

dotenv.config();

const CONNECTION_ID = process.argv[2] || 'm365people3';
const CSV_PATH = process.argv[3] || 'config/textcraft-europe.csv';
const MAX_WAIT_MINUTES = 10;

const { AZURE_TENANT_ID, AZURE_CLIENT_ID, AZURE_CLIENT_SECRET } = process.env;

if (!AZURE_TENANT_ID || !AZURE_CLIENT_ID || !AZURE_CLIENT_SECRET) {
  console.error('Missing required environment variables');
  process.exit(1);
}

const credential = new ClientSecretCredential(
  AZURE_TENANT_ID,
  AZURE_CLIENT_ID,
  AZURE_CLIENT_SECRET
);

const authProvider = new TokenCredentialAuthenticationProvider(credential, {
  scopes: ['https://graph.microsoft.com/.default'],
});

const betaClient = Client.initWithMiddleware({ authProvider, defaultVersion: 'beta' });

console.log(`Connection ID: ${CONNECTION_ID}`);
console.log(`CSV Path: ${CSV_PATH}`);
console.log(`Max wait: ${MAX_WAIT_MINUTES} minutes\n`);

// Wait for schema to be ready
console.log('Waiting for schema to be ready...');
const startTime = Date.now();
const maxWaitMs = MAX_WAIT_MINUTES * 60 * 1000;

while (Date.now() - startTime < maxWaitMs) {
  try {
    const connection = await betaClient.api(`/external/connections/${CONNECTION_ID}`).get();
    const elapsed = Math.round((Date.now() - startTime) / 1000);

    if (connection.state === 'ready') {
      console.log(`✅ Schema is ready! (after ${elapsed}s)`);
      console.log(`   Content Category: ${connection.contentCategory}`);
      break;
    } else if (connection.state === 'failed') {
      console.log(`❌ Schema failed: ${connection.failureReason}`);
      process.exit(1);
    }

    console.log(`   State: ${connection.state} (${elapsed}s elapsed)`);
  } catch (e) {
    console.log(`   Error checking status: ${e.message}`);
  }

  await new Promise(resolve => setTimeout(resolve, 5000));
}

// Check final state
const finalConnection = await betaClient.api(`/external/connections/${CONNECTION_ID}`).get();
if (finalConnection.state !== 'ready') {
  console.log(`\n⚠️  Schema still not ready after ${MAX_WAIT_MINUTES} minutes`);
  console.log(`   Current state: ${finalConnection.state}`);
  console.log(`   Run this script again later or check the admin portal`);
  process.exit(1);
}

// Check schema labels
console.log('\nChecking schema labels...');
const schema = await betaClient.api(`/external/connections/${CONNECTION_ID}/schema`).get();
for (const prop of schema.properties || []) {
  const labels = prop.labels?.join(', ') || 'none';
  console.log(`   ${prop.name}: [${labels}]`);
}

// Load and parse CSV
console.log(`\nLoading profiles from ${CSV_PATH}...`);
const content = await fs.readFile(CSV_PATH, 'utf-8');
const records = parse(content, {
  columns: true,
  skip_empty_lines: true,
  trim: true,
});

console.log(`Loaded ${records.length} profiles`);

// Ingest items
console.log('\nIngesting items...\n');
let successful = 0;
let failed = 0;

for (let i = 0; i < records.length; i++) {
  const row = records[i];
  const email = row.email;
  const itemId = `person-${email.replace(/@/g, '-').replace(/\./g, '-')}`;

  // Build properties
  const properties = {
    accountInformation: JSON.stringify({ userPrincipalName: email }),
  };

  // Skills
  if (row.skills) {
    let skills = [];
    try {
      skills = JSON.parse(row.skills.replace(/'/g, '"'));
    } catch {
      skills = row.skills.split(',').map(s => s.trim()).filter(Boolean);
    }
    if (skills.length > 0) {
      properties['skills@odata.type'] = 'Collection(String)';
      properties.skills = skills.map(s => JSON.stringify({ displayName: s }));
    }
  }

  // AboutMe
  if (row.aboutMe) {
    properties.aboutMe = JSON.stringify({
      detail: { contentType: 'text', content: row.aboutMe }
    });
  }

  // Custom properties
  const customFields = ['VTeam', 'BenefitPlan', 'CostCenter', 'BuildingAccess', 'ProjectCode', 'WritingStyle', 'Specialization'];
  for (const field of customFields) {
    if (row[field]) {
      properties[field] = row[field];
    }
  }

  const item = {
    id: itemId,
    properties,
    acl: [{ type: 'everyone', value: 'everyone', accessType: 'grant' }],
  };

  try {
    await betaClient.api(`/external/connections/${CONNECTION_ID}/items/${itemId}`).put(item);
    console.log(`[${i + 1}/${records.length}] ✓ ${email}`);
    successful++;
  } catch (error) {
    console.log(`[${i + 1}/${records.length}] ✗ ${email}: ${error.message}`);
    failed++;
  }

  // Small delay
  await new Promise(resolve => setTimeout(resolve, 100));
}

console.log(`\n${'='.repeat(50)}`);
console.log('INGESTION COMPLETE');
console.log(`${'='.repeat(50)}`);
console.log(`Successful: ${successful}`);
console.log(`Failed: ${failed}`);
console.log(`\nNext steps:`);
console.log(`1. Wait 6+ hours for Microsoft 365 to index the data`);
console.log(`2. Test in Copilot: "Find people with TypeScript skills"`);
