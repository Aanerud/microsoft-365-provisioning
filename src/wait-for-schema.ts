#!/usr/bin/env node
import { GraphClient } from './graph-client.js';
import { PeopleConnectionManager } from './people-connector/connection-manager.js';
import dotenv from 'dotenv';

dotenv.config();

async function waitForSchema() {
  const tenantId = process.env.AZURE_TENANT_ID || '';
  const clientId = process.env.AZURE_CLIENT_ID || '';
  const connectionId = 'm365provisionpeople';
  const clientSecret = process.env.AZURE_CLIENT_SECRET || '';

  if (!tenantId || !clientId || !clientSecret) {
    throw new Error('AZURE_TENANT_ID, AZURE_CLIENT_ID, and AZURE_CLIENT_SECRET are required.');
  }

  // Authenticate with Graph Connector scopes (app-only)
  console.log('🔐 Authenticating (app-only)...');
  const graphClient = new GraphClient({
    tenantId,
    clientId,
    clientSecret,
  });
  const { betaClient } = graphClient.getClients();

  const connectionManager = new PeopleConnectionManager(betaClient, connectionId);

  console.log('\n⏳ Waiting for schema to be ready...');
  console.log('   This can take 5-10 minutes for Graph Connectors\n');

  let attempts = 0;
  const maxAttempts = 120; // 120 * 5 seconds = 10 minutes

  while (attempts < maxAttempts) {
    try {
      const schemaStatus = await connectionManager.getSchema();
      const state = schemaStatus?.state ?? schemaStatus?.status;

      console.log(`[${new Date().toLocaleTimeString()}] Schema state: ${state}`);

      if (state === 'ready') {
        console.log('\n✅ Schema is READY! You can now run the ingestion.');
        console.log('\nRun: npm run enrich-profiles -- --csv config/agents-test-enrichment.csv');
        process.exit(0);
      } else if (state === 'failed') {
        console.error(`\n❌ Schema registration FAILED: ${schemaStatus?.failureReason || 'Unknown reason'}`);
        process.exit(1);
      }

      attempts++;
      await new Promise(resolve => setTimeout(resolve, 5000)); // Wait 5 seconds

    } catch (error: any) {
      console.error(`Error checking connection: ${error.message}`);
      process.exit(1);
    }
  }

  console.error('\n❌ Timeout: Schema did not become ready within 10 minutes');
  process.exit(1);
}

waitForSchema().catch(console.error);
