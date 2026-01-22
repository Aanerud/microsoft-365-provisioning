#!/usr/bin/env node
import { GraphClient } from './graph-client.js';
import { BrowserAuthServer } from './auth/browser-auth-server.js';
import { PeopleConnectionManager } from './people-connector/connection-manager.js';
import dotenv from 'dotenv';

dotenv.config();

async function waitForSchema() {
  const tenantId = process.env.AZURE_TENANT_ID || '';
  const clientId = process.env.AZURE_CLIENT_ID || '';
  const connectionId = 'm365provisionpeople';

  // Authenticate with Graph Connector scopes
  console.log('🔐 Authenticating...');
  const authServer = new BrowserAuthServer({
    tenantId,
    clientId,
    scopes: [
      'User.Read.All',
      'Directory.Read.All',
      'ExternalConnection.ReadWrite.All',
      'ExternalItem.ReadWrite.All'
    ]
  });
  const authResult = await authServer.authenticate();

  // Initialize clients
  const graphClient = new GraphClient({ accessToken: authResult.accessToken });
  const { client, betaClient } = graphClient.getClients();

  const connectionManager = new PeopleConnectionManager(client, betaClient, connectionId);

  console.log('\n⏳ Waiting for schema to be ready...');
  console.log('   This can take 5-10 minutes for Graph Connectors\n');

  let attempts = 0;
  const maxAttempts = 120; // 120 * 5 seconds = 10 minutes

  while (attempts < maxAttempts) {
    try {
      const connection = await connectionManager.getConnection();
      const state = connection.state;

      console.log(`[${new Date().toLocaleTimeString()}] Schema state: ${state}`);

      if (state === 'ready') {
        console.log('\n✅ Schema is READY! You can now run the ingestion.');
        console.log('\nRun: npm run enrich-profiles -- --csv config/agents-test-enrichment.csv');
        process.exit(0);
      } else if (state === 'failed') {
        console.error(`\n❌ Schema registration FAILED: ${connection.failureReason}`);
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
