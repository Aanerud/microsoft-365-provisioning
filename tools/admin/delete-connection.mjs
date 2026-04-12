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

const args = process.argv.slice(2);
const CONNECTION_ID = args[0] || 'm365provisionpeople';
const SKIP_CLEANUP = args.includes('--skip-cleanup');
const AAD_SOURCE_ID = '4ce763dd-9214-4eff-af7c-da491cc3782d';

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

function extractSourceId(url) {
  const match = url?.match(/sourceId='([^']+)'/);
  return match ? match[1] : null;
}

async function cleanupProfileSources() {
  console.log('🧹 Cleaning up profile sources/prioritization...');

  const connections = await client.api('/external/connections').get();
  const existingConnectionIds = new Set((connections.value || []).map(c => c.id));
  existingConnectionIds.delete(CONNECTION_ID);

  const sources = await client.api('/admin/people/profileSources').get();
  const profileSources = sources.value || [];
  const profileSourceIds = new Set(profileSources.map(source => source.sourceId));

  const settings = await client.api('/admin/people/profilePropertySettings').get();
  const setting = settings.value?.[0];
  let prioritizedSourceIds = new Set();

  if (setting) {
    const currentUrls = setting.prioritizedSourceUrls || [];
    const validSourceIds = new Set([...existingConnectionIds, AAD_SOURCE_ID]);
    const validUrls = currentUrls.filter(url => {
      const sourceId = extractSourceId(url);
      if (!sourceId) return false;
      if (sourceId === CONNECTION_ID) return false;
      if (!profileSourceIds.has(sourceId)) return false;
      return validSourceIds.has(sourceId);
    });

    const changed =
      currentUrls.length !== validUrls.length ||
      currentUrls.some((url, index) => url !== validUrls[index]);

    if (changed) {
      await client.api(`/admin/people/profilePropertySettings/${setting.id}`).patch({
        prioritizedSourceUrls: validUrls,
      });
      console.log('✅ Prioritized sources updated');
    } else {
      console.log('✅ Prioritized sources already clean');
    }

    prioritizedSourceIds = new Set(validUrls.map(extractSourceId).filter(Boolean));
  } else {
    console.log('⚠️  No profile property settings found');
  }

  const orphaned = profileSources.filter(
    source => source.sourceId !== AAD_SOURCE_ID && !existingConnectionIds.has(source.sourceId)
  );

  if (orphaned.length > 0) {
    console.log(`Deleting ${orphaned.length} orphaned profile sources...`);
    for (const source of orphaned) {
      if (prioritizedSourceIds.has(source.sourceId)) {
        console.log(`  ⚠️  Skipping ${source.sourceId} (still in prioritized list)`);
        continue;
      }
      try {
        await client.api(`/admin/people/profileSources(sourceId='${source.sourceId}')`).delete();
        console.log(`  ✅ Deleted ${source.sourceId}`);
      } catch (error) {
        console.log(`  ❌ Failed to delete ${source.sourceId}: ${error.message}`);
      }
    }
  }
}

console.log(`🗑️  Deleting connection: ${CONNECTION_ID}...`);

try {
  await client.api(`/external/connections/${CONNECTION_ID}`).delete();
  console.log('✅ Connection deleted successfully');
  console.log('');
  console.log('⏳ Wait 5-15 minutes for deletion to complete before recreating.');
  console.log('');
  console.log('Next steps:');
  console.log('  1. npm run enrich:connector-only -- --csv config/textcraft-europe.csv --setup');
  console.log('     (Creates connection, registers schema, registers as profile source, ingests items)');
  console.log('');
  console.log('  2. Wait 6+ hours for Microsoft 365 to index the data');
  console.log('');
  console.log('  3. Test in Copilot with queries like:');
  console.log('     - "Find people with skills in Strategic Planning"');
  console.log('     - "Who has TypeScript skills?"');
  if (!SKIP_CLEANUP) {
    await cleanupProfileSources();
  } else {
    console.log('ℹ️  Skipping profile source cleanup (--skip-cleanup)');
  }
} catch (error) {
  if (error.statusCode === 404) {
    console.log('ℹ️  Connection does not exist (already deleted or never created)');
    if (!SKIP_CLEANUP) {
      await cleanupProfileSources();
    }
  } else {
    console.log(`❌ Failed: ${error.statusCode} - ${error.message}`);
    if (error.body) {
      console.log('   Details:', error.body);
    }
  }
}
