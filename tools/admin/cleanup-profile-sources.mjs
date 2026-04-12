#!/usr/bin/env node
/**
 * Clean up orphaned profile sources from the prioritization list
 */
import { ClientSecretCredential } from '@azure/identity';
import { Client } from '@microsoft/microsoft-graph-client';
import { TokenCredentialAuthenticationProvider } from '@microsoft/microsoft-graph-client/authProviders/azureTokenCredentials/index.js';
import dotenv from 'dotenv';

dotenv.config();

const KEEP_CONNECTION_ID = process.argv[2] || null;
const AAD_SOURCE_ID = '4ce763dd-9214-4eff-af7c-da491cc3782d';

const { AZURE_TENANT_ID, AZURE_CLIENT_ID, AZURE_CLIENT_SECRET } = process.env;

if (!AZURE_TENANT_ID || !AZURE_CLIENT_ID || !AZURE_CLIENT_SECRET) {
  console.error('❌ Missing required environment variables: AZURE_TENANT_ID, AZURE_CLIENT_ID, AZURE_CLIENT_SECRET');
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

function extractSourceId(url) {
  const match = url?.match(/sourceId='([^']+)'/);
  return match ? match[1] : null;
}

console.log('Cleaning up profile sources...\n');

// Get current connections
const connections = await betaClient.api('/external/connections').get();
const existingConnectionIds = new Set((connections.value || []).map(c => c.id));
console.log('Existing connections:', [...existingConnectionIds]);

// Get current profile sources
const sources = await betaClient.api('/admin/people/profileSources').get();
const profileSources = sources.value || [];
const profileSourceIds = new Set(profileSources.map(source => source.sourceId));

console.log('\nCurrent profile sources:');
for (const source of profileSources) {
  const exists = existingConnectionIds.has(source.sourceId);
  console.log(`  ${source.sourceId}: ${exists ? '✅ exists' : '❌ orphaned'}`);
}

// Update profile property settings to only include existing sources
console.log('\nUpdating profile property settings...');
const settings = await betaClient.api('/admin/people/profilePropertySettings').get();
const setting = settings.value?.[0];
let prioritizedSourceIds = new Set();

if (setting) {
  const currentUrls = setting.prioritizedSourceUrls || [];
  console.log('Current prioritized sources:', currentUrls.length);

  const validSourceIds = new Set([...existingConnectionIds, AAD_SOURCE_ID]);
  const validUrls = currentUrls.filter(url => {
    const sourceId = extractSourceId(url);
    if (!sourceId) return false;
    if (!profileSourceIds.has(sourceId)) return false;
    return validSourceIds.has(sourceId);
  });

  if (KEEP_CONNECTION_ID) {
    const keepIsValid =
      existingConnectionIds.has(KEEP_CONNECTION_ID) &&
      profileSourceIds.has(KEEP_CONNECTION_ID);
    if (!keepIsValid) {
      console.log(`⚠️  Keep connection "${KEEP_CONNECTION_ID}" is not a valid profile source; skipping priority pin.`);
    } else {
      const keepUrl = `https://graph.microsoft.com/beta/admin/people/profileSources(sourceId='${KEEP_CONNECTION_ID}')`;
      const existingIndex = validUrls.indexOf(keepUrl);
      if (existingIndex === -1) {
        validUrls.unshift(keepUrl);
      } else if (existingIndex > 0) {
        validUrls.splice(existingIndex, 1);
        validUrls.unshift(keepUrl);
      }
    }
  }

  const changed =
    currentUrls.length !== validUrls.length ||
    currentUrls.some((url, index) => url !== validUrls[index]);

  if (changed) {
    console.log('Updated prioritized sources:', validUrls.length);
    for (let i = 0; i < validUrls.length; i++) {
      console.log(`  ${i + 1}. ${validUrls[i]}`);
    }

    await betaClient.api(`/admin/people/profilePropertySettings/${setting.id}`).patch({
      prioritizedSourceUrls: validUrls,
    });
    console.log('✅ Profile property settings updated');
  } else {
    console.log('✅ Prioritized sources already clean');
  }

  prioritizedSourceIds = new Set(validUrls.map(extractSourceId).filter(Boolean));
} else {
  console.log('No profile property settings found');
}

// Delete orphaned profile sources (after prioritization cleanup)
const orphaned = profileSources.filter(
  source => source.sourceId !== AAD_SOURCE_ID && !existingConnectionIds.has(source.sourceId)
);
if (orphaned.length > 0) {
  console.log(`\nDeleting ${orphaned.length} orphaned profile sources...`);
  for (const source of orphaned) {
    if (prioritizedSourceIds.has(source.sourceId)) {
      console.log(`  ⚠️  Skipping ${source.sourceId} (still in prioritized list)`);
      continue;
    }
    try {
      await betaClient.api(`/admin/people/profileSources(sourceId='${source.sourceId}')`).delete();
      console.log(`  ✅ Deleted ${source.sourceId}`);
    } catch (e) {
      console.log(`  ❌ Failed to delete ${source.sourceId}: ${e.message}`);
    }
  }
}

console.log('\nDone!');
