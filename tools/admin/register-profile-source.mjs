#!/usr/bin/env node
/**
 * Register Profile Source
 *
 * Fixes the Graph Connector profile experience by:
 * 1. Registering the connection as a profile source
 * 2. Adding it to the prioritized sources list
 *
 * This enables Profile Experience (id: 16) and should fix:
 * - "Users: -" showing in Admin Portal
 * - Data not appearing in user profile cards
 * - Copilot not finding people by connector data
 *
 * Requirements:
 * - PeopleSettings.ReadWrite.All application permission
 * - Connection must exist with contentCategory: 'people'
 */
import { ClientSecretCredential } from '@azure/identity';
import { Client } from '@microsoft/microsoft-graph-client';
import { TokenCredentialAuthenticationProvider } from '@microsoft/microsoft-graph-client/authProviders/azureTokenCredentials/index.js';
import dotenv from 'dotenv';

dotenv.config();

const CONNECTION_ID = process.argv[2] || 'm365people3';
const DISPLAY_NAME = process.argv[3] || 'M365 Custom Properties';
const WEB_URL = process.argv[4] || process.env.SHAREPOINT_URL || 'https://textcraft.sharepoint.com';

const { AZURE_TENANT_ID, AZURE_CLIENT_ID, AZURE_CLIENT_SECRET } = process.env;

if (!AZURE_TENANT_ID || !AZURE_CLIENT_ID || !AZURE_CLIENT_SECRET) {
  console.error('Missing required environment variables: AZURE_TENANT_ID, AZURE_CLIENT_ID, AZURE_CLIENT_SECRET');
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

console.log('='.repeat(70));
console.log('REGISTER PROFILE SOURCE');
console.log('='.repeat(70));
console.log(`Connection ID: ${CONNECTION_ID}`);
console.log(`Display Name: ${DISPLAY_NAME}`);
console.log(`Web URL: ${WEB_URL}`);
console.log();

// Step 0: Verify connection exists
console.log('0. VERIFYING CONNECTION EXISTS');
console.log('-'.repeat(50));
try {
  const connection = await betaClient.api(`/external/connections/${CONNECTION_ID}`).get();
  console.log(`   ✅ Connection exists: ${connection.name}`);
  console.log(`   State: ${connection.state}`);
  console.log(`   Content Category: ${connection.contentCategory}`);

  if (connection.contentCategory !== 'people') {
    console.log();
    console.log('   ⚠️  WARNING: contentCategory is not "people"');
    console.log('   Profile source registration may not work as expected');
    console.log('   Consider recreating the connection with contentCategory: "people"');
  }
} catch (e) {
  console.log(`   ❌ Connection not found: ${e.message}`);
  console.log('   Create the connection first with: npm run enrich:setup');
  process.exit(1);
}

// Step 1: Register as profile source
console.log('\n1. REGISTERING AS PROFILE SOURCE');
console.log('-'.repeat(50));

const profileSourcePayload = {
  sourceId: CONNECTION_ID,
  displayName: DISPLAY_NAME,
  kind: 'Connector',  // REQUIRED - Microsoft internal pipeline checks for this property
  webUrl: WEB_URL,
};

console.log('   Payload:', JSON.stringify(profileSourcePayload, null, 2).replace(/\n/g, '\n   '));

try {
  await betaClient.api('/admin/people/profileSources').post(profileSourcePayload);
  console.log('   ✅ Successfully registered as profile source (with kind: Connector)');
} catch (e) {
  if (e.statusCode === 409) {
    console.log('   ℹ️  Profile source already exists (409 Conflict)');
    // Verify kind is set - if not, delete and re-register
    try {
      const sources = await betaClient.api('/admin/people/profileSources').get();
      const existing = sources.value?.find(s => s.sourceId === CONNECTION_ID);
      if (existing && !existing.kind) {
        console.log('   ⚠️  Existing source is MISSING kind property - re-registering...');
        await betaClient.api(`/admin/people/profileSources(sourceId='${CONNECTION_ID}')`).delete();
        await betaClient.api('/admin/people/profileSources').post(profileSourcePayload);
        console.log('   ✅ Re-registered with kind: Connector');
      } else {
        console.log(`   ✅ kind is set: ${existing?.kind}`);
      }
    } catch (verifyErr) {
      console.log(`   ⚠️  Could not verify kind: ${verifyErr.message}`);
    }
  } else if (e.statusCode === 403) {
    console.log(`   ❌ Permission denied (403)`);
    console.log('   Ensure PeopleSettings.ReadWrite.All application permission is granted');
    console.log();
    console.log('   To grant permission:');
    console.log('   1. Go to Azure Portal > App Registrations > Your App');
    console.log('   2. API Permissions > Add > Microsoft Graph > Application');
    console.log('   3. Search for PeopleSettings.ReadWrite.All');
    console.log('   4. Grant admin consent');
    process.exit(1);
  } else {
    console.log(`   ❌ Failed: ${e.message}`);
    if (e.body?.error?.message) {
      console.log(`   Details: ${e.body.error.message}`);
    }
    if (e.statusCode === 400) {
      console.log('   This may be a tenant configuration issue or API format problem');
    }
    process.exit(1);
  }
}

// Step 2: Get current profile property settings
console.log('\n2. GETTING PROFILE PROPERTY SETTINGS');
console.log('-'.repeat(50));

let settingsId = null;
let currentSources = [];

try {
  const response = await betaClient.api('/admin/people/profilePropertySettings').get();
  const settings = response.value?.[0];

  if (settings) {
    settingsId = settings.id;
    currentSources = settings.prioritizedSourceUrls || [];
    console.log(`   ✅ Found settings with ID: ${settingsId}`);
    console.log(`   Current prioritized sources: ${currentSources.length}`);
  } else {
    console.log('   ⚠️  No profile property settings found');
    console.log('   Skipping prioritization step');
  }
} catch (e) {
  console.log(`   ❌ Failed to get settings: ${e.message}`);
  if (e.statusCode === 403) {
    console.log('   Missing PeopleSettings.ReadWrite.All permission');
  }
}

// Step 3: Add to prioritized sources
if (settingsId) {
  console.log('\n3. ADDING TO PRIORITIZED SOURCES');
  console.log('-'.repeat(50));

  const sourceUrl = `https://graph.microsoft.com/beta/admin/people/profileSources(sourceId='${CONNECTION_ID}')`;

  // Check if already in list
  if (currentSources.some(url => url.includes(CONNECTION_ID))) {
    console.log('   ✅ Already in prioritized sources');
  } else {
    // Add to front of list (highest priority)
    const updatedSources = [sourceUrl, ...currentSources];

    console.log(`   Adding to priority position 1 (highest)`);
    console.log(`   New source URL: ${sourceUrl}`);

    try {
      await betaClient.api(`/admin/people/profilePropertySettings/${settingsId}`).patch({
        prioritizedSourceUrls: updatedSources,
      });
      console.log('   ✅ Successfully added to prioritized sources');
    } catch (e) {
      console.log(`   ❌ Failed to update prioritization: ${e.message}`);
      if (e.body?.error?.message) {
        console.log(`   Details: ${e.body.error.message}`);
      }
    }
  }
}

// Step 4: Verify registration
console.log('\n4. VERIFYING REGISTRATION');
console.log('-'.repeat(50));

try {
  const sources = await betaClient.api('/admin/people/profileSources').get();
  const ourSource = sources.value?.find(s => s.sourceId === CONNECTION_ID);

  if (ourSource) {
    console.log('   ✅ Profile source verified:');
    console.log(`      Source ID: ${ourSource.sourceId}`);
    console.log(`      Display Name: ${ourSource.displayName}`);
    console.log(`      Kind: ${ourSource.kind || '⚠️  NOT SET!'}`);
    console.log(`      Web URL: ${ourSource.webUrl}`);
    if (!ourSource.kind) {
      console.log('   ⚠️  CRITICAL: kind property was not persisted by the API!');
    }
  } else {
    console.log('   ⚠️  Profile source not found in list');
  }
} catch (e) {
  console.log(`   ⚠️  Could not verify: ${e.message}`);
}

// Verify prioritization
if (settingsId) {
  try {
    const response = await betaClient.api('/admin/people/profilePropertySettings').get();
    const settings = response.value?.[0];
    const urls = settings?.prioritizedSourceUrls || [];

    const ourIndex = urls.findIndex(url => url.includes(CONNECTION_ID));
    if (ourIndex >= 0) {
      console.log(`   ✅ In prioritized sources at position ${ourIndex + 1}`);
    } else {
      console.log('   ⚠️  Not in prioritized sources');
    }
  } catch (e) {
    // Ignore
  }
}

// Summary
console.log('\n' + '='.repeat(70));
console.log('NEXT STEPS');
console.log('='.repeat(70));
console.log(`
1. Wait 6-24 hours for Microsoft to process the profile source registration
   and enable Profile Experience (id: 16) for the connection.

2. Check Admin Portal:
   - Go to Microsoft 365 Admin Center > Search & Intelligence > Data Sources
   - Look for "${DISPLAY_NAME}" connection
   - "Users" column should show a number instead of "-"

3. Verify in user profiles:
   - Open a user's profile card in Teams or Outlook
   - Look for skills/notes from the connector

4. Test with Copilot:
   - Try: "Find people with skills in TypeScript"
   - Try: "Who knows about data migration?"

5. If issues persist, run the diagnostic:
   node tools/debug/check-profile-source.mjs ${CONNECTION_ID}

Note: Profile Experience enablement is handled by Microsoft based on
profile source registration. You cannot directly set enabledContentExperiences.
`);
