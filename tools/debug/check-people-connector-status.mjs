#!/usr/bin/env node
/**
 * Check People Connector Status
 * Verifies:
 * 1. Connection exists with contentCategory: 'people'
 * 2. Schema has personAccount label
 * 3. Profile source is registered
 * 4. Sample item format is correct
 */
import { ClientSecretCredential } from '@azure/identity';
import { Client } from '@microsoft/microsoft-graph-client';
import { TokenCredentialAuthenticationProvider } from '@microsoft/microsoft-graph-client/authProviders/azureTokenCredentials/index.js';
import dotenv from 'dotenv';

dotenv.config();

const CONNECTION_ID = resolveConnectionId();

const { AZURE_TENANT_ID, AZURE_CLIENT_ID, AZURE_CLIENT_SECRET } = process.env;

if (!AZURE_TENANT_ID || !AZURE_CLIENT_ID || !AZURE_CLIENT_SECRET) {
  console.error('Missing required environment variables');
  process.exit(1);
}

if (!CONNECTION_ID) {
  console.error('Missing required connection ID. Provide --connection-id or set CONNECTION_ID/M365_CONNECTION_ID.');
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

console.log('='.repeat(60));
console.log('PEOPLE CONNECTOR STATUS CHECK');
console.log('='.repeat(60));
console.log(`Connection ID: ${CONNECTION_ID}\n`);

// 1. Check connection
console.log('1. CONNECTION DETAILS');
console.log('-'.repeat(40));
try {
  const connection = await betaClient.api(`/external/connections/${CONNECTION_ID}`).get();
  console.log(`   Name: ${connection.name}`);
  console.log(`   State: ${connection.state}`);
  console.log(`   Content Category: ${connection.contentCategory || 'NOT SET'}`);

  if (connection.contentCategory !== 'people') {
    console.log('   ⚠️  WARNING: contentCategory should be "people" for user mapping!');
  } else {
    console.log('   ✅ contentCategory is correctly set to "people"');
  }
} catch (e) {
  console.log(`   ❌ Failed to get connection: ${e.message}`);
}

// 2. Check schema
console.log('\n2. SCHEMA PROPERTIES');
console.log('-'.repeat(40));
try {
  const schema = await betaClient.api(`/external/connections/${CONNECTION_ID}/schema`).header('Prefer', 'include-unknown-enum-members').get();

  let hasPersonAccount = false;
  for (const prop of schema.properties || []) {
    const labels = prop.labels?.join(', ') || 'none';
    console.log(`   ${prop.name} (${prop.type}) - labels: [${labels}]`);

    if (prop.labels?.includes('personAccount')) {
      hasPersonAccount = true;
    }
  }

  if (!hasPersonAccount) {
    console.log('   ⚠️  WARNING: No property with personAccount label found!');
  } else {
    console.log('   ✅ personAccount label found');
  }
} catch (e) {
  console.log(`   ❌ Failed to get schema: ${e.message}`);
}

// 3. Check profile source registration
console.log('\n3. PROFILE SOURCE REGISTRATION');
console.log('-'.repeat(40));
try {
  const sources = await betaClient.api('/admin/people/profileSources').get();
  const ourSource = sources.value?.find(s => s.sourceId === CONNECTION_ID);

  if (ourSource) {
    console.log(`   ✅ Registered as profile source`);
    console.log(`   Display Name: ${ourSource.displayName}`);
    console.log(`   Web URL: ${ourSource.webUrl}`);
  } else {
    console.log('   ⚠️  NOT registered as profile source');
    console.log('   Available sources:');
    for (const s of sources.value || []) {
      console.log(`      - ${s.sourceId}: ${s.displayName}`);
    }
  }
} catch (e) {
  const msg = e.message || e.body?.error?.message || e.code || JSON.stringify(e);
  console.log(`   ❌ Failed to check profile sources: ${msg}`);
  if (e.statusCode) console.log(`      Status: ${e.statusCode}`);
  if (e.body?.error) console.log(`      Error detail: ${JSON.stringify(e.body.error)}`);
  if (e.statusCode === 403) {
    console.log('   (May need PeopleSettings.Read.All permission)');
  }
}

// 4. Check profile property settings (prioritization)
console.log('\n4. PROFILE PROPERTY SETTINGS (PRIORITIZATION)');
console.log('-'.repeat(40));
try {
  const settings = await betaClient.api('/admin/people/profilePropertySettings').get();
  const setting = settings.value?.[0];

  if (setting) {
    const urls = setting.prioritizedSourceUrls || [];
    console.log(`   Prioritized sources (${urls.length}):`);

    const ourUrl = `https://graph.microsoft.com/beta/admin/people/profileSources(sourceId='${CONNECTION_ID}')`;
    let foundOurs = false;

    for (let i = 0; i < urls.length; i++) {
      const isOurs = urls[i].includes(CONNECTION_ID);
      if (isOurs) foundOurs = true;
      console.log(`   ${i + 1}. ${urls[i]}${isOurs ? ' ← OUR CONNECTION' : ''}`);
    }

    if (!foundOurs) {
      console.log('   ⚠️  Our connection is NOT in prioritized sources');
    } else {
      console.log('   ✅ Our connection is in prioritized sources');
    }
  } else {
    console.log('   No profile property settings found');
  }
} catch (e) {
  const msg = e.message || e.body?.error?.message || e.code || JSON.stringify(e);
  console.log(`   ❌ Failed to check settings: ${msg}`);
  if (e.statusCode) console.log(`      Status: ${e.statusCode}`);
  if (e.body?.error) console.log(`      Error detail: ${JSON.stringify(e.body.error)}`);
}

// 5. Check sample items
console.log('\n5. SAMPLE EXTERNAL ITEMS');
console.log('-'.repeat(40));
try {
  const items = await betaClient.api(`/external/connections/${CONNECTION_ID}/items`)
    .top(3)
    .get();

  console.log(`   Total items found: ${items.value?.length || 0}`);

  for (const item of items.value || []) {
    console.log(`\n   Item ID: ${item.id}`);
    console.log(`   Properties:`);

    for (const [key, value] of Object.entries(item.properties || {})) {
      if (key === 'accountInformation') {
        console.log(`      accountInformation: ${value}`);
        // Parse and check format
        try {
          const parsed = JSON.parse(value);
          if (parsed.userPrincipalName) {
            console.log(`         ✅ Has userPrincipalName: ${parsed.userPrincipalName}`);
          } else if (parsed.externalDirectoryObjectId) {
            console.log(`         ✅ Has externalDirectoryObjectId`);
          } else {
            console.log(`         ⚠️  Missing userPrincipalName or externalDirectoryObjectId`);
          }
        } catch {
          console.log(`         ⚠️  Could not parse JSON`);
        }
      } else if (key.endsWith('@odata.type')) {
        console.log(`      ${key}: ${value}`);
      } else if (Array.isArray(value)) {
        console.log(`      ${key}: [${value.length} items]`);
        if (value.length > 0) {
          console.log(`         First: ${value[0].substring(0, 80)}...`);
        }
      } else if (typeof value === 'string' && value.length > 100) {
        console.log(`      ${key}: ${value.substring(0, 100)}...`);
      } else {
        console.log(`      ${key}: ${value}`);
      }
    }

    // Check ACL
    if (item.acl) {
      const everyoneGrant = item.acl.find(a => a.type === 'everyone' && a.accessType === 'grant');
      if (everyoneGrant) {
        console.log(`   ACL: ✅ everyone/grant`);
      } else {
        console.log(`   ACL: ⚠️  ${JSON.stringify(item.acl)}`);
      }
    }
  }
} catch (e) {
  const msg = e.message || e.body?.error?.message || e.code || JSON.stringify(e);
  console.log(`   ❌ Failed to get items: ${msg}`);
  if (e.statusCode) console.log(`      Status: ${e.statusCode}`);
  if (e.body?.error) console.log(`      Error detail: ${JSON.stringify(e.body.error)}`);
  if (e.statusCode === 404) {
    console.log('   (Items endpoint may require different permissions or beta API)');
  }
}

// 6. Summary
console.log('\n' + '='.repeat(60));
console.log('SUMMARY');
console.log('='.repeat(60));
console.log(`
The "Users: -" in Admin Portal typically means one of:
1. The personAccount label is not being recognized
2. The userPrincipalName in accountInformation doesn't match Entra ID users
3. Profile source registration is missing
4. Need more time for Microsoft to process user mapping (can take 24+ hours)

If all checks above pass, wait another 24 hours and check if user profiles
show the connector data.

To verify data appears on profiles:
1. Go to a user's profile card in Teams/Outlook
2. Check if skills/notes from the connector appear
3. Try Copilot queries like "Find people with skills in TypeScript"
`);

function resolveConnectionId() {
  return process.argv[2] || process.env.CONNECTION_ID || process.env.M365_CONNECTION_ID || null;
}
