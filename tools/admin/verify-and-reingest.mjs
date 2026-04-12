#!/usr/bin/env node
/**
 * Verify and Re-ingest Missing Items
 *
 * This script:
 * 1. Loads users from CSV
 * 2. Verifies each user exists in Entra ID
 * 3. Checks which items exist in the Graph Connector
 * 4. Re-ingests missing items
 *
 * Use this when the Admin Portal shows fewer indexed items than expected.
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
const DRY_RUN = process.argv.includes('--dry-run');
const VERBOSE = process.argv.includes('--verbose');

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

console.log('='.repeat(70));
console.log('VERIFY AND RE-INGEST MISSING ITEMS');
console.log('='.repeat(70));
console.log(`Connection ID: ${CONNECTION_ID}`);
console.log(`CSV Path: ${CSV_PATH}`);
console.log(`Dry Run: ${DRY_RUN}`);
console.log();

// Load CSV
const content = await fs.readFile(CSV_PATH, 'utf-8');
const records = parse(content, { columns: true, skip_empty_lines: true, trim: true });
const columns = Object.keys(records[0] || {});

console.log(`Users in CSV: ${records.length}`);
console.log();

// Step 1: Verify users exist in Entra ID (optional - may fail due to permissions)
console.log('1. VERIFYING USERS IN ENTRA ID');
console.log('-'.repeat(50));

const usersInEntra = new Map();
const usersMissing = [];
let skipUserVerification = false;

// Try first user to check if we have permission
try {
  const firstEmail = records[0]?.email;
  if (firstEmail) {
    const user = await betaClient.api(`/users/${firstEmail}`).select('id,userPrincipalName,displayName').get();
    usersInEntra.set(firstEmail, user);
  }
} catch (e) {
  if (e.message?.includes('Insufficient privileges') || e.statusCode === 403) {
    console.log('   ⚠️ No User.Read.All permission - skipping user verification');
    console.log('   Will assume all CSV users exist in Entra ID');
    skipUserVerification = true;
    // Add all records to usersInEntra
    for (const row of records) {
      usersInEntra.set(row.email, { userPrincipalName: row.email });
    }
  } else if (e.statusCode === 404) {
    usersMissing.push(records[0]?.email);
  }
}

if (!skipUserVerification && records.length > 1) {
  // Continue checking remaining users
  for (const row of records.slice(1)) {
    const email = row.email;
    try {
      const user = await betaClient.api(`/users/${email}`).select('id,userPrincipalName,displayName').get();
      usersInEntra.set(email, user);
      if (VERBOSE) console.log(`   ✅ ${email}`);
    } catch (e) {
      if (e.statusCode === 404) {
        usersMissing.push(email);
        if (VERBOSE) console.log(`   ❌ ${email} - NOT FOUND`);
      } else {
        console.log(`   ⚠️ ${email} - Error: ${e.message}`);
      }
    }
  }
}

console.log(`   Users to process: ${usersInEntra.size}`);
if (!skipUserVerification && usersMissing.length > 0) {
  console.log(`   Missing from Entra: ${usersMissing.length}`);
  console.log(`\n   Missing users (will be skipped):`);
  for (const email of usersMissing.slice(0, 5)) {
    console.log(`      - ${email}`);
  }
  if (usersMissing.length > 5) {
    console.log(`      ... and ${usersMissing.length - 5} more`);
  }
}

// Step 2: Check which items exist in Graph Connector
console.log('\n2. CHECKING EXISTING ITEMS');
console.log('-'.repeat(50));

const itemsExisting = new Set();
const itemsMissing = [];

for (const row of records) {
  const email = row.email;
  if (!usersInEntra.has(email)) continue; // Skip users not in Entra

  const itemId = `person-${email.replace(/@/g, '-').replace(/\./g, '-')}`;

  try {
    await betaClient.api(`/external/connections/${CONNECTION_ID}/items/${itemId}`).get();
    itemsExisting.add(itemId);
    if (VERBOSE) console.log(`   ✅ ${email}`);
  } catch (e) {
    if (e.statusCode === 404) {
      itemsMissing.push({ email, itemId, row });
      if (VERBOSE) console.log(`   ❌ ${email}`);
    } else {
      console.log(`   ⚠️ ${email} - Error: ${e.message}`);
    }
  }
}

console.log(`   Existing items: ${itemsExisting.size}`);
console.log(`   Missing items: ${itemsMissing.length}`);

// Step 3: Re-ingest missing items
if (itemsMissing.length === 0) {
  console.log('\n✅ All items already exist!');
  process.exit(0);
}

console.log('\n3. RE-INGESTING MISSING ITEMS');
console.log('-'.repeat(50));

if (DRY_RUN) {
  console.log('   (DRY RUN - no changes will be made)');
  console.log(`   Would re-ingest ${itemsMissing.length} items:`);
  for (const { email } of itemsMissing.slice(0, 10)) {
    console.log(`      - ${email}`);
  }
  if (itemsMissing.length > 10) {
    console.log(`      ... and ${itemsMissing.length - 10} more`);
  }
  process.exit(0);
}

// Get custom fields (columns not in standard set)
const standardFields = new Set([
  'email', 'firstName', 'lastName', 'displayName', 'jobTitle', 'department',
  'officeLocation', 'manager', 'skills', 'aboutMe', 'interests', 'languages'
]);
const customFields = columns.filter(col => !standardFields.has(col));

let successCount = 0;
let failCount = 0;

for (const { email, itemId, row } of itemsMissing) {
  // Build item (simplified - just the essentials)
  const item = buildItem(email, itemId, row, columns);

  try {
    await betaClient.api(`/external/connections/${CONNECTION_ID}/items/${itemId}`).put(item);
    successCount++;
    console.log(`   ✅ ${email}`);
  } catch (e) {
    failCount++;
    console.log(`   ❌ ${email}: ${e.statusCode} - ${e.message?.substring(0, 100)}`);
    if (VERBOSE && e.body) {
      console.log(`      ${JSON.stringify(e.body).substring(0, 200)}`);
    }
  }

  // Small delay to avoid rate limiting
  await new Promise(resolve => setTimeout(resolve, 200));
}

console.log(`\n   Success: ${successCount}`);
console.log(`   Failed: ${failCount}`);

// Summary
console.log('\n' + '='.repeat(70));
console.log('SUMMARY');
console.log('='.repeat(70));
console.log(`
Total users in CSV:        ${records.length}
Users found in Entra ID:   ${usersInEntra.size}
Items already in index:    ${itemsExisting.size}
Items re-ingested:         ${successCount}
Items failed:              ${failCount}

${failCount > 0 ? `
⚠️  Some items failed to ingest. Common causes:
- Invalid data format in skills/languages/aboutMe
- Rate limiting (try again later)
- Schema validation errors

Run with --verbose for more details.
` : ''}
Next steps:
1. Wait 15-30 minutes for Microsoft to process
2. Check Admin Portal for updated item count
3. Run check-profile-source.mjs to verify status
`);

/**
 * Build external item from CSV row
 */
function buildItem(email, itemId, row, csvColumns) {
  const contentParts = [];
  if (row.aboutMe) contentParts.push(row.aboutMe);
  if (row.skills) contentParts.push(`Skills: ${row.skills}`);
  if (row.interests) contentParts.push(`Interests: ${row.interests}`);

  const properties = {
    // REQUIRED: Link to Entra ID user
    accountInformation: JSON.stringify({
      userPrincipalName: email
    })
  };

  // Add skills (with personSkills label format)
  if (row.skills) {
    let skills = parseArray(row.skills);
    properties['skills@odata.type'] = 'Collection(String)';
    properties.skills = skills.map(s => JSON.stringify({ displayName: s }));
  }

  // Add aboutMe (with personNote label format)
  if (row.aboutMe) {
    properties.aboutMe = JSON.stringify({
      detail: { contentType: 'text', content: row.aboutMe }
    });
  }

  // Add languages (with language format)
  if (row.languages) {
    let languages = parseArray(row.languages);
    properties['languages@odata.type'] = 'Collection(String)';
    properties.languages = languages.map(lang => {
      // Parse "Italian (Native)" format
      const match = lang.match(/^(.+?)\s*\(([^)]+)\)$/);
      if (match) {
        const [, langName, prof] = match;
        return JSON.stringify({
          displayName: langName.trim(),
          proficiency: mapProficiency(prof)
        });
      }
      return JSON.stringify({
        displayName: lang.trim(),
        proficiency: 'professionalWorking'
      });
    });
  }

  // Add interests (plain strings)
  if (row.interests) {
    let interests = parseArray(row.interests);
    properties['interests@odata.type'] = 'Collection(String)';
    properties.interests = interests;
  }

  // Add custom fields (VTeam, CostCenter, etc.)
  const standardFields = new Set([
    'email', 'firstName', 'lastName', 'displayName', 'jobTitle', 'department',
    'officeLocation', 'manager', 'skills', 'aboutMe', 'interests', 'languages'
  ]);

  for (const col of csvColumns) {
    if (standardFields.has(col)) continue;
    if (row[col] && row[col].trim()) {
      properties[col] = String(row[col]).trim();
      contentParts.push(`${col}: ${row[col]}`);
    }
  }

  return {
    id: itemId,
    content: {
      value: contentParts.join('. ') || `Profile for ${email}`,
      type: 'text'
    },
    properties,
    acl: [
      {
        type: 'everyone',
        value: 'everyone',
        accessType: 'grant'
      }
    ]
  };
}

/**
 * Parse comma-separated or JSON array
 */
function parseArray(value) {
  if (!value) return [];
  if (typeof value !== 'string') return Array.isArray(value) ? value : [value];

  try {
    const parsed = JSON.parse(value.replace(/'/g, '"'));
    return Array.isArray(parsed) ? parsed : [value];
  } catch {
    return value.split(',').map(v => v.trim()).filter(v => v);
  }
}

/**
 * Map proficiency string to Graph API value
 */
function mapProficiency(prof) {
  const normalized = prof.toLowerCase().trim();
  if (normalized.includes('native') || normalized.includes('bilingual')) return 'nativeOrBilingual';
  if (normalized.includes('fluent') || normalized === 'full') return 'fullProfessional';
  if (normalized.includes('professional') || normalized.includes('working')) return 'professionalWorking';
  if (normalized.includes('conversational') || normalized.includes('limited')) return 'limitedWorking';
  if (normalized.includes('basic') || normalized.includes('elementary')) return 'elementary';
  return 'professionalWorking';
}
