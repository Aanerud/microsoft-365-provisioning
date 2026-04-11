#!/usr/bin/env node
/**
 * Verify ingestion progress for a people connector.
 * - Counts expected items from CSV (unique emails)
 * - Queries Microsoft Search for externalItem total
 * - Computes percent indexed
 * - Optionally fetches a sample item to confirm properties
 *
 * Requires app-only auth (AZURE_CLIENT_SECRET) and Graph beta endpoints.
 */
import dotenv from 'dotenv';
import fs from 'fs/promises';
import { parse } from 'csv-parse/sync';
import { ClientSecretCredential } from '@azure/identity';
import { Client } from '@microsoft/microsoft-graph-client';
import { TokenCredentialAuthenticationProvider } from '@microsoft/microsoft-graph-client/authProviders/azureTokenCredentials/index.js';

dotenv.config();

const args = process.argv.slice(2);
const showHelp = args.includes('--help') || args.includes('-h');

const connectionId = getArgValue(args, '--connection-id') || process.env.CONNECTION_ID || 'm365people';
const csvPath = getArgValue(args, '--csv') || 'config/textcraft-europe.csv';
const region = getArgValue(args, '--region') || process.env.GRAPH_SEARCH_REGION || 'EMEA';
const queryString = getArgValue(args, '--query') || '*';
const sampleEmail = getArgValue(args, '--sample-email');

if (showHelp) {
  console.log('Usage: node tools/debug/verify-ingestion-progress.mjs [options]');
  console.log('');
  console.log('Options:');
  console.log('  --connection-id <id>    Graph connector id (default: m365people)');
  console.log('  --csv <path>            CSV path (default: config/textcraft-europe.csv)');
  console.log('  --region <region>       Search region for app-only search (default: EMEA; ignored for delegated)');
  console.log('  --query <string>        Search query string (default: *)');
  console.log('  --search-auth <mode>    Search auth: app|delegated|token (default: app)');
  console.log('  --search-token <token>  Delegated token for search (skips browser login)');
  console.log('  --delegated-scopes <s>  Comma-separated delegated scopes (default: ExternalItem.Read.All,offline_access)');
  console.log('  --sample-email <email>  Fetch a sample item and print property info');
  process.exit(0);
}

const searchToken = getArgValue(args, '--search-token') || process.env.GRAPH_SEARCH_TOKEN;
const searchAuth = (getArgValue(args, '--search-auth') || (searchToken ? 'token' : 'app')).toLowerCase();
const delegatedScopes = parseCsvList(getArgValue(args, '--delegated-scopes')) || [
  'ExternalItem.Read.All',
  'offline_access',
];
const supportedSearchAuth = new Set(['app', 'delegated', 'token']);

if (!supportedSearchAuth.has(searchAuth)) {
  console.error('Invalid --search-auth value. Use app, delegated, or token.');
  process.exit(1);
}

if (searchAuth === 'token' && !searchToken) {
  console.error('Missing --search-token (or GRAPH_SEARCH_TOKEN) when using --search-auth token.');
  process.exit(1);
}

const searchAuthLabel = searchAuth === 'app' ? 'app-only' : 'delegated';
const searchRegion = searchAuthLabel === 'app-only' ? region : null;

const { AZURE_TENANT_ID, AZURE_CLIENT_ID, AZURE_CLIENT_SECRET } = process.env;

if (!AZURE_TENANT_ID || !AZURE_CLIENT_ID || !AZURE_CLIENT_SECRET) {
  console.error('Missing required environment variables: AZURE_TENANT_ID, AZURE_CLIENT_ID, AZURE_CLIENT_SECRET');
  process.exit(1);
}

const credential = new ClientSecretCredential(AZURE_TENANT_ID, AZURE_CLIENT_ID, AZURE_CLIENT_SECRET);
const authProvider = new TokenCredentialAuthenticationProvider(credential, {
  scopes: ['https://graph.microsoft.com/.default'],
});
const betaClient = Client.initWithMiddleware({ authProvider, defaultVersion: 'beta' });

let records = [];
try {
  const content = await fs.readFile(csvPath, 'utf-8');
  records = parse(content, { columns: true, skip_empty_lines: true, trim: true });
} catch (error) {
  console.warn(`Warning: Could not read CSV at ${csvPath}. ${error.message}`);
}

const emails = records
  .map((row) => (row.email ? String(row.email).trim() : ''))
  .filter((email) => email.length > 0);
const normalizedEmails = emails.map((email) => email.toLowerCase());
const uniqueEmails = new Set(normalizedEmails);
const expectedCount = uniqueEmails.size;
const duplicateCount = normalizedEmails.length - uniqueEmails.size;
const missingEmailCount = records.length - emails.length;

console.log('INGESTION VERIFICATION');
console.log('----------------------');
console.log(`Connection ID: ${connectionId}`);
console.log(`CSV path: ${csvPath}`);
console.log(`Search auth: ${searchAuthLabel}`);
if (searchRegion) {
  console.log(`Search region: ${searchRegion}`);
} else {
  console.log('Search region: (ignored for delegated auth)');
}
console.log(`Search query: ${queryString}`);
console.log(`CSV rows: ${records.length}`);
console.log(`Unique emails: ${expectedCount}`);
if (missingEmailCount > 0) {
  console.log(`Rows missing email: ${missingEmailCount}`);
}
if (duplicateCount > 0) {
  console.log(`Duplicate emails: ${duplicateCount} (items will overwrite by item ID)`);
}
console.log('');

let indexedTotal = null;
let moreResultsAvailable = null;
let connectionItemCount = null;

try {
  const connection = await betaClient.api(`/external/connections/${connectionId}`).get();
  connectionItemCount = connection.itemCount ?? null;
} catch (error) {
  console.warn(`Warning: Failed to read connection itemCount. ${error.message}`);
}

let searchClient = betaClient;
if (searchAuth !== 'app') {
  const delegatedAccessToken = searchToken || await getDelegatedAccessToken(delegatedScopes);
  searchClient = Client.init({
    authProvider: (done) => done(null, delegatedAccessToken),
    defaultVersion: 'beta',
  });
}

try {
  const searchRequest = {
    entityTypes: ['externalItem'],
    contentSources: [`/external/connections/${connectionId}`],
    query: { queryString },
    from: 0,
    size: 1,
    ...(searchRegion ? { region: searchRegion } : {}),
  };

  const searchResponse = await searchClient.api('/search/query').post({
    requests: [searchRequest],
  });

  const container = searchResponse?.value?.[0]?.hitsContainers?.[0];
  indexedTotal = container?.total ?? null;
  moreResultsAvailable = container?.moreResultsAvailable ?? null;
} catch (error) {
  console.error('Search query failed.');
  const detailMessage = error.body?.error?.message || error.message;
  if (detailMessage) {
    console.error(`  ${detailMessage}`);
  }
  if (error.statusCode) {
    console.error(`  Status: ${error.statusCode}`);
  }
  if (error.statusCode === 403 && `${detailMessage || ''}`.includes('Application permission is only supported')) {
    console.error('  Hint: externalItem search requires delegated auth. Use --search-auth delegated or --search-token.');
  }
}

if (indexedTotal !== null) {
  console.log(`Search index total: ${indexedTotal}`);
  if (moreResultsAvailable !== null) {
    console.log(`More results available: ${moreResultsAvailable}`);
  }
  if (expectedCount > 0) {
    const percent = (indexedTotal / expectedCount) * 100;
    console.log(`Indexing progress: ${percent.toFixed(1)}%`);
    if (indexedTotal > expectedCount) {
      console.log('Note: search total exceeds expected count (old items or prior runs may exist).');
    }
  }
} else {
  console.log('Search index total: unavailable');
}

if (connectionItemCount !== null && (indexedTotal === null || connectionItemCount !== indexedTotal)) {
  console.log(`Connection itemCount (ingested): ${connectionItemCount}`);
  if (indexedTotal === null && expectedCount > 0) {
    const percent = (connectionItemCount / expectedCount) * 100;
    console.log(`Ingestion progress (connection itemCount): ${percent.toFixed(1)}%`);
  }
}

const sampleToCheck = sampleEmail || (emails.length > 0 ? emails[0] : null);
if (sampleToCheck) {
  console.log('');
  console.log('SAMPLE ITEM CHECK');
  console.log('-----------------');
  const itemId = toItemId(sampleToCheck);
  try {
    const item = await betaClient.api(`/external/connections/${connectionId}/items/${itemId}`).get();
    const properties = item.properties || {};
    const propertyKeys = Object.keys(properties);
    const populatedCount = propertyKeys.filter((key) => isPopulated(properties[key])).length;

    console.log(`Sample email: ${sampleToCheck}`);
    console.log(`Item ID: ${itemId}`);
    console.log(`Properties: ${propertyKeys.length} (${populatedCount} populated)`);
    console.log(`Has accountInformation: ${propertyKeys.includes('accountInformation')}`);

    const hasEveryoneAcl = Array.isArray(item.acl)
      ? item.acl.some((entry) => entry?.type === 'everyone' && entry?.value === 'everyone')
      : false;
    console.log(`ACL includes everyone: ${hasEveryoneAcl}`);

    if (propertyKeys.length > 0) {
      const preview = propertyKeys.slice(0, 12).join(', ');
      const remaining = propertyKeys.length - 12;
      console.log(`Property keys: ${preview}${remaining > 0 ? ` ... (+${remaining} more)` : ''}`);
    }
  } catch (error) {
    console.error(`Failed to fetch sample item: ${error.message}`);
    if (error.statusCode) {
      console.error(`Status: ${error.statusCode}`);
    }
  }
} else {
  console.log('');
  console.log('No sample email found. Provide --sample-email to validate a specific item.');
}

function getArgValue(argv, flag) {
  const index = argv.indexOf(flag);
  if (index === -1) return null;
  const value = argv[index + 1];
  if (!value || value.startsWith('--')) return null;
  return value;
}

function parseCsvList(value) {
  if (!value) return null;
  const list = value
    .split(',')
    .map((item) => item.trim())
    .filter(Boolean);
  return list.length > 0 ? list : null;
}

async function getDelegatedAccessToken(scopes) {
  const tenantId = process.env.AZURE_TENANT_ID;
  const clientId = process.env.AZURE_CLIENT_ID;
  const authPort = parseInt(process.env.AUTH_SERVER_PORT || '5544', 10);

  if (!tenantId || !clientId) {
    throw new Error('AZURE_TENANT_ID and AZURE_CLIENT_ID are required for delegated search.');
  }

  try {
    const { BrowserAuthServer } = await import('../../dist/auth/browser-auth-server.js');
    const authServer = new BrowserAuthServer({
      tenantId,
      clientId,
      port: authPort,
      scopes,
    });
    const authResult = await authServer.authenticate();
    return authResult.accessToken;
  } catch (error) {
    if (error.code === 'ERR_MODULE_NOT_FOUND') {
      throw new Error('Delegated search requires build output. Run: npm run build');
    }
    throw error;
  }
}

function toItemId(email) {
  return `person-${email.replace(/@/g, '-').replace(/\./g, '-')}`;
}

function isPopulated(value) {
  if (value === null || value === undefined) return false;
  if (typeof value === 'string') return value.trim().length > 0;
  if (Array.isArray(value)) return value.length > 0;
  return true;
}
