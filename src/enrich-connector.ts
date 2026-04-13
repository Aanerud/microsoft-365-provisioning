#!/usr/bin/env node
/**
 * Option B: Graph Connector Enrichment
 *
 * Ingests people-labeled data into Microsoft Graph Connectors for Copilot searchability.
 * Uses application-only authentication (client secret).
 *
 * Handles (properties with people data labels):
 *   - skills (personSkills label)
 *   - aboutMe/notes (personNote label)
 *   - certifications (personCertifications label)
 *   - awards (personAwards label)
 *   - projects (personProjects label)
 *   - birthday (personAnniversaries label)
 *   - mySite (personWebSite label)
 *
 * Key insight: Only data from system sources (connectors with people data labels)
 * is Copilot-searchable. Profile API data is NOT Copilot-searchable.
 */

import fs from 'fs/promises';
import path from 'path';
import { parse } from 'csv-parse/sync';
import dotenv from 'dotenv';
import { GraphClient } from './graph-client.js';
import { initializeLogger, Logger } from './utils/logger.js';
import { PeopleConnectionManager } from './people-connector/connection-manager.js';
import { PeopleItemIngester } from './people-connector/item-ingester.js';
import { PeopleSchemaBuilder } from './people-connector/schema-builder.js';
import { getOptionBProperties } from './schema/user-property-schema.js';
import { applyOidCacheToRows, ensureOidCacheWithAuth, loadOidCache, getOidCachePath } from './oid-cache.js';
import { loadRowsFromJson, validateJsonInput } from './json-loader.js';

dotenv.config();

// Standard user fields (handled by Option A provisioning, not enrichment)
const STANDARD_USER_FIELDS = new Set([
  'name', 'email', 'role', 'department', 'givenName', 'surname', 'jobTitle',
  'employeeType', 'companyName', 'officeLocation', 'streetAddress', 'city',
  'state', 'country', 'postalCode', 'usageLocation', 'preferredLanguage',
  'mobilePhone', 'businessPhones', 'employeeId', 'employeeHireDate', 'ManagerEmail',
  // Used by composite labels (personEmails, personName) or Option A only:
  'mail', 'displayName',
  // Handled by Option A license assignment, not connector:
  'licenses',
  // Used by personCurrentPosition (positions array with relatedPerson):
  'positions',
]);

// Profile API only — no people data label, can't go through connector
const PROFILE_API_FIELDS = new Set<string>([]);

interface ConnectorOptions {
  csvPath?: string;
  jsonPath?: string;
  connectionId: string;
  setupConnector: boolean;
  dryRun: boolean;
}

function getLabeledOptionBFieldNames(): string[] {
  return getOptionBProperties()
    .filter(prop => prop.peopleDataLabel)
    .map(prop => prop.name);
}

function parseArgs(): ConnectorOptions {
  const args = process.argv.slice(2);
  const options: ConnectorOptions = {
    connectionId: 'm365people3',
    setupConnector: false,
    dryRun: false,
  };

  for (let i = 0; i < args.length; i++) {
    switch (args[i]) {
      case '--csv':
        options.csvPath = args[++i];
        break;
      case '--json':
        options.jsonPath = args[++i];
        break;
      case '--connection-id':
        options.connectionId = args[++i];
        break;
      case '--setup':
        options.setupConnector = true;
        break;
      case '--dry-run':
        options.dryRun = true;
        break;
    }
  }

  // Default to CSV if neither specified
  if (!options.csvPath && !options.jsonPath) {
    options.csvPath = 'config/textcraft-europe.csv';
  }

  return options;
}

async function loadConnectorRows(csvPath: string): Promise<{
  rows: any[];
  csvColumns: string[];
  ignoredFields: string[];
}> {
  const content = await fs.readFile(csvPath, 'utf-8');
  const records = parse(content, {
    columns: true,
    skip_empty_lines: true,
    trim: true,
  });

  const csvColumns = records.length > 0 ? Object.keys(records[0]) : [];
  const labeledOptionBFields = new Set(getLabeledOptionBFieldNames());
  const customPropertyNames = new Set(PeopleSchemaBuilder.getCustomPropertyNames(csvColumns));
  const ignoredFields = csvColumns.filter(col =>
    !STANDARD_USER_FIELDS.has(col) &&
    !PROFILE_API_FIELDS.has(col) &&
    !labeledOptionBFields.has(col) &&
    !customPropertyNames.has(col)
  );

  return { rows: records, csvColumns, ignoredFields };
}

async function loadConnectorRowsFromJson(jsonPath: string): Promise<{
  rows: any[];
  allColumns: string[];
  ignoredFields: string[];
}> {
  const records = await loadRowsFromJson(jsonPath);

  // Validate
  const validation = validateJsonInput(records);
  if (!validation.valid) {
    throw new Error(`JSON validation failed:\n  ${validation.errors.join('\n  ')}`);
  }
  if (validation.warnings.length > 0) {
    console.log(`JSON warnings:\n  ${validation.warnings.join('\n  ')}`);
  }

  // Collect all unique keys across all records
  const allColumns = [...new Set(records.flatMap(r => Object.keys(r)))];

  // Field routing
  const labeledOptionBFields = new Set(getLabeledOptionBFieldNames());
  const customPropertyNames = new Set(PeopleSchemaBuilder.getCustomPropertyNames(allColumns));
  const ignoredFields = allColumns.filter(col =>
    col !== 'email' && col !== 'externalDirectoryObjectId' &&
    !STANDARD_USER_FIELDS.has(col) &&
    !PROFILE_API_FIELDS.has(col) &&
    !labeledOptionBFields.has(col) &&
    !customPropertyNames.has(col)
  );

  return { rows: records, allColumns, ignoredFields };
}

async function mergeInputSources(csvPath: string, jsonPath: string): Promise<{
  rows: any[];
  allColumns: string[];
  ignoredFields: string[];
}> {
  const csv = await loadConnectorRows(csvPath);
  const json = await loadConnectorRowsFromJson(jsonPath);

  // Index JSON records by email (lowercase)
  const jsonByEmail = new Map(json.rows.map(r => [r.email.toLowerCase(), r]));

  // Merge: CSV provides base, JSON overrides per-property
  const merged = csv.rows.map(csvRow => {
    const jsonRow = jsonByEmail.get(csvRow.email?.toLowerCase());
    if (!jsonRow) return csvRow;
    return { ...csvRow, ...jsonRow };
  });

  // Add JSON-only records (not in CSV)
  const csvEmails = new Set(csv.rows.map((r: any) => r.email?.toLowerCase()));
  for (const jsonRow of json.rows) {
    if (!csvEmails.has(jsonRow.email.toLowerCase())) {
      merged.push(jsonRow);
    }
  }

  const allColumns = [...new Set([...csv.csvColumns, ...json.allColumns])];
  const ignoredFields = [...new Set([...csv.ignoredFields, ...json.ignoredFields])];

  return { rows: merged, allColumns, ignoredFields };
}

async function enrichViaConnector(
  connectorRows: any[],
  csvColumns: string[],
  ignoredFields: string[],
  options: ConnectorOptions,
  graphClient: GraphClient,
  logger: Logger
): Promise<{ successful: number; failed: number }> {
  console.log('\n' + '='.repeat(60));
  console.log('Option B: Graph Connector Enrichment (Copilot-Searchable)');
  console.log('='.repeat(60));
  console.log('Writing people-labeled + custom searchable properties.');
  const labeledOptionBFields = getLabeledOptionBFieldNames();
  const connectorFields = labeledOptionBFields.filter(field => csvColumns.includes(field));
  const customPropertyNames = PeopleSchemaBuilder.getCustomPropertyNames(csvColumns);
  const customFields = customPropertyNames.filter(field => csvColumns.includes(field));
  if (connectorFields.length > 0) {
    console.log(`  - Labeled fields: ${connectorFields.join(', ')}`);
  }
  if (customFields.length > 0) {
    console.log(`  - Custom fields: ${customFields.join(', ')}`);
  }
  if (connectorFields.length === 0 && customFields.length === 0) {
    console.log('  - Connector fields: (none detected in CSV)');
  }
  if (ignoredFields.length > 0) {
    console.log(`  - Ignored: ${ignoredFields.join(', ')}`);
  }
  console.log('');

  const { betaClient } = graphClient.getClients();
  const connectionManager = new PeopleConnectionManager(betaClient, options.connectionId);

  // Setup connector if requested
  if (options.setupConnector) {
    console.log('Setting up Graph Connector...');

    await connectionManager.createConnection(
      'M365 Provision People Data',
      'People data enrichment for Copilot search (labeled + custom properties)'
    );

    // Register schema with people data labels + custom properties
    const schema = PeopleSchemaBuilder.buildPeopleSchema(csvColumns);
    const labeledCount = schema.filter((p: any) => p.labels).length;
    const customCount = schema.length - labeledCount;
    console.log(`Schema: ${labeledCount} labeled properties + ${customCount} custom searchable properties.`);
    await connectionManager.registerSchema(schema);

    // Verify schema labels were registered correctly
    await connectionManager.verifySchemaLabels();

    // Enable Copilot/Search visibility (idempotent — safe on existing connections)
    await connectionManager.enableSearchExperience();

    // Register as profile source after schema is ready
    const userDomain = process.env.USER_DOMAIN || 'yourdomain.onmicrosoft.com';
    const webUrl = `https://${userDomain.replace('.onmicrosoft.com', '.sharepoint.com')}`;
    await connectionManager.registerAsProfileSource('M365 Agent Provisioning', webUrl);
  }

  const itemIngester = new PeopleItemIngester(betaClient, options.connectionId, logger, csvColumns);

  // Create external items for profiles with labeled or custom property data
  const allConnectorFields = [...labeledOptionBFields, ...customPropertyNames];
  const items = connectorRows
    .filter(row => allConnectorFields.some(field => row[field] && row[field] !== ''))
    .map(row => itemIngester.createExternalItem(row));

  const dumpItemEmail = process.env.DUMP_CONNECTOR_ITEM_EMAIL?.toLowerCase();
  const dumpItems = dumpItemEmail
    ? items.filter(item => {
      const accountInfo = item?.properties?.accountInformation;
      if (!accountInfo) return false;
      try {
        const parsed = JSON.parse(accountInfo);
        return String(parsed.userPrincipalName || '').toLowerCase() === dumpItemEmail;
      } catch {
        return false;
      }
    })
    : items;

  if (process.env.DUMP_CONNECTOR_ITEMS === 'true') {
    await fs.mkdir('debug', { recursive: true });
    const suffix = dumpItemEmail ? `-${dumpItemEmail.replace(/[@.]/g, '-')}` : '';
    const dumpPath = path.join('debug', `external-items-${options.connectionId}${suffix}.json`);
    await fs.writeFile(dumpPath, JSON.stringify(dumpItems, null, 2), 'utf-8');
    if (dumpItemEmail && dumpItems.length === 0) {
      console.log(`No connector items matched ${dumpItemEmail}`);
    }
    console.log(`Wrote connector payloads to ${dumpPath}`);
  }

  if (options.dryRun) {
    console.log('\nDry Run - Sample Item:');
    if (dumpItems.length > 0) {
      console.log(JSON.stringify(dumpItems[0], null, 2));
    } else if (dumpItemEmail) {
      console.log(`No connector items matched ${dumpItemEmail}`);
    }
    console.log(`\nWould ingest ${items.length} items`);
    return { successful: items.length, failed: 0 };
  }

  // Ingest items
  console.log(`\nIngesting ${items.length} items...\n`);
  const { successful, failed } = await itemIngester.batchIngestItems(items);

  return { successful: successful.length, failed: failed.length };
}

async function run(): Promise<void> {
  const options = parseArgs();

  const tenantId = process.env.AZURE_TENANT_ID || '';
  const clientId = process.env.AZURE_CLIENT_ID || '';

  if (!tenantId || !clientId) {
    throw new Error('Missing Azure AD configuration in .env file');
  }

  console.log('Option B: Graph Connector Enrichment\n');
  console.log('Configuration:');
  if (options.csvPath) console.log(`  CSV: ${options.csvPath}`);
  if (options.jsonPath) console.log(`  JSON: ${options.jsonPath}`);
  console.log(`  Connection ID: ${options.connectionId}`);
  console.log(`  Setup Connector: ${options.setupConnector ? 'Yes' : 'No'}`);
  console.log(`  Dry Run: ${options.dryRun ? 'Yes' : 'No'}`);
  console.log('');
  console.log('Note: Run Option A provisioning before Option B enrichment.');
  console.log('');

  // Verify input files exist
  if (options.csvPath) {
    try {
      await fs.access(options.csvPath);
    } catch {
      console.error(`Error: CSV file not found: ${options.csvPath}`);
      process.exit(1);
    }
  }
  if (options.jsonPath) {
    try {
      await fs.access(options.jsonPath);
    } catch {
      console.error(`Error: JSON file not found: ${options.jsonPath}`);
      process.exit(1);
    }
  }

  // Load rows from CSV, JSON, or both
  let connectorRows: any[];
  let csvColumns: string[];
  let ignoredFields: string[];

  if (options.csvPath && options.jsonPath) {
    console.log('Loading profiles from CSV + JSON (merge mode)...');
    const merged = await mergeInputSources(options.csvPath, options.jsonPath);
    connectorRows = merged.rows;
    csvColumns = merged.allColumns;
    ignoredFields = merged.ignoredFields;
  } else if (options.jsonPath) {
    console.log('Loading profiles from JSON...');
    const json = await loadConnectorRowsFromJson(options.jsonPath);
    connectorRows = json.rows;
    csvColumns = json.allColumns;
    ignoredFields = json.ignoredFields;
  } else {
    console.log('Loading profiles from CSV...');
    const csv = await loadConnectorRows(options.csvPath!);
    connectorRows = csv.rows;
    csvColumns = csv.csvColumns;
    ignoredFields = csv.ignoredFields;
  }
  console.log(`Loaded ${connectorRows.length} profiles`);
  const labeledOptionBFields = getLabeledOptionBFieldNames();
  const connectorFields = labeledOptionBFields.filter(field => csvColumns.includes(field));
  const customPropertyNames = PeopleSchemaBuilder.getCustomPropertyNames(csvColumns);
  const customFields = customPropertyNames.filter(field => csvColumns.includes(field));
  console.log('Field routing:');
  console.log(`  Labeled: ${connectorFields.join(', ') || '(none detected)'}`);
  if (customFields.length > 0) {
    console.log(`  Custom: ${customFields.join(', ')}`);
  }
  if (ignoredFields.length > 0) {
    console.log(`  Ignored: ${ignoredFields.join(', ')}`);
  }
  console.log('');

  const logger = await initializeLogger('logs');

  // Load OID cache from disk (built by Option A provisioning or npm run build-oid-cache)
  // Falls back to browser login if cache doesn't exist
  const inputPath = options.csvPath || options.jsonPath!;
  const cachePath = getOidCachePath(inputPath);
  let oidCacheResult;
  const existingCache = await loadOidCache(cachePath);
  if (existingCache) {
    oidCacheResult = { cache: existingCache, cachePath, rebuilt: false };
  } else {
    console.log('OID cache not found on disk. Will authenticate to build it...');
    oidCacheResult = await ensureOidCacheWithAuth({
      csvPath: inputPath,
      tenantId,
      clientId,
      authPort: parseInt(process.env.AUTH_SERVER_PORT || '5544', 10),
    });
  }
  console.log(`OID cache: ${oidCacheResult.rebuilt ? 'built' : 'loaded'} (${oidCacheResult.cachePath})`);
  const oidSummary = applyOidCacheToRows(connectorRows, oidCacheResult.cache);
  console.log(
    `OID mapping: ${oidSummary.hits} matched, ${oidSummary.misses} missing, ${oidSummary.existing} prefilled\n`
  );

  // Gate: abort if too many OID misses (items without OID can't map to users)
  const totalRows = connectorRows.length;
  const matchedRows = oidSummary.hits + oidSummary.existing;
  if (matchedRows === 0 && totalRows > 0) {
    throw new Error(
      `OID mapping failed: 0 of ${totalRows} users matched. ` +
      `Items without externalDirectoryObjectId cannot be mapped to Entra ID profiles.\n` +
      `  Fix: run 'npm run build-oid-cache -- --csv ${inputPath}' or run Option A first.`
    );
  }
  if (oidSummary.misses > 0) {
    const pct = Math.round((oidSummary.misses / totalRows) * 100);
    console.warn(`⚠ ${oidSummary.misses} of ${totalRows} users (${pct}%) have no OID — those items will lack user mapping.`);
    if (pct > 50 && !options.dryRun) {
      throw new Error(
        `Too many OID misses: ${oidSummary.misses} of ${totalRows} (${pct}%). ` +
        `Rebuild the cache: npm run build-oid-cache -- --csv ${inputPath}`
      );
    }
  }

  // Authenticate with app-only auth (client secret)
  const clientSecret = process.env.AZURE_CLIENT_SECRET;
  if (!clientSecret) {
    throw new Error('AZURE_CLIENT_SECRET is required for Graph Connector ingestion (app-only auth).');
  }
  console.log('Authenticating for Graph Connectors (app-only)...');
  const graphClient = new GraphClient({
    tenantId,
    clientId,
    clientSecret,
  });
  console.log('Authenticated\n');

  const connectorResults = await enrichViaConnector(
    connectorRows,
    csvColumns,
    ignoredFields,
    options,
    graphClient,
    logger
  );

  // Summary
  console.log('\n' + '='.repeat(60));
  console.log('ENRICHMENT SUMMARY');
  console.log('='.repeat(60));
  console.log('\nGraph Connectors (people-labeled props only - Copilot-searchable):');
  console.log(`  Successful: ${connectorResults.successful}`);
  console.log(`  Failed: ${connectorResults.failed}`);
  console.log('='.repeat(60));

  await logger.close();

  if (connectorResults.failed > 0) {
    process.exit(1);
  }
}

run().catch(error => {
  console.error('Fatal error:', error.message);
  process.exit(1);
});
