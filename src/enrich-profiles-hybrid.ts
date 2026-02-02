#!/usr/bin/env node
/**
 * Hybrid Profile Enrichment
 *
 * Uses the optimal approach for each data type to maximize Copilot searchability:
 *
 * 1. Graph Connectors (people data labels only) - Copilot-searchable:
 *    - skills (personSkills label)
 *    - aboutMe/notes (personNote label)
 *    - certifications, awards, projects, birthday, mySite (labels only)
 *    - NOTE: Unlabeled/custom fields are ignored in strict-by-doc mode
 *
 * 2. Profile API (Direct POST) - Profile cards only (not Copilot-searchable):
 *    - languages (no connector label available)
 *    - interests (no connector label available)
 *    - skills, aboutMe (also written here for profile cards)
 *
 * Key insight: Data written via Profile API with delegated auth is stored as
 * source.type: "User" with isSearchable: false. Only data from system sources
 * (connectors with people data labels) is Copilot-searchable.
 *
 * Available people data labels:
 *   - personAccount (required for user mapping)
 *   - personSkills, personNote, personCertifications, personAwards, personProjects
 *   - personAddresses, personEmails, personPhones
 *
 * NOT available: personLanguages, personInterests (platform limitation)
 */

import fs from 'fs/promises';
import { parse } from 'csv-parse/sync';
import dotenv from 'dotenv';
import { GraphClient } from './graph-client.js';
import { BrowserAuthServer } from './auth/browser-auth-server.js';
import { initializeLogger, Logger } from './utils/logger.js';
import { PeopleConnectionManager } from './people-connector/connection-manager.js';
import { PeopleItemIngester } from './people-connector/item-ingester.js';
import { PeopleSchemaBuilder } from './people-connector/schema-builder.js';
import { getOptionBProperties } from './schema/user-property-schema.js';
import { applyOidCacheToRows, ensureOidCacheWithAuth, ensureOidCacheWithClient } from './oid-cache.js';
import {
  ProfileWriter,
  SkillProficiency,
  PersonInterest,
  PersonNote,
  LanguageProficiency,
} from './profile-writer.js';

dotenv.config();

// Profile API fields (handled directly)
const PROFILE_API_FIELDS = ['skills', 'languages', 'aboutMe', 'interests'];

// Standard user fields (handled by Option A provisioning, not enrichment)
const STANDARD_USER_FIELDS = new Set([
  'name', 'email', 'role', 'department', 'givenName', 'surname', 'jobTitle',
  'employeeType', 'companyName', 'officeLocation', 'streetAddress', 'city',
  'state', 'country', 'postalCode', 'usageLocation', 'preferredLanguage',
  'mobilePhone', 'businessPhones', 'employeeId', 'employeeHireDate', 'ManagerEmail'
]);

interface EnrichOptions {
  csvPath: string;
  connectionId: string;
  setupConnector: boolean;
  skipProfileApi: boolean;
  skipConnector: boolean;
  dryRun: boolean;
}

interface UserProfile {
  email: string;
  skills?: string[];
  aboutMe?: string;
  interests?: string[];
  languages?: LanguageProficiency[];
}

function getLabeledOptionBFieldNames(): string[] {
  return getOptionBProperties()
    .filter(prop => prop.peopleDataLabel)
    .map(prop => prop.name);
}

class HybridProfileEnrichment {
  private tenantId: string;
  private clientId: string;
  private logger: Logger | null = null;

  constructor() {
    this.tenantId = process.env.AZURE_TENANT_ID || '';
    this.clientId = process.env.AZURE_CLIENT_ID || '';

    if (!this.tenantId || !this.clientId) {
      throw new Error('Missing Azure AD configuration in .env file');
    }
  }

  parseArgs(): EnrichOptions {
    const args = process.argv.slice(2);
    const options: EnrichOptions = {
      csvPath: 'config/textcraft-europe.csv',
      connectionId: 'm365people3',
      setupConnector: false,
      skipProfileApi: false,
      skipConnector: false,
      dryRun: false,
    };

    for (let i = 0; i < args.length; i++) {
      switch (args[i]) {
        case '--csv':
          options.csvPath = args[++i];
          break;
        case '--connection-id':
          options.connectionId = args[++i];
          break;
        case '--setup':
          options.setupConnector = true;
          break;
        case '--skip-profile-api':
          options.skipProfileApi = true;
          break;
        case '--skip-connector':
          options.skipConnector = true;
          break;
        case '--dry-run':
          options.dryRun = true;
          break;
      }
    }

    return options;
  }

  /**
   * Load CSV and categorize data for Profile API vs Graph Connectors
   */
  async loadProfiles(csvPath: string): Promise<{
    profiles: UserProfile[];
    connectorRows: any[];
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
    const ignoredFields = csvColumns.filter(col =>
      !STANDARD_USER_FIELDS.has(col) &&
      !PROFILE_API_FIELDS.includes(col) &&
      !labeledOptionBFields.has(col)
    );

    const profiles: UserProfile[] = records.map((row: any) => {
      const profile: UserProfile = {
        email: row.email,
      };

      // Parse Profile API fields
      if (row.skills) {
        profile.skills = this.parseArrayValue(row.skills);
      }
      if (row.aboutMe) {
        profile.aboutMe = row.aboutMe;
      }
      if (row.interests) {
        profile.interests = this.parseArrayValue(row.interests);
      }
      if (row.languages) {
        profile.languages = ProfileWriter.parseLanguages(row.languages);
      }

      return profile;
    });

    return { profiles, connectorRows: records, csvColumns, ignoredFields };
  }

  private parseArrayValue(value: string): string[] {
    if (!value || value.trim() === '') return [];

    try {
      const normalized = value.replace(/'/g, '"');
      const parsed = JSON.parse(normalized);
      if (Array.isArray(parsed)) {
        return parsed.map(String);
      }
    } catch {
      // Fall back to comma-separated
    }

    return value.split(',').map(s => s.trim()).filter(Boolean);
  }

  /**
   * Enrich via Profile API (delegated auth)
   */
  async enrichViaProfileApi(
    profiles: UserProfile[],
    accessToken: string,
    dryRun: boolean
  ): Promise<{ successful: number; failed: number }> {
    console.log('\n' + '='.repeat(60));
    console.log('PHASE 1: Profile API (Direct POST)');
    console.log('='.repeat(60));
    console.log('Writing: skills, languages, notes (aboutMe), interests\n');

    const profileWriter = new ProfileWriter(accessToken);
    let successful = 0;
    let failed = 0;

    for (let i = 0; i < profiles.length; i++) {
      const profile = profiles[i];
      const hasProfileData =
        (profile.skills && profile.skills.length > 0) ||
        profile.aboutMe ||
        (profile.interests && profile.interests.length > 0) ||
        (profile.languages && profile.languages.length > 0);

      if (!hasProfileData) continue;

      console.log(`[${i + 1}/${profiles.length}] ${profile.email}`);

      if (dryRun) {
        if (profile.skills?.length) console.log(`  [DRY RUN] Would POST ${profile.skills.length} skills`);
        if (profile.aboutMe) console.log(`  [DRY RUN] Would POST aboutMe`);
        if (profile.interests?.length) console.log(`  [DRY RUN] Would POST ${profile.interests.length} interests`);
        if (profile.languages?.length) console.log(`  [DRY RUN] Would POST ${profile.languages.length} languages`);
        successful++;
        continue;
      }

      let userFailed = false;

      // Skills
      if (profile.skills && profile.skills.length > 0) {
        const skillObjects: SkillProficiency[] = profile.skills.map(s => ({
          displayName: s,
          allowedAudiences: 'organization',
        }));
        const result = await profileWriter.writeSkills(profile.email, skillObjects);
        if (result.failed > 0) userFailed = true;
      }

      // Notes (aboutMe)
      if (profile.aboutMe) {
        const noteObjects: PersonNote[] = [{
          detail: profile.aboutMe,
          displayName: 'About Me',
          allowedAudiences: 'organization',
        }];
        const result = await profileWriter.writeNotes(profile.email, noteObjects);
        if (result.failed > 0) userFailed = true;
      }

      // Interests
      if (profile.interests && profile.interests.length > 0) {
        const interestObjects: PersonInterest[] = profile.interests.map(i => ({
          displayName: i,
          allowedAudiences: 'organization',
        }));
        const result = await profileWriter.writeInterests(profile.email, interestObjects);
        if (result.failed > 0) userFailed = true;
      }

      // Languages
      if (profile.languages && profile.languages.length > 0) {
        const result = await profileWriter.writeLanguages(profile.email, profile.languages);
        if (result.failed > 0) userFailed = true;
      }

      if (userFailed) {
        failed++;
      } else {
        successful++;
      }

      // Small delay between users
      await new Promise(resolve => setTimeout(resolve, 100));
    }

    return { successful, failed };
  }

  /**
   * Enrich via Graph Connectors (app-only auth)
   */
  async enrichViaConnector(
    connectorRows: any[],
    csvColumns: string[],
    ignoredFields: string[],
    options: EnrichOptions,
    graphClient: GraphClient,
    logger: Logger
  ): Promise<{ successful: number; failed: number }> {
    console.log('\n' + '='.repeat(60));
    console.log('PHASE 2: Graph Connectors (Copilot-Searchable Data)');
    console.log('='.repeat(60));
    console.log('Writing people-labeled properties only (strict-by-doc).');
    const labeledOptionBFields = getLabeledOptionBFieldNames();
    const connectorFields = labeledOptionBFields.filter(field => csvColumns.includes(field));
    if (connectorFields.length > 0) {
      console.log(`  - Connector fields: ${connectorFields.join(', ')}`);
    } else {
      console.log('  - Connector fields: (none detected in CSV)');
    }
    if (ignoredFields.length > 0) {
      console.log(`  - Ignored (no people label): ${ignoredFields.join(', ')}`);
    }
    console.log('');

    const { betaClient } = graphClient.getClients();
    const connectionManager = new PeopleConnectionManager(betaClient, options.connectionId);

    // Setup connector if requested
    if (options.setupConnector) {
      console.log('Setting up Graph Connector...');

      await connectionManager.createConnection(
        'M365 People Enrichment',
        'People data enrichment for Copilot search (labeled properties only)'
      );

      // Register schema with people data labels only
      const schema = PeopleSchemaBuilder.buildPeopleSchema();
      console.log(`Schema includes ${schema.length} labeled properties (personAccount + ${schema.length - 1} people labels).`);
      await connectionManager.registerSchema(schema);

      // Register as profile source after schema is ready
      const userDomain = process.env.USER_DOMAIN || 'yourdomain.onmicrosoft.com';
      const webUrl = `https://${userDomain.replace('.onmicrosoft.com', '.sharepoint.com')}`;
      await connectionManager.registerAsProfileSource('M365 Agent Provisioning', webUrl);
    }

    const itemIngester = new PeopleItemIngester(betaClient, options.connectionId, logger);

    // Create external items for profiles with labeled data
    const items = connectorRows
      .filter(row => labeledOptionBFields.some(field => row[field] && row[field] !== ''))
      .map(row => itemIngester.createExternalItem(row));

    if (options.dryRun) {
      console.log('\nDry Run - Sample Item:');
      if (items.length > 0) {
        console.log(JSON.stringify(items[0], null, 2));
      }
      console.log(`\nWould ingest ${items.length} items`);
      return { successful: items.length, failed: 0 };
    }

    // Ingest items
    console.log(`\nIngesting ${items.length} items...\n`);
    const { successful, failed } = await itemIngester.batchIngestItems(items);

    return { successful: successful.length, failed: failed.length };
  }

  async run(): Promise<void> {
    const options = this.parseArgs();

    console.log('Hybrid Profile Enrichment\n');
    console.log('Configuration:');
    console.log(`  CSV: ${options.csvPath}`);
    console.log(`  Connection ID: ${options.connectionId}`);
    console.log(`  Setup Connector: ${options.setupConnector ? 'Yes' : 'No'}`);
    console.log(`  Skip Profile API: ${options.skipProfileApi ? 'Yes' : 'No'}`);
    console.log(`  Skip Connector: ${options.skipConnector ? 'Yes' : 'No'}`);
    console.log(`  Dry Run: ${options.dryRun ? 'Yes' : 'No'}`);
    console.log('');
    console.log('Note: Run Option A provisioning before Option B enrichment.');
    console.log('');

    // Verify CSV exists
    try {
      await fs.access(options.csvPath);
    } catch {
      console.error(`Error: CSV file not found: ${options.csvPath}`);
      process.exit(1);
    }

    // Load profiles
    console.log('Loading profiles from CSV...');
    const { profiles, connectorRows, csvColumns, ignoredFields } = await this.loadProfiles(options.csvPath);
    console.log(`Loaded ${profiles.length} profiles`);
    const labeledOptionBFields = getLabeledOptionBFieldNames();
    const connectorFields = labeledOptionBFields.filter(field => csvColumns.includes(field));
    console.log('Field routing:');
    console.log('  Profile API only (cards): languages, interests');
    console.log(`  Connector (people labels): ${connectorFields.join(', ') || '(none detected)'}`);
    if (ignoredFields.length > 0) {
      console.log(`  Ignored (no people label): ${ignoredFields.join(', ')}`);
    }
    console.log('');

    this.logger = await initializeLogger('logs');

    let profileApiResults = { successful: 0, failed: 0 };
    let connectorResults = { successful: 0, failed: 0 };
    let delegatedAccessToken: string | null = null;

    // Phase 1: Profile API (requires delegated auth)
    if (!options.skipProfileApi) {
      console.log('Authenticating for Profile API (delegated - browser login)...');
      const authServer = new BrowserAuthServer({
        tenantId: this.tenantId,
        clientId: this.clientId,
        port: parseInt(process.env.AUTH_SERVER_PORT || '5544', 10),
        scopes: ['User.ReadWrite.All', 'People.Read.All', 'offline_access'],
      });
      const authResult = await authServer.authenticate();
      console.log('Authenticated\n');
      delegatedAccessToken = authResult.accessToken;

      profileApiResults = await this.enrichViaProfileApi(
        profiles,
        authResult.accessToken,
        options.dryRun
      );
    }

    // Phase 2: Graph Connectors (requires app-only auth)
    // Always run for labeled people data (Copilot-searchable)
    if (!options.skipConnector) {
      let oidCacheResult;
      if (delegatedAccessToken) {
        const delegatedClient = new GraphClient({ accessToken: delegatedAccessToken, useBeta: true });
        oidCacheResult = await ensureOidCacheWithClient({
          csvPath: options.csvPath,
          tenantId: this.tenantId,
          graphClient: delegatedClient,
        });
      } else {
        oidCacheResult = await ensureOidCacheWithAuth({
          csvPath: options.csvPath,
          tenantId: this.tenantId,
          clientId: this.clientId,
          authPort: parseInt(process.env.AUTH_SERVER_PORT || '5544', 10),
        });
      }
      console.log(`🧭 OID cache: ${oidCacheResult.rebuilt ? 'built' : 'loaded'} (${oidCacheResult.cachePath})`);
      const oidSummary = applyOidCacheToRows(connectorRows, oidCacheResult.cache);
      console.log(
        `🔗 OID mapping: ${oidSummary.hits} matched, ${oidSummary.misses} missing, ${oidSummary.existing} prefilled\n`
      );

      const clientSecret = process.env.AZURE_CLIENT_SECRET;
      if (!clientSecret) {
        throw new Error('AZURE_CLIENT_SECRET is required for Graph Connector ingestion (app-only auth).');
      }
      console.log('\nAuthenticating for Graph Connectors (app-only)...');
      const graphClient = new GraphClient({
        tenantId: this.tenantId,
        clientId: this.clientId,
        clientSecret,
      });
      console.log('Authenticated\n');

      connectorResults = await this.enrichViaConnector(
        connectorRows,
        csvColumns,
        ignoredFields,
        options,
        graphClient,
        this.logger
      );
    }

    // Final Summary
    console.log('\n' + '='.repeat(60));
    console.log('ENRICHMENT SUMMARY');
    console.log('='.repeat(60));
    console.log('\nProfile API (languages, interests - profile cards only):');
    console.log(`  Successful: ${profileApiResults.successful}`);
    console.log(`  Failed: ${profileApiResults.failed}`);
    console.log('\nGraph Connectors (people-labeled props only - Copilot-searchable):');
    console.log(`  Successful: ${connectorResults.successful}`);
    console.log(`  Failed: ${connectorResults.failed}`);
    console.log('\nNote: Languages/interests have no people data labels, so they remain');
    console.log('Profile API only (visible on cards, not Copilot-searchable).');
    console.log('='.repeat(60));

    await this.logger.close();

    if (profileApiResults.failed > 0 || connectorResults.failed > 0) {
      process.exit(1);
    }
  }
}

// Run
const enrichment = new HybridProfileEnrichment();
enrichment.run().catch(error => {
  console.error('Fatal error:', error.message);
  process.exit(1);
});
