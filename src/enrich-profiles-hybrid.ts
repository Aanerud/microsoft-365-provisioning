#!/usr/bin/env node
/**
 * Hybrid Profile Enrichment
 *
 * Uses the optimal approach for each data type:
 * 1. Profile API (Direct POST) - For standard profile fields:
 *    - skills, languages, notes (aboutMe), interests, certifications, awards, etc.
 * 2. Graph Connectors - For custom organizational properties:
 *    - VTeam, CostCenter, BenefitPlan, BuildingAccess, ProjectCode, etc.
 *
 * This hybrid approach provides:
 * - Reliable direct writes for standard profile data
 * - Copilot-searchable custom properties via Graph Connectors
 */

import fs from 'fs/promises';
import { parse } from 'csv-parse/sync';
import dotenv from 'dotenv';
import { GraphClient } from './graph-client.js';
import { BrowserAuthServer } from './auth/browser-auth-server.js';
import { initializeLogger, Logger } from './utils/logger.js';
import { PeopleConnectionManager } from './people-connector/connection-manager.js';
import { PeopleItemIngester } from './people-connector/item-ingester.js';
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
  customProperties: Record<string, string>;
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
      connectionId: 'm365provisionpeople',
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
    customFields: string[];
  }> {
    const content = await fs.readFile(csvPath, 'utf-8');
    const records = parse(content, {
      columns: true,
      skip_empty_lines: true,
      trim: true,
    });

    const csvColumns = records.length > 0 ? Object.keys(records[0]) : [];

    // Identify custom fields (not standard user fields, not Profile API fields)
    const customFields = csvColumns.filter(col =>
      !STANDARD_USER_FIELDS.has(col) &&
      !PROFILE_API_FIELDS.includes(col)
    );

    const profiles: UserProfile[] = records.map((row: any) => {
      const profile: UserProfile = {
        email: row.email,
        customProperties: {},
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

      // Collect custom properties for Graph Connector
      for (const field of customFields) {
        if (row[field] && row[field] !== '') {
          profile.customProperties[field] = String(row[field]);
        }
      }

      return profile;
    });

    return { profiles, customFields };
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
    profiles: UserProfile[],
    customFields: string[],
    options: EnrichOptions,
    graphClient: GraphClient,
    logger: Logger
  ): Promise<{ successful: number; failed: number }> {
    console.log('\n' + '='.repeat(60));
    console.log('PHASE 2: Graph Connectors (Custom Properties)');
    console.log('='.repeat(60));
    console.log(`Writing: ${customFields.join(', ')}\n`);

    const { client, betaClient } = graphClient.getClients();
    const connectionManager = new PeopleConnectionManager(client, betaClient, options.connectionId);

    // Setup connector if requested
    if (options.setupConnector) {
      console.log('Setting up Graph Connector...');

      await connectionManager.createConnection(
        'M365 Custom Properties',
        'Custom organizational properties for Copilot search'
      );

      // Register as profile source
      const userDomain = process.env.USER_DOMAIN || 'yourdomain.onmicrosoft.com';
      const webUrl = `https://${userDomain.replace('.onmicrosoft.com', '.sharepoint.com')}`;
      await connectionManager.registerAsProfileSource('M365 Agent Provisioning', webUrl);

      // Register schema with only custom properties
      const schema = this.buildCustomPropertiesSchema(customFields);
      console.log(`Schema includes ${schema.length - 1} custom properties`);
      await connectionManager.registerSchema(schema);
    }

    const itemIngester = new PeopleItemIngester(client, betaClient, options.connectionId, logger);

    // Create external items with only custom properties
    const items = profiles
      .filter(p => Object.keys(p.customProperties).length > 0)
      .map(p => this.createCustomPropertyItem(p, customFields));

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

  /**
   * Build schema for custom properties only
   */
  private buildCustomPropertiesSchema(customFields: string[]): any[] {
    const schema: any[] = [
      // Required: Link to Entra ID user
      {
        name: 'accountInformation',
        type: 'String',
        isSearchable: false,
        isRetrievable: true,
        labels: ['userAccountInformation'],
      },
    ];

    // Add custom properties
    for (const field of customFields) {
      schema.push({
        name: field,
        type: 'String',
        isSearchable: true,
        isRetrievable: true,
        isQueryable: true,
      });
    }

    return schema;
  }

  /**
   * Create external item with only custom properties
   */
  private createCustomPropertyItem(profile: UserProfile, customFields: string[]): any {
    const itemId = `person-${profile.email.replace(/@/g, '-').replace(/\./g, '-')}`;

    const properties: any = {
      accountInformation: JSON.stringify({
        userPrincipalName: profile.email,
      }),
    };

    const contentParts: string[] = [];

    for (const field of customFields) {
      const value = profile.customProperties[field];
      if (value) {
        properties[field] = value;
        contentParts.push(`${field}: ${value}`);
      }
    }

    return {
      id: itemId,
      content: {
        value: contentParts.join('. '),
        type: 'text',
      },
      properties,
      acl: [
        {
          type: 'everyone',
          value: 'everyone',
          accessType: 'grant',
        },
      ],
    };
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

    // Verify CSV exists
    try {
      await fs.access(options.csvPath);
    } catch {
      console.error(`Error: CSV file not found: ${options.csvPath}`);
      process.exit(1);
    }

    // Load profiles
    console.log('Loading profiles from CSV...');
    const { profiles, customFields } = await this.loadProfiles(options.csvPath);
    console.log(`Loaded ${profiles.length} profiles`);
    console.log(`Profile API fields: ${PROFILE_API_FIELDS.join(', ')}`);
    console.log(`Custom fields for Connector: ${customFields.join(', ')}`);
    console.log('');

    this.logger = await initializeLogger('logs');

    let profileApiResults = { successful: 0, failed: 0 };
    let connectorResults = { successful: 0, failed: 0 };

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

      profileApiResults = await this.enrichViaProfileApi(
        profiles,
        authResult.accessToken,
        options.dryRun
      );
    }

    // Phase 2: Graph Connectors (requires app-only auth)
    if (!options.skipConnector && customFields.length > 0) {
      const clientSecret = process.env.AZURE_CLIENT_SECRET;
      if (!clientSecret) {
        console.log('\nSkipping Graph Connector: AZURE_CLIENT_SECRET not configured');
      } else {
        console.log('\nAuthenticating for Graph Connectors (app-only)...');
        const graphClient = new GraphClient({
          tenantId: this.tenantId,
          clientId: this.clientId,
          clientSecret,
        });
        console.log('Authenticated\n');

        connectorResults = await this.enrichViaConnector(
          profiles,
          customFields,
          options,
          graphClient,
          this.logger
        );
      }
    }

    // Final Summary
    console.log('\n' + '='.repeat(60));
    console.log('ENRICHMENT SUMMARY');
    console.log('='.repeat(60));
    console.log('\nProfile API (skills, languages, notes, interests):');
    console.log(`  Successful: ${profileApiResults.successful}`);
    console.log(`  Failed: ${profileApiResults.failed}`);
    console.log('\nGraph Connectors (custom properties):');
    console.log(`  Successful: ${connectorResults.successful}`);
    console.log(`  Failed: ${connectorResults.failed}`);
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
