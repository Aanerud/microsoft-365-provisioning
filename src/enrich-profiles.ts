#!/usr/bin/env node

import fs from 'fs/promises';
import { parse } from 'csv-parse/sync';
import dotenv from 'dotenv';
import { GraphClient } from './graph-client.js';
import { BrowserAuthServer } from './auth/browser-auth-server.js';
import { initializeLogger } from './utils/logger.js';
import { PeopleConnectionManager } from './people-connector/connection-manager.js';
import { PeopleSchemaBuilder } from './people-connector/schema-builder.js';
import { PeopleItemIngester } from './people-connector/item-ingester.js';
import { parsePropertyValue } from './schema/user-property-schema.js';

dotenv.config();

interface EnrichOptions {
  csvPath: string;
  connectionId: string;
  createConnection: boolean;
  registerSchema: boolean;
  dryRun: boolean;
}

class ProfileEnrichment {
  private graphClient: GraphClient | null = null;
  private connectionManager: PeopleConnectionManager | null = null;
  private itemIngester: PeopleItemIngester | null = null;
  private tenantId: string;
  private clientId: string;

  constructor() {
    this.tenantId = process.env.AZURE_TENANT_ID || '';
    this.clientId = process.env.AZURE_CLIENT_ID || '';

    if (!this.tenantId || !this.clientId) {
      throw new Error('Missing Azure AD configuration in .env file');
    }
  }

  /**
   * Parse CLI arguments
   */
  parseArgs(): EnrichOptions {
    const args = process.argv.slice(2);
    const options: EnrichOptions = {
      csvPath: 'config/agents-template.csv',
      connectionId: 'm365provisionpeople',
      createConnection: false,
      registerSchema: false,
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
          options.createConnection = true;
          options.registerSchema = true;
          break;
        case '--dry-run':
          options.dryRun = true;
          break;
      }
    }

    return options;
  }

  /**
   * Load CSV and parse Option B properties
   */
  async loadPeopleData(csvPath: string): Promise<any[]> {
    const content = await fs.readFile(csvPath, 'utf-8');
    const records = parse(content, {
      columns: true,
      skip_empty_lines: true,
      trim: true,
    });

    // Parse array values
    return records.map((record: any) => {
      const parsed = { ...record };

      // Parse skills, interests, etc.
      if (record.skills) parsed.skills = parsePropertyValue('skills', record.skills);
      if (record.interests) parsed.interests = parsePropertyValue('interests', record.interests);
      if (record.pastProjects) parsed.pastProjects = parsePropertyValue('pastProjects', record.pastProjects);
      if (record.responsibilities) parsed.responsibilities = parsePropertyValue('responsibilities', record.responsibilities);
      if (record.schools) parsed.schools = parsePropertyValue('schools', record.schools);
      if (record.certifications) parsed.certifications = parsePropertyValue('certifications', record.certifications);
      if (record.awards) parsed.awards = parsePropertyValue('awards', record.awards);

      // Parse languages with proficiency
      // Supports formats:
      // - JSON: [{"language":"Norwegian","proficiency":"native"},{"language":"English","proficiency":"fullProfessional"}]
      // - Simple: ["Norwegian:native","English:fullProfessional"]
      // - Plain: ["Norwegian","English"] (defaults to professionalWorking proficiency)
      if (record.languages) parsed.languages = parsePropertyValue('languages', record.languages);

      return parsed;
    });
  }

  /**
   * Main enrichment process
   */
  async run(): Promise<void> {
    const options = this.parseArgs();

    console.log('🚀 Profile Enrichment (Option B)\n');
    console.log('Configuration:');
    console.log(`  CSV: ${options.csvPath}`);
    console.log(`  Connection ID: ${options.connectionId}`);
    console.log(`  Dry Run: ${options.dryRun ? 'Yes' : 'No'}`);
    console.log('');

    // Check authentication method
    const clientSecret = process.env.AZURE_CLIENT_SECRET;

    if (clientSecret) {
      // Use OAuth 2.0 Client Credentials Flow (app-only)
      console.log('🔐 Using application-only authentication (OAuth 2.0 client credentials)...');
      this.graphClient = new GraphClient({
        tenantId: this.tenantId,
        clientId: this.clientId,
        clientSecret: clientSecret
      });
      console.log('✓ Authenticated with client credentials\n');
    } else {
      // Use OAuth 2.0 Authorization Code Flow (delegated)
      console.log('🔐 Authenticating with browser (OAuth 2.0 delegated)...');
      const authServer = new BrowserAuthServer({
        tenantId: this.tenantId,
        clientId: this.clientId,
        scopes: [
          'User.Read.All',
          'Directory.Read.All',
          'ExternalConnection.ReadWrite.All',
          'ExternalItem.ReadWrite.All'
        ]
      });
      const authResult = await authServer.authenticate();
      this.graphClient = new GraphClient({ accessToken: authResult.accessToken });
      console.log('✓ Authenticated\n');
    }

    // Initialize clients
    const { client, betaClient } = this.graphClient.getClients();
    this.connectionManager = new PeopleConnectionManager(client, betaClient, options.connectionId);

    // Load people data first to get CSV columns
    console.log('📖 Loading people data from CSV...');
    const peopleData = await this.loadPeopleData(options.csvPath);
    console.log(`✓ Loaded ${peopleData.length} people\n`);

    // Get CSV columns for schema building
    const csvColumns = peopleData.length > 0 ? Object.keys(peopleData[0]) : [];

    // Setup connection if requested
    if (options.createConnection || options.registerSchema) {
      console.log('📋 Setting up Graph Connector...\n');

      if (options.createConnection) {
        await this.connectionManager.createConnection(
          'M365 Provision People Data',
          'People data enrichment from M365-Agent-Provisioning'
        );

        // Register as profile source to fix "unknownFutureValue" label issue
        // This requires PeopleSettings.ReadWrite.All application permission
        const userDomain = process.env.USER_DOMAIN || 'yourdomain.onmicrosoft.com';
        const webUrl = `https://${userDomain.replace('.onmicrosoft.com', '.sharepoint.com')}`;
        await this.connectionManager.registerAsProfileSource(
          'M365 Agent Provisioning',
          webUrl
        );
      }

      if (options.registerSchema) {
        const schema = PeopleSchemaBuilder.buildPeopleSchema(csvColumns);
        console.log(`  Schema includes ${schema.length - 1} properties (${schema.filter(p => p.labels).length - 1} with labels, ${schema.filter(p => !p.labels && p.name !== 'accountInformation').length} custom)`);
        await this.connectionManager.registerSchema(schema);
      }
    }

    const logger = await initializeLogger('logs');
    this.itemIngester = new PeopleItemIngester(client, betaClient, options.connectionId, logger);

    // Create external items
    console.log('🔨 Creating external items...');
    const items = peopleData.map(person => this.itemIngester!.createExternalItem(person, csvColumns));

    if (options.dryRun) {
      console.log('\n📄 Dry Run - Sample Item:');
      console.log(JSON.stringify(items[0], null, 2));
      console.log(`\nWould ingest ${items.length} items`);
      await logger.close();
      return;
    }

    // Batch ingest
    console.log(`\n📤 Ingesting ${items.length} items (batch size: 20)...\n`);
    const { successful, failed } = await this.itemIngester.batchIngestItems(items);

    // Cleanup orphaned items (items not in CSV)
    console.log(`\n🔍 Checking for orphaned external items...`);
    const csvEmails = new Set(peopleData.map(p => p.email.toLowerCase()));
    const deletedItems = await this.itemIngester.deleteOrphanedItems(csvEmails);

    if (deletedItems.length > 0) {
      console.log(`🗑️  Deleted ${deletedItems.length} orphaned items`);
      deletedItems.forEach(id => console.log(`  ✓ ${id}`));
    } else {
      console.log(`✓ No orphaned items found`);
    }

    // Summary
    console.log('\n' + '='.repeat(60));
    console.log('📊 Enrichment Summary');
    console.log('='.repeat(60));
    console.log(`✅ Successful: ${successful.length}`);
    console.log(`❌ Failed:     ${failed.length}`);
    console.log(`🗑️  Deleted:    ${deletedItems.length}`);
    console.log('='.repeat(60));

    if (failed.length > 0) {
      console.log('\n❌ Failed Items:');
      failed.forEach(f => console.log(`  - ${f.id}: ${f.error}`));
    }

    await logger.close();
  }
}

// Run
const enrichment = new ProfileEnrichment();
enrichment.run().catch(console.error);
