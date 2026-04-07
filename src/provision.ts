#!/usr/bin/env node

import fs from 'fs/promises';
import { parse } from 'csv-parse/sync';
import dotenv from 'dotenv';
import { fileURLToPath } from 'url';
import path from 'path';
import { GraphClient } from './graph-client.js';
import { ConfigExporter, type AgentConfig } from './export.js';
import { BrowserAuthServer } from './auth/browser-auth-server.js';
import { StateManager } from './state-manager.js';
import { initializeLogger } from './utils/logger.js';
import { ensureOidCacheWithClient } from './oid-cache.js';
import {
  getOptionBProperties,
  getCustomProperties,
} from './schema/user-property-schema.js';
import { loadRowsFromJson } from './json-loader.js';
import { buildLicenseResolver } from './license-resolver.js';

dotenv.config();

interface AgentDefinition {
  name: string;
  email: string;
  role: string;
  department: string;
  // Beta-only fields (optional)
  employeeType?: string;
  companyName?: string;
  officeLocation?: string;
}

interface ProvisionOptions {
  csvPath: string;
  outputPath: string;
  dryRun: boolean;
  skipLicenses: boolean;
  force: boolean;
  auth: boolean;
  // New state management flags
  skipDelete: boolean;    // Don't delete users not in CSV
  skipUpdate: boolean;    // Only create new users
  skipCreate: boolean;    // Only update/delete existing
  showDiff: boolean;      // Show detailed diff even in non-dry-run
}

class AgentProvisioner {
  private graphClient: GraphClient;
  private exporter: ConfigExporter;
  private licenseSkuIds: string[];
  private perUserLicenseResolver: ((userLicenses: string[]) => string[]) | null = null;

  constructor(graphClient: GraphClient) {
    this.graphClient = graphClient;
    this.exporter = new ConfigExporter();

    // Support both LICENSE_SKU_IDS (comma-separated) and legacy LICENSE_SKU_ID
    const skuIdsEnv = process.env.LICENSE_SKU_IDS || process.env.LICENSE_SKU_ID || '';
    this.licenseSkuIds = skuIdsEnv
      .split(',')
      .map(id => id.trim())
      .filter(id => id.length > 0);

    if (this.licenseSkuIds.length === 0) {
      console.warn('⚠ WARNING: LICENSE_SKU_IDS not set in .env - will check for per-user licenses in JSON');
    } else {
      console.log(`📋 Configured ${this.licenseSkuIds.length} license(s) from .env`);
    }
  }

  /**
   * Load agent definitions from CSV or JSON file
   */
  async loadAgentsFromFile(filePath: string): Promise<AgentDefinition[]> {
    try {
      let records: any[];

      if (filePath.endsWith('.json')) {
        records = await loadRowsFromJson(filePath);
        console.log(`✓ Loaded ${records.length} agent definitions from JSON ${filePath}\n`);
      } else {
        const content = await fs.readFile(filePath, 'utf-8');
        records = parse(content, {
          columns: true,
          skip_empty_lines: true,
          trim: true,
        });
        console.log(`✓ Loaded ${records.length} agent definitions from ${filePath}\n`);
      }

      // Validate required columns
      for (let i = 0; i < records.length; i++) {
        const record = records[i];
        if (!record.name || !record.email || !record.role || !record.department) {
          throw new Error(
            `Row ${i + 2} missing required columns (name, email, role, department)`
          );
        }
      }

      // Analyze CSV columns for deferred properties
      const csvColumns = records.length > 0 ? Object.keys(records[0]) : [];
      const optionBProps = getOptionBProperties().map(p => p.name);
      const customProps = getCustomProperties(csvColumns);

      // Find deferred properties (Option B standard + custom)
      const deferredStandard = csvColumns.filter(col => optionBProps.includes(col));
      const deferredCustom = customProps;
      const allDeferred = [...deferredStandard, ...deferredCustom];

      if (allDeferred.length > 0) {
        console.log('');
        console.log('📋 Enrichment Properties Detected (Option B):');
        console.log('   These properties will be deferred and handled by Option B:');
        console.log('');

        if (deferredStandard.length > 0) {
          console.log('   Standard enrichment properties:');
          deferredStandard.forEach(prop => {
            const metadata = getOptionBProperties().find(p => p.name === prop);
            const label = metadata?.peopleDataLabel || '(custom - no label)';
            console.log(`     - ${prop} → ${label}`);
          });
          console.log('');
        }

        if (deferredCustom.length > 0) {
          console.log('   Custom organization properties:');
          deferredCustom.forEach(prop => {
            console.log(`     - ${prop} → (searchable custom property)`);
          });
          console.log('');
        }

        console.log('   To enrich profiles with these properties, run:');
        console.log('   npm run enrich-profiles');
        console.log('');
      }

      return records;
    } catch (error: any) {
      throw new Error(`Failed to load CSV: ${error.message}`);
    }
  }

  /**
   * Provision a single agent
   */
  async provisionAgent(
    agent: AgentDefinition,
    options: Partial<ProvisionOptions>
  ): Promise<AgentConfig> {
    console.log(`\n📝 Provisioning: ${agent.name} (${agent.email})`);

    // Step 1: Create user (with beta fields if provided)
    const password = this.generatePassword();
    const user = await this.graphClient.createUser({
      displayName: agent.name,
      email: agent.email,
      password,
      employeeType: agent.employeeType,
      companyName: agent.companyName,
      officeLocation: agent.officeLocation,
    });

    // Step 2: Assign licenses (per-user from JSON, or global from .env)
    if (!options.skipLicenses) {
      let skuIds = this.licenseSkuIds; // default from .env
      if (this.perUserLicenseResolver && Array.isArray((agent as any).licenses)) {
        const resolved = this.perUserLicenseResolver((agent as any).licenses);
        if (resolved.length > 0) skuIds = resolved;
      }
      if (skuIds.length > 0) {
        await this.graphClient.assignLicenses(user.id, skuIds);
      }
    }

    const agentConfig: AgentConfig = {
      name: agent.name,
      email: agent.email,
      role: agent.role,
      department: agent.department,
      userId: user.id,
      password,
      createdAt: new Date().toISOString(),
      // Include beta fields if present
      ...(agent.employeeType && { employeeType: agent.employeeType }),
      ...(agent.companyName && { companyName: agent.companyName }),
      ...(agent.officeLocation && { officeLocation: agent.officeLocation }),
    };

    console.log(`✅ Completed: ${agent.name}`);

    return agentConfig;
  }

  /**
   * Provision all agents from CSV using state management
   * Handles CREATE, UPDATE, and DELETE operations
   */
  async provisionAll(options: ProvisionOptions): Promise<void> {
    console.log('🚀 Starting Agent Provisioning (State Management Mode)\n');
    // Determine input file path (JSON takes precedence if both provided)
    const inputPath = (options as any).jsonPath || options.csvPath;

    console.log('Configuration:');
    console.log(`  Input: ${inputPath}`);
    console.log(`  Output: ${options.outputPath}`);
    console.log(`  Mode: ${options.dryRun ? 'DRY RUN' : 'LIVE'}`);
    console.log(`  Skip Licenses: ${options.skipLicenses}`);
    console.log(`  Skip Delete: ${options.skipDelete}`);
    console.log(`  Skip Update: ${options.skipUpdate}`);
    console.log(`  Skip Create: ${options.skipCreate}`);
    console.log('  Beta Features: ✓ Enabled (always)');
    console.log('');

    // Load agent definitions from CSV or JSON
    const agents = await this.loadAgentsFromFile(inputPath);

    // Resolve per-user licenses from JSON (if present)
    const hasPerUserLicenses = agents.some((a: any) => Array.isArray(a.licenses));
    if (hasPerUserLicenses && !options.skipLicenses) {
      this.perUserLicenseResolver = await buildLicenseResolver(this.graphClient, agents);
    }

    // Get column names for custom property detection
    const csvColumns = agents.length > 0 ? Object.keys(agents[0]) : [];

    // Initialize state management system
    const stateManager = new StateManager({
      graphClient: this.graphClient,
    });

    // Fetch current Azure AD state (including custom properties)
    const azureAdUsers = await stateManager.fetchAzureAdState();

    // Calculate delta (what needs to be created, updated, deleted)
    // This also applies account protection filters
    const delta = await stateManager.calculateDelta(agents, csvColumns, azureAdUsers);

    // Show diff report
    if (options.dryRun || options.showDiff) {
      console.log('\n' + stateManager.generateDiffReport(delta));
    } else {
      console.log(`\n📊 ${stateManager.generateSummary(delta)}\n`);
    }

    // If dry-run, stop here
    if (options.dryRun) {
      console.log('\n🔍 DRY RUN - No changes were made\n');
      return;
    }

    // Confirmation prompt for DELETE operations
    if (delta.delete.length > 0 && !options.force && !options.skipDelete) {
      console.log(`\n⚠️  WARNING: ${delta.delete.length} users will be DELETED!`);
      console.log('Users to delete:');
      delta.delete.slice(0, 10).forEach(action => {
        console.log(`  - ${action.user.displayName} (${action.user.email})`);
      });
      if (delta.delete.length > 10) {
        console.log(`  ... and ${delta.delete.length - 10} more`);
      }
      console.log('\nRun with --force flag to proceed with deletion.');
      console.log('Run with --skip-delete to only CREATE and UPDATE users.');
      return;
    }

    // Apply changes
    console.log('\n📦 Applying changes...\n');

    // Initialize logger
    const logger = await initializeLogger();

    const provisionedAgents: AgentConfig[] = [];
    let createCount = 0;
    let updateCount = 0;
    let deleteCount = 0;

    // 1. Handle CREATE operations
    if (!options.skipCreate && delta.create.length > 0) {
      console.log(`📝 Creating ${delta.create.length} users...`);

      // Prepare users for batch creation
      const usersToCreate = delta.create.map(action => ({
        ...action.user, // Include all standard properties
        password: this.generatePassword(),
      }));

      // Create users in batch
      const { successful, failed } = await this.graphClient.createUsersBatch(usersToCreate);
      createCount = successful.length;

      // Assign licenses (per-user from JSON, or global from .env)
      if (!options.skipLicenses) {
        const hasAnyLicenses = this.licenseSkuIds.length > 0 || this.perUserLicenseResolver;
        if (hasAnyLicenses) {
          console.log(`  Assigning licenses to each user...`);
        }
        for (const user of successful) {
          // Find the agent definition to check for per-user licenses
          const agent = agents.find((a: any) => a.email === user.userPrincipalName);
          let skuIds = this.licenseSkuIds;
          if (this.perUserLicenseResolver && agent && Array.isArray((agent as any).licenses)) {
            const resolved = this.perUserLicenseResolver((agent as any).licenses);
            if (resolved.length > 0) skuIds = resolved;
          }
          if (skuIds.length === 0) continue;
          try {
            const result = await this.graphClient.assignLicenses(user.id, skuIds);
            if (result.failed.length > 0) {
              for (const failure of result.failed) {
                logger.warn(`Failed to assign license ${failure.skuId} to ${user.displayName}`, {
                  userId: user.id,
                  skuId: failure.skuId,
                  error: failure.error,
                });
                console.warn(`⚠ Failed to assign license ${failure.skuId} to ${user.displayName}: ${failure.error}`);
              }
            }
          } catch (error: any) {
            logger.warn(`Failed to assign licenses to ${user.displayName}`, {
              userId: user.id,
              error: error.message,
            });
            console.warn(`⚠ Failed to assign licenses to ${user.displayName}: ${error.message}`);
          }
        }
      }

      // Assign managers (must be done after all users are created)
      console.log(`  Checking for manager assignments...`);
      const managerAssignments: Array<{ userId: string; managerId: string; userName: string }> = [];

      for (const user of successful) {
        const csvUser = agents.find(a => a.email === user.userPrincipalName);
        const managerEmail = (csvUser as any)?.ManagerEmail;

        if (managerEmail && managerEmail.trim() !== '') {
          // Find manager by email (might be in the same batch or already exists)
          const managerUser = successful.find(u => u.userPrincipalName === managerEmail);

          if (managerUser) {
            managerAssignments.push({
              userId: user.id,
              managerId: managerUser.id,
              userName: user.displayName,
            });
          } else {
            // Manager might already exist in Azure AD
            try {
              const existingManager = await this.graphClient.getUserByEmail(managerEmail);
              managerAssignments.push({
                userId: user.id,
                managerId: existingManager.id,
                userName: user.displayName,
              });
            } catch (error) {
              logger.warn(`Manager not found for ${user.displayName}`, {
                userId: user.id,
                managerEmail,
              });
              console.warn(`⚠ Manager ${managerEmail} not found for ${user.displayName}`);
            }
          }
        }
      }

      if (managerAssignments.length > 0) {
        console.log(`  Assigning ${managerAssignments.length} manager relationships...`);
        const { successful: assignedManagers, failed: failedManagers } = await this.graphClient.assignManagersBatch(
          managerAssignments.map(m => ({ userId: m.userId, managerId: m.managerId }))
        );

        if (failedManagers.length > 0) {
          logger.error(`Failed to assign ${failedManagers.length} managers`, {
            failures: failedManagers,
          });
        }

        logger.info(`Assigned ${assignedManagers.length} manager relationships`);
      }


      // Track for export
      for (const user of successful) {
        const agent = usersToCreate.find(a => a.email === user.userPrincipalName);
        if (agent) {
          const csvUser = agents.find(a => a.email === agent.email);
          provisionedAgents.push({
            ...agent,
            name: agent.displayName,
            role: csvUser?.role || '',
            department: csvUser?.department || '',
            userId: user.id,
            password: agent.password,
            createdAt: new Date().toISOString(),
            lastAction: 'CREATE',
          });
        }
      }

      if (failed.length > 0) {
        console.warn(`⚠ ${failed.length} users failed to create`);
      }
    }

    // 2. Handle UPDATE operations
    if (!options.skipUpdate && delta.update.length > 0) {
      console.log(`✏️  Updating ${delta.update.length} users...`);

      // Prepare updates for batch (Option A - standard properties only)
      const standardUpdates: Array<{ userId: string; updates: Record<string, any> }> = [];

      for (const action of delta.update) {
        const userId = action.azureAdUser.id;

        // Only handle standard property changes (Option A)
        // Custom properties and enrichment data are handled by Option B (Graph Connectors)
        const standardChanges = action.changes?.filter(c => !c.isCustomProperty) || [];

        // Prepare standard property updates
        if (standardChanges.length > 0) {
          const updates: Record<string, any> = {};

          for (const change of standardChanges) {
            updates[change.field] = change.newValue;
          }

          standardUpdates.push({ userId, updates });
        }
      }

      // Apply standard updates in batch
      if (standardUpdates.length > 0) {
        const { successful } = await this.graphClient.updateUsersBatch(standardUpdates);
        updateCount += successful.length;
      }

      // Check and update manager assignments for updated users
      console.log(`  Checking manager assignments for updated users...`);
      const managerAssignments: Array<{ userId: string; managerId: string; userName: string }> = [];

      for (const action of delta.update) {
        const userId = action.azureAdUser.id;
        const email = action.user.email;
        const csvUser = agents.find(a => a.email === email);
        const managerEmail = (csvUser as any)?.ManagerEmail;

        // Check current manager
        const currentManager = await this.graphClient.getManager(userId);
        const currentManagerEmail = currentManager?.userPrincipalName;

        // If manager should be assigned/changed
        if (managerEmail && managerEmail.trim() !== '') {
          // Only update if manager has changed
          if (currentManagerEmail !== managerEmail) {
            // Find manager by email
            try {
              const managerUser = await this.graphClient.getUserByEmail(managerEmail);
              managerAssignments.push({
                userId,
                managerId: managerUser.id,
                userName: action.user.displayName,
              });
            } catch (error) {
              logger.warn(`Manager not found for ${action.user.displayName}`, {
                userId,
                managerEmail,
              });
              console.warn(`⚠ Manager ${managerEmail} not found for ${action.user.displayName}`);
            }
          }
        } else if (currentManager && (!managerEmail || managerEmail.trim() === '')) {
          // Manager should be removed (CSV has no manager, but Azure AD has one)
          await this.graphClient.removeManager(userId);
          logger.info(`Removed manager from ${action.user.displayName}`);
        }
      }

      if (managerAssignments.length > 0) {
        console.log(`  Updating ${managerAssignments.length} manager relationships...`);
        const { successful: assignedManagers, failed: failedManagers } = await this.graphClient.assignManagersBatch(
          managerAssignments.map(m => ({ userId: m.userId, managerId: m.managerId }))
        );

        if (failedManagers.length > 0) {
          logger.error(`Failed to update ${failedManagers.length} managers`, {
            failures: failedManagers,
          });
        }

        logger.info(`Updated ${assignedManagers.length} manager relationships`);
      }
    }

    // 3. Handle DELETE operations
    if (!options.skipDelete && delta.delete.length > 0) {
      console.log(`🗑️  Deleting ${delta.delete.length} users...`);

      const userIds = delta.delete.map(action => action.azureAdUser.id);
      const { successful } = await this.graphClient.deleteUsersBatch(userIds);
      deleteCount = successful.length;
    }

    // Export configuration (only for created/existing users)
    if (provisionedAgents.length > 0 || delta.update.length > 0 || delta.noChange.length > 0) {
      await this.exporter.exportConfig(provisionedAgents, options.outputPath);
      await this.exporter.exportPasswords(provisionedAgents);
      await this.exporter.generateReport();
    }

    // Summary
    console.log('\n' + '='.repeat(60));
    console.log('📊 Provisioning Summary');
    console.log('='.repeat(60));
    console.log(`✅ Created:   ${createCount} users`);
    console.log(`✏️  Updated:   ${updateCount} users`);
    console.log(`🗑️  Deleted:   ${deleteCount} users`);
    console.log(`➖ Unchanged: ${delta.noChange.length} users`);
    console.log('='.repeat(60));

    console.log('\n✨ Provisioning complete!\n');

    if (createCount > 0) {
      console.log('Next steps:');
      console.log('1. Review output/agents-config.json');
      console.log('2. Secure output/passwords.txt (do not commit!)');
      console.log('3. Wait 15-30 minutes for mailbox provisioning');
    }

    // Write summary and close logger
    const errorCount = logger.getLogBuffer().filter(e => e.level === 'error').length;
    const warnCount = logger.getLogBuffer().filter(e => e.level === 'warn').length;

    await logger.writeSummary({
      created: createCount,
      updated: updateCount,
      deleted: deleteCount,
      errors: errorCount,
      warnings: warnCount,
    });

    console.log(`\n📄 Log file: ${logger.getLogFilePath()}`);

    await logger.close();
  }

  /**
   * Clean up provisioned agents
   */
  async cleanup(): Promise<void> {
    console.log('🧹 Cleaning up provisioned agents\n');

    try {
      const config = await this.exporter.loadConfig();

      console.log(`Found ${config.agents.length} agents to delete\n`);
      console.log('⚠️  WARNING: This will permanently delete users and mailboxes!');
      console.log('Continue? (yes/no): ');

      // In a real implementation, you would prompt for confirmation
      // For now, we'll just list what would be deleted
      console.log('\nWould delete:');
      config.agents.forEach(agent => {
        console.log(`  - ${agent.name} (${agent.email})`);
      });

      console.log('\n❌ Cleanup cancelled (not implemented for safety)');
      console.log('To manually delete users, use Azure Portal or PowerShell');
    } catch (error: any) {
      throw new Error(`Cleanup failed: ${error.message}`);
    }
  }

  /**
   * Generate a secure random password
   */
  private generatePassword(): string {
    const length = 16;
    const charset = 'abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789!@#$%^&*';
    let password = '';

    password += 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'[Math.floor(Math.random() * 26)];
    password += 'abcdefghijklmnopqrstuvwxyz'[Math.floor(Math.random() * 26)];
    password += '0123456789'[Math.floor(Math.random() * 10)];
    password += '!@#$%^&*'[Math.floor(Math.random() * 8)];

    for (let i = password.length; i < length; i++) {
      password += charset[Math.floor(Math.random() * charset.length)];
    }

    return password.split('').sort(() => Math.random() - 0.5).join('');
  }
}

// Parse CLI arguments
function parseArgs(): Partial<ProvisionOptions> & { command?: string } {
  const args = process.argv.slice(2);
  const options: Partial<ProvisionOptions> & { command?: string } = {
    csvPath: 'config/agents-template.csv',
    outputPath: 'output/agents-config.json',
    dryRun: false,
    skipLicenses: false,
    force: false,
    auth: false,
    skipDelete: false,
    skipUpdate: false,
    skipCreate: false,
    showDiff: false,
  };

  for (let i = 0; i < args.length; i++) {
    const arg = args[i];

    switch (arg) {
      case '--dry-run':
        options.dryRun = true;
        break;
      case '--csv':
        options.csvPath = args[++i];
        break;
      case '--json':
        (options as any).jsonPath = args[++i];
        break;
      case '--output':
        options.outputPath = args[++i];
        break;
      case '--skip-licenses':
        options.skipLicenses = true;
        break;
      case '--force':
        options.force = true;
        break;
      case '--auth':
        options.auth = true;
        break;
      case '--skip-delete':
        options.skipDelete = true;
        break;
      case '--skip-update':
        options.skipUpdate = true;
        break;
      case '--skip-create':
        options.skipCreate = true;
        break;
      case '--show-diff':
        options.showDiff = true;
        break;
      case '--logout':
        options.command = 'logout';
        break;
      case '--cleanup':
        options.command = 'cleanup';
        break;
      case '--help':
      case '-h':
        options.command = 'help';
        break;
      default:
        if (arg.startsWith('--')) {
          console.error(`Unknown option: ${arg}`);
          process.exit(1);
        }
    }
  }

  return options;
}

function printHelp(): void {
  console.log(`
M365 Agent Provisioning Tool - State Management Mode

Usage: npm run provision [options]

State Management:
  CSV file is the source of truth. The tool syncs Azure AD to match the CSV exactly:
  - CREATE: Users in CSV but not in Azure AD
  - UPDATE: Users in both with differing attributes
  - DELETE: Users in Azure AD but not in CSV (requires --force)

Options:
  --dry-run              Preview changes without applying them
  --csv <path>           Use custom CSV file (default: config/agents-template.csv)
  --output <path>        Custom output path (default: output/agents-config.json)
  --skip-licenses        Skip license assignment
  --force                Proceed with deletion without confirmation
  --auth                 Force re-authentication (ignore cached token)
  --skip-delete          Don't delete users not in CSV (only CREATE and UPDATE)
  --skip-update          Only create new users (skip updates)
  --skip-create          Only update/delete existing users (skip creation)
  --show-diff            Show detailed diff report even in non-dry-run mode
  --logout               Clear cached authentication token
  --cleanup              Delete provisioned users
  --help, -h             Show this help message

Authentication:
  First run will prompt for browser-based authentication.
  Token is cached in ~/.m365-provision/ for future runs.

Examples:
  npm run provision                              # Full sync (with delete confirmation)
  npm run provision -- --dry-run                 # Preview what would change
  npm run provision -- --skip-delete             # Only CREATE and UPDATE users
  npm run provision -- --force                   # Skip delete confirmation
  npm run provision -- --show-diff               # Verbose diff output
  npm run provision -- --csv custom.csv          # Use custom CSV file
  npm run provision --                           # Uses Microsoft Graph beta endpoints

Custom Properties:
  Any CSV column not in the standard Microsoft Graph schema is deferred to Option B
  (Graph Connector). Option A ignores these columns during provisioning. For example:

  name,email,jobTitle,DeploymentManager,FavoriteColor
  John Doe,john@domain.com,Engineer,true,Blue

  Here, 'jobTitle' is a standard property, while 'DeploymentManager' and
  'FavoriteColor' are custom properties handled by Option B.

For more information, see USAGE.md
`);
}

// Main execution
async function main() {
  const options = parseArgs();

  if (options.command === 'help') {
    printHelp();
    return;
  }

  // Handle logout command
  if (options.command === 'logout') {
    const tenantId = process.env.AZURE_TENANT_ID;
    const clientId = process.env.AZURE_CLIENT_ID;

    if (!tenantId || !clientId) {
      console.error('❌ Error: AZURE_TENANT_ID and AZURE_CLIENT_ID must be set in .env file');
      process.exit(1);
    }

    const authServer = new BrowserAuthServer({ tenantId, clientId });
    authServer.clearCache();
    console.log('\n✅ Logged out successfully. Run provision again to re-authenticate.\n');
    return;
  }

  // Authenticate with browser-based MSAL
  const tenantId = process.env.AZURE_TENANT_ID;
  const clientId = process.env.AZURE_CLIENT_ID;
  const authPort = parseInt(process.env.AUTH_SERVER_PORT || '5544');

  if (!tenantId || !clientId) {
    console.error('❌ Error: AZURE_TENANT_ID and AZURE_CLIENT_ID must be set in .env file');
    process.exit(1);
  }

  try {
    // Start authentication server and wait for user to authenticate
    const authServer = new BrowserAuthServer({
      tenantId,
      clientId,
      port: authPort,
      forceRefresh: options.auth || false,
    });

    const authResult = await authServer.authenticate();

    // Create GraphClient with access token (beta endpoints enforced)
    const graphClient = new GraphClient({
      accessToken: authResult.accessToken,
    });

    if (options.command !== 'cleanup') {
      const csvPath = options.csvPath || 'config/agents-template.csv';
      const oidCacheResult = await ensureOidCacheWithClient({
        csvPath,
        tenantId,
        graphClient,
      });
      console.log(`🧭 OID cache: ${oidCacheResult.rebuilt ? 'built' : 'loaded'} (${oidCacheResult.cachePath})\n`);
    }

    const provisioner = new AgentProvisioner(graphClient);

    if (options.command === 'cleanup') {
      await provisioner.cleanup();
    } else {
      await provisioner.provisionAll(options as ProvisionOptions);
    }
  } catch (error: any) {
    console.error(`\n❌ Error: ${error.message}\n`);
    process.exit(1);
  }
}

// Run if executed directly
const __filename_prov = fileURLToPath(import.meta.url);
if (__filename_prov === path.resolve(process.argv[1])) {
  main();
}
