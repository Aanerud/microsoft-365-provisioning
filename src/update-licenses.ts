#!/usr/bin/env node

/**
 * Update Licenses
 *
 * Assigns licenses from LICENSE_SKU_IDS to existing users.
 * Useful when:
 * - Adding a new license to all users (e.g., Copilot license)
 * - Ensuring all users have the required licenses
 *
 * This command is idempotent - already-assigned licenses are skipped gracefully.
 */

import fs from 'fs/promises';
import { parse } from 'csv-parse/sync';
import dotenv from 'dotenv';
import { GraphClient } from './graph-client.js';
import { BrowserAuthServer } from './auth/browser-auth-server.js';

dotenv.config();

interface UpdateLicensesOptions {
  csvPath: string;
  dryRun: boolean;
  auth: boolean;
}

interface LicenseUpdateResult {
  email: string;
  displayName: string;
  userId: string;
  licensesAdded: string[];
  licensesSkipped: string[];
  errors: string[];
}

class LicenseUpdater {
  private graphClient: GraphClient;
  private licenseSkuIds: string[];

  constructor(graphClient: GraphClient) {
    this.graphClient = graphClient;

    // Support both LICENSE_SKU_IDS (comma-separated) and legacy LICENSE_SKU_ID
    const skuIdsEnv = process.env.LICENSE_SKU_IDS || process.env.LICENSE_SKU_ID || '';
    this.licenseSkuIds = skuIdsEnv
      .split(',')
      .map(id => id.trim())
      .filter(id => id.length > 0);

    if (this.licenseSkuIds.length === 0) {
      throw new Error('LICENSE_SKU_IDS not set in .env - cannot update licenses');
    }

    console.log(`📋 Configured ${this.licenseSkuIds.length} license(s) for assignment:`);
    this.licenseSkuIds.forEach(sku => console.log(`   - ${sku}`));
    console.log('');
  }

  /**
   * Load user emails from CSV file
   */
  async loadUsersFromCsv(csvPath: string): Promise<Array<{ email: string; name: string }>> {
    const content = await fs.readFile(csvPath, 'utf-8');
    const records = parse(content, {
      columns: true,
      skip_empty_lines: true,
      trim: true,
    });

    console.log(`✓ Loaded ${records.length} users from ${csvPath}\n`);

    return records.map((record: any) => ({
      email: record.email,
      name: record.name || record.displayName,
    }));
  }

  /**
   * Get current licenses for a user
   */
  async getUserLicenses(userId: string): Promise<string[]> {
    try {
      const { client } = this.graphClient.getClients();
      const response = await client.api(`/users/${userId}/licenseDetails`).get();
      return response.value.map((license: any) => license.skuId);
    } catch (error: any) {
      console.warn(`  ⚠ Could not fetch licenses for user ${userId}: ${error.message}`);
      return [];
    }
  }

  /**
   * Update licenses for all users in CSV
   */
  async updateAll(options: UpdateLicensesOptions): Promise<void> {
    console.log('🔄 Starting License Update\n');
    console.log('='.repeat(60));

    if (options.dryRun) {
      console.log('🔍 DRY RUN MODE - No changes will be made\n');
    }

    // Load users from CSV
    const csvUsers = await this.loadUsersFromCsv(options.csvPath);

    // Fetch all Azure AD users to map emails to user IDs
    console.log('🔍 Fetching Azure AD users...');
    const azureUsers = await this.graphClient.listUsers();
    const userMap = new Map<string, { id: string; displayName: string }>();
    for (const user of azureUsers) {
      userMap.set(user.userPrincipalName.toLowerCase(), {
        id: user.id,
        displayName: user.displayName,
      });
    }
    console.log(`  Found ${azureUsers.length} users in Azure AD\n`);

    // Process each user
    const results: LicenseUpdateResult[] = [];
    let totalAdded = 0;
    let totalSkipped = 0;
    let totalErrors = 0;
    let notFound = 0;

    console.log('📝 Processing users...\n');

    for (const csvUser of csvUsers) {
      const azureUser = userMap.get(csvUser.email.toLowerCase());

      if (!azureUser) {
        console.log(`  ⚠ User not found in Azure AD: ${csvUser.email}`);
        notFound++;
        continue;
      }

      const result: LicenseUpdateResult = {
        email: csvUser.email,
        displayName: azureUser.displayName,
        userId: azureUser.id,
        licensesAdded: [],
        licensesSkipped: [],
        errors: [],
      };

      // Get current licenses
      const currentLicenses = await this.getUserLicenses(azureUser.id);
      const currentLicenseSet = new Set(currentLicenses.map(l => l.toLowerCase()));

      // Determine which licenses need to be added
      const licensesToAdd: string[] = [];
      for (const skuId of this.licenseSkuIds) {
        if (currentLicenseSet.has(skuId.toLowerCase())) {
          result.licensesSkipped.push(skuId);
          totalSkipped++;
        } else {
          licensesToAdd.push(skuId);
        }
      }

      if (licensesToAdd.length === 0) {
        console.log(`  ✓ ${azureUser.displayName} - all licenses already assigned`);
        results.push(result);
        continue;
      }

      // Assign missing licenses
      if (options.dryRun) {
        console.log(`  📋 ${azureUser.displayName} - would add ${licensesToAdd.length} license(s)`);
        result.licensesAdded = licensesToAdd;
        totalAdded += licensesToAdd.length;
      } else {
        console.log(`  🔄 ${azureUser.displayName} - adding ${licensesToAdd.length} license(s)...`);

        for (const skuId of licensesToAdd) {
          try {
            await this.graphClient.assignLicense(azureUser.id, skuId);
            result.licensesAdded.push(skuId);
            totalAdded++;
          } catch (error: any) {
            const errorMsg = error.message || 'Unknown error';
            result.errors.push(`${skuId}: ${errorMsg}`);
            totalErrors++;
            console.log(`    ✗ Failed to assign ${skuId}: ${errorMsg}`);
          }
        }

        if (result.licensesAdded.length > 0) {
          console.log(`    ✓ Added ${result.licensesAdded.length} license(s)`);
        }
      }

      results.push(result);
    }

    // Print summary
    console.log('\n' + '='.repeat(60));
    console.log('📊 License Update Summary');
    console.log('='.repeat(60));
    console.log(`  Users in CSV:        ${csvUsers.length}`);
    console.log(`  Users found:         ${csvUsers.length - notFound}`);
    console.log(`  Users not found:     ${notFound}`);
    console.log(`  Licenses added:      ${totalAdded}`);
    console.log(`  Already assigned:    ${totalSkipped}`);
    if (totalErrors > 0) {
      console.log(`  Errors:              ${totalErrors}`);
    }
    console.log('='.repeat(60));

    if (options.dryRun) {
      console.log('\n💡 Run without --dry-run to apply changes');
    } else {
      console.log('\n✨ License update complete!');
    }
  }
}

function printHelp(): void {
  console.log(`
Update Licenses - Assign licenses to existing users

Usage:
  npm run update-licenses -- [options]

Description:
  Reads users from a CSV file and assigns any missing licenses from LICENSE_SKU_IDS.
  This is useful when you've added a new license (e.g., Copilot) to your .env and
  want to assign it to all existing users.

  The command is idempotent - already-assigned licenses are skipped gracefully.

Options:
  --csv <path>           CSV file with user emails (default: config/agents-template.csv)
  --dry-run              Preview which licenses would be assigned (no changes)
  --auth                 Force re-authentication (ignore cached token)
  --help, -h             Show this help message

Examples:
  # Preview license changes
  npm run update-licenses -- --dry-run --csv config/textcraft-europe.csv

  # Apply license updates
  npm run update-licenses -- --csv config/textcraft-europe.csv

  # Update licenses for default CSV
  npm run update-licenses

Environment:
  LICENSE_SKU_IDS        Comma-separated list of license SKU IDs to assign
                         Example: LICENSE_SKU_IDS=sku-id-1,sku-id-2

  Run 'npm run list-licenses' to see available licenses in your tenant.
`);
}

async function main(): Promise<void> {
  const args = process.argv.slice(2);

  // Parse arguments
  const options: UpdateLicensesOptions = {
    csvPath: 'config/agents-template.csv',
    dryRun: false,
    auth: false,
  };

  for (let i = 0; i < args.length; i++) {
    switch (args[i]) {
      case '--csv':
        options.csvPath = args[++i];
        break;
      case '--dry-run':
        options.dryRun = true;
        break;
      case '--auth':
        options.auth = true;
        break;
      case '--help':
      case '-h':
        printHelp();
        process.exit(0);
    }
  }

  // Check required environment variables
  const tenantId = process.env.AZURE_TENANT_ID;
  const clientId = process.env.AZURE_CLIENT_ID;

  if (!tenantId || !clientId) {
    console.error('Error: AZURE_TENANT_ID and AZURE_CLIENT_ID must be set in .env');
    process.exit(1);
  }

  // Check for license SKU IDs
  const skuIdsEnv = process.env.LICENSE_SKU_IDS || process.env.LICENSE_SKU_ID || '';
  if (!skuIdsEnv.trim()) {
    console.error('Error: LICENSE_SKU_IDS must be set in .env');
    console.error('Run "npm run list-licenses" to see available licenses');
    process.exit(1);
  }

  // Check CSV file exists
  try {
    await fs.access(options.csvPath);
  } catch {
    console.error(`Error: CSV file not found: ${options.csvPath}`);
    process.exit(1);
  }

  console.log('🔑 License Update Tool\n');
  console.log('Configuration:');
  console.log(`  CSV: ${options.csvPath}`);
  console.log(`  Dry Run: ${options.dryRun ? 'Yes' : 'No'}`);
  console.log('');

  // Authenticate using browser-based flow (delegated permissions)
  const authServer = new BrowserAuthServer({
    tenantId,
    clientId,
    forceRefresh: options.auth,
  });

  console.log('Starting browser authentication (delegated permissions)...\n');
  const authResult = await authServer.authenticate();
  console.log(`✅ Authenticated as: ${authResult.account.username}\n`);

  // Create Graph client
  const graphClient = new GraphClient({ accessToken: authResult.accessToken });

  // Run license update
  const updater = new LicenseUpdater(graphClient);
  await updater.updateAll(options);
}

main().catch(error => {
  console.error('Fatal error:', error.message);
  process.exit(1);
});
