#!/usr/bin/env node

/**
 * Tenant Reset Script
 *
 * Safely removes all provisioned users from the tenant while protecting
 * admin accounts and other critical users.
 *
 * Features:
 * - Uses AccountProtectionService to protect admin accounts
 * - Dry-run mode by default (requires --confirm to actually delete)
 * - Batch deletion for efficiency
 * - Optional Graph Connector cleanup
 *
 * Usage:
 *   npm run reset-tenant              # Dry run - shows what would be deleted
 *   npm run reset-tenant -- --confirm # Actually delete users
 *   npm run reset-tenant -- --confirm --include-connector  # Also clean connector items
 */

import dotenv from 'dotenv';
import fs from 'fs/promises';
import path from 'path';
import { GraphClient } from './graph-client.js';
import { BrowserAuthServer } from './auth/browser-auth-server.js';
import { AccountProtectionService } from './safety/account-protection.js';

dotenv.config();

interface ResetOptions {
  confirm: boolean;
  includeConnector: boolean;
  verbose: boolean;
}

interface User {
  id: string;
  displayName: string;
  userPrincipalName: string;
  mail?: string;
}

async function resetTenant(options: ResetOptions): Promise<void> {
  console.log('\n' + '═'.repeat(60));
  console.log('🔄 TENANT RESET SCRIPT');
  console.log('═'.repeat(60));
  console.log(`\nMode: ${options.confirm ? '⚠️  LIVE - WILL DELETE USERS' : '👀 DRY RUN - Preview only'}`);
  console.log('');

  // Check for required environment variables
  const tenantId = process.env.AZURE_TENANT_ID;
  const clientId = process.env.AZURE_CLIENT_ID;

  if (!tenantId || !clientId) {
    console.error('❌ Error: AZURE_TENANT_ID and AZURE_CLIENT_ID must be set in .env file');
    process.exit(1);
  }

  // Authenticate using browser-based auth (requires delegated permissions)
  console.log('🔐 Authenticating...');
  const authServer = new BrowserAuthServer({ tenantId, clientId });
  const authResult = await authServer.authenticate();
  console.log(`✅ Authenticated as: ${authResult.account.username}\n`);

  // Initialize Graph client
  const graphClient = new GraphClient({ accessToken: authResult.accessToken });

  // Initialize account protection service
  const protectionService = AccountProtectionService.fromEnvironment(graphClient);
  console.log(protectionService.getProtectionConfig());
  console.log('');

  // Fetch all users
  console.log('📋 Fetching all users from tenant...');
  const allUsers = await graphClient.listUsers();
  console.log(`   Found ${allUsers.length} total users\n`);

  if (allUsers.length === 0) {
    console.log('✅ No users found in tenant. Nothing to do.');
    return;
  }

  // Filter protected accounts
  console.log('🛡️  Checking account protection...');
  const accountsToCheck = allUsers.map(user => ({
    email: user.userPrincipalName,
    userId: user.id,
    displayName: user.displayName,
  }));

  const { allowed, protectedAccounts } = await protectionService.filterProtectedAccounts(accountsToCheck);

  // Display protected accounts
  if (protectedAccounts.length > 0) {
    console.log(`\n🛡️  PROTECTED ACCOUNTS (${protectedAccounts.length}) - Will NOT be deleted:`);
    console.log('─'.repeat(60));
    protectedAccounts.forEach((account, i) => {
      console.log(`   ${i + 1}. ${account.email}`);
      console.log(`      Reason: ${account.reason}`);
      if (account.role) {
        console.log(`      Role: ${account.role}`);
      }
    });
    console.log('');
  }

  // Get full user objects for users to delete
  const usersToDelete = allUsers.filter(user =>
    allowed.some(a => a.email === user.userPrincipalName)
  );

  // Display users to delete
  console.log(`\n🗑️  USERS TO DELETE (${usersToDelete.length}):`);
  console.log('─'.repeat(60));

  if (usersToDelete.length === 0) {
    console.log('   No users to delete (all users are protected)');
    return;
  }

  // Group by domain for display
  const byDomain: Record<string, User[]> = {};
  usersToDelete.forEach(user => {
    const domain = user.userPrincipalName.split('@')[1] || 'unknown';
    if (!byDomain[domain]) byDomain[domain] = [];
    byDomain[domain].push(user);
  });

  Object.entries(byDomain).forEach(([domain, users]) => {
    console.log(`\n   📁 ${domain} (${users.length} users)`);
    if (options.verbose || users.length <= 10) {
      users.forEach(user => {
        console.log(`      - ${user.displayName} (${user.userPrincipalName})`);
      });
    } else {
      users.slice(0, 5).forEach(user => {
        console.log(`      - ${user.displayName} (${user.userPrincipalName})`);
      });
      console.log(`      ... and ${users.length - 5} more`);
    }
  });

  console.log('\n' + '═'.repeat(60));
  console.log(`📊 SUMMARY:`);
  console.log(`   Total users in tenant: ${allUsers.length}`);
  console.log(`   Protected (keep):      ${protectedAccounts.length}`);
  console.log(`   To be deleted:         ${usersToDelete.length}`);
  console.log('═'.repeat(60));

  // Check for dry run
  if (!options.confirm) {
    console.log('\n⚠️  DRY RUN MODE - No users were deleted');
    console.log('   To actually delete users, run with --confirm flag:');
    console.log('   npm run reset-tenant -- --confirm\n');
    return;
  }

  // Confirm deletion
  console.log('\n⚠️  WARNING: This action cannot be undone!');
  console.log(`   About to delete ${usersToDelete.length} users from the tenant.`);
  console.log('');

  // Actually delete users
  console.log('🗑️  Deleting users...\n');

  const userIds = usersToDelete.map(u => u.id);
  const { successful, failed } = await graphClient.deleteUsersBatch(userIds);

  // Report results
  console.log('\n' + '═'.repeat(60));
  console.log('📊 DELETION RESULTS:');
  console.log(`   ✅ Successfully deleted: ${successful.length}`);
  console.log(`   ❌ Failed to delete:     ${failed.length}`);

  if (failed.length > 0) {
    console.log('\n   Failed deletions:');
    failed.forEach(f => {
      const user = usersToDelete.find(u => u.id === f.userId);
      console.log(`   - ${user?.displayName || f.userId}: ${f.error}`);
    });
  }

  // Clean up Graph Connector items if requested
  if (options.includeConnector) {
    console.log('\n🔗 Cleaning up Graph Connector items...');
    await cleanupConnectorItems(usersToDelete);
  }

  // Clean up state files
  await cleanupStateFiles();

  console.log('\n✅ Tenant reset complete!\n');
}

async function cleanupConnectorItems(deletedUsers: User[]): Promise<void> {
  const stateFilePath = path.join(process.cwd(), 'state', 'external-items-state.json');

  try {
    const stateContent = await fs.readFile(stateFilePath, 'utf-8');
    const state = JSON.parse(stateContent);

    // Remove items for deleted users
    const deletedEmails = new Set(deletedUsers.map(u => u.userPrincipalName.toLowerCase()));
    const originalCount = Object.keys(state.items || {}).length;

    if (state.items) {
      Object.keys(state.items).forEach(key => {
        const item = state.items[key];
        if (item.email && deletedEmails.has(item.email.toLowerCase())) {
          delete state.items[key];
        }
      });
    }

    const newCount = Object.keys(state.items || {}).length;
    const removed = originalCount - newCount;

    await fs.writeFile(stateFilePath, JSON.stringify(state, null, 2));
    console.log(`   Removed ${removed} items from connector state file`);
  } catch (error: any) {
    if (error.code === 'ENOENT') {
      console.log('   No connector state file found (skipping)');
    } else {
      console.warn(`   ⚠️ Could not clean connector state: ${error.message}`);
    }
  }
}

async function cleanupStateFiles(): Promise<void> {
  console.log('\n🧹 Cleaning up state files...');

  const filesToClean = [
    'output/agents-config.json',
    'output/provisioning-report.md',
    'output/passwords.txt',
  ];

  for (const file of filesToClean) {
    const filePath = path.join(process.cwd(), file);
    try {
      await fs.unlink(filePath);
      console.log(`   Deleted: ${file}`);
    } catch (error: any) {
      if (error.code !== 'ENOENT') {
        console.warn(`   ⚠️ Could not delete ${file}: ${error.message}`);
      }
    }
  }
}

// Parse command line arguments
function parseArgs(): ResetOptions {
  const args = process.argv.slice(2);
  return {
    confirm: args.includes('--confirm') || args.includes('-y'),
    includeConnector: args.includes('--include-connector') || args.includes('-c'),
    verbose: args.includes('--verbose') || args.includes('-v'),
  };
}

// Main execution
const options = parseArgs();

resetTenant(options).catch(error => {
  console.error(`\n❌ Error: ${error.message}\n`);
  process.exit(1);
});
