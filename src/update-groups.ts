#!/usr/bin/env node

/**
 * Update Group Memberships
 *
 * Reads users from JSON or CSV and ensures they are members of the specified groups.
 * Groups must already exist in Entra ID. This command is idempotent.
 *
 * Usage:
 *   npm run update-groups -- --json config/users.config.json
 *   npm run update-groups -- --json config/users.config.json --dry-run
 */

import fs from 'fs/promises';
import dotenv from 'dotenv';
import { GraphClient } from './graph-client.js';
import { BrowserAuthServer } from './auth/browser-auth-server.js';
import { loadRowsFromJson } from './json-loader.js';

dotenv.config();

interface UpdateGroupsOptions {
  jsonPath?: string;
  csvPath: string;
  dryRun: boolean;
  auth: boolean;
}

async function loadUsers(options: UpdateGroupsOptions): Promise<any[]> {
  const inputPath = options.jsonPath || options.csvPath;

  if (inputPath.endsWith('.json')) {
    const records = await loadRowsFromJson(inputPath);
    console.log(`✓ Loaded ${records.length} users from JSON ${inputPath}\n`);
    return records;
  }

  // CSV — parse groups as comma-separated string
  const { parse } = await import('csv-parse/sync');
  const content = await fs.readFile(inputPath, 'utf-8');
  const records = parse(content, { columns: true, skip_empty_lines: true, trim: true });
  console.log(`✓ Loaded ${records.length} users from ${inputPath}\n`);
  return records;
}

async function main(): Promise<void> {
  const args = process.argv.slice(2);
  const options: UpdateGroupsOptions = {
    csvPath: 'config/agents-template.csv',
    dryRun: false,
    auth: false,
  };

  for (let i = 0; i < args.length; i++) {
    switch (args[i]) {
      case '--csv': options.csvPath = args[++i]; break;
      case '--json': options.jsonPath = args[++i]; break;
      case '--dry-run': options.dryRun = true; break;
      case '--auth': options.auth = true; break;
      case '--help': case '-h':
        console.log(`
Update Group Memberships — Assign users to Entra ID groups

Usage:
  npm run update-groups -- --json config/users.config.json
  npm run update-groups -- --json config/users.config.json --dry-run

Options:
  --json <path>   JSON file with user data (PascalCase auto-detected)
  --csv <path>    CSV file with user data
  --dry-run       Preview changes without applying
  --auth          Force re-authentication

Groups must already exist in Entra ID. The JSON/CSV "Groups" field
should be an array of group display names.
`);
        process.exit(0);
    }
  }

  const tenantId = process.env.AZURE_TENANT_ID;
  const clientId = process.env.AZURE_CLIENT_ID;
  if (!tenantId || !clientId) {
    console.error('Error: AZURE_TENANT_ID and AZURE_CLIENT_ID must be set in .env');
    process.exit(1);
  }

  const inputPath = options.jsonPath || options.csvPath;
  try {
    await fs.access(inputPath);
  } catch {
    console.error(`Error: Input file not found: ${inputPath}`);
    process.exit(1);
  }

  console.log('🔑 Group Membership Update Tool\n');
  console.log('Configuration:');
  console.log(`  Input: ${inputPath}`);
  console.log(`  Dry Run: ${options.dryRun ? 'Yes' : 'No'}`);
  console.log('');

  // Authenticate
  const authServer = new BrowserAuthServer({ tenantId, clientId, forceRefresh: options.auth });
  console.log('Starting browser authentication (delegated permissions)...\n');
  const authResult = await authServer.authenticate();
  console.log(`✅ Authenticated as: ${authResult.account.username}\n`);

  const graphClient = new GraphClient({ accessToken: authResult.accessToken });

  // Load users
  const users = await loadUsers(options);

  // Collect unique group names across all users
  const allGroupNames = new Set<string>();
  const usersWithGroups: Array<{ email: string; name: string; groups: string[] }> = [];

  for (const user of users) {
    const groups = user.groups;
    if (!Array.isArray(groups) || groups.length === 0) continue;

    // Handle both array of strings and comma-separated string
    const groupList = groups.flatMap((g: any) =>
      typeof g === 'string' ? g.split(',').map((s: string) => s.trim()).filter(Boolean) : []
    );

    if (groupList.length > 0) {
      usersWithGroups.push({
        email: user.email,
        name: user.name || user.displayName || user.email,
        groups: groupList,
      });
      groupList.forEach(g => allGroupNames.add(g));
    }
  }

  if (usersWithGroups.length === 0) {
    console.log('No users with group assignments found.');
    return;
  }

  console.log(`Found ${usersWithGroups.length} users with group assignments`);
  console.log(`Unique groups: ${allGroupNames.size}\n`);

  // Resolve group names → IDs
  console.log('🔍 Resolving group names...');
  const groupMap = await graphClient.getGroupsByNames([...allGroupNames]);

  for (const name of allGroupNames) {
    if (groupMap.has(name)) {
      console.log(`  ✓ "${name}" → ${groupMap.get(name)}`);
    } else {
      console.warn(`  ⚠ "${name}" → NOT FOUND in Entra ID`);
    }
  }
  console.log('');

  // Fetch Azure AD users to map emails → OIDs
  console.log('🔍 Fetching Azure AD users...');
  const azureUsers = await graphClient.listUsers();
  const userMap = new Map<string, { id: string; displayName: string }>();
  for (const u of azureUsers) {
    userMap.set(u.userPrincipalName.toLowerCase(), { id: u.id, displayName: u.displayName });
  }
  console.log(`  Found ${azureUsers.length} users in Azure AD\n`);

  // Build assignment list
  const assignments: Array<{ groupId: string; userId: string; groupName: string; userName: string }> = [];
  let notFound = 0;
  let noGroup = 0;

  for (const user of usersWithGroups) {
    const azureUser = userMap.get(user.email.toLowerCase());
    if (!azureUser) {
      console.warn(`  ⚠ User not found in Azure AD: ${user.email}`);
      notFound++;
      continue;
    }

    for (const groupName of user.groups) {
      const groupId = groupMap.get(groupName);
      if (!groupId) {
        noGroup++;
        continue; // Already warned above
      }

      assignments.push({
        groupId,
        userId: azureUser.id,
        groupName,
        userName: azureUser.displayName,
      });
    }
  }

  console.log(`📝 ${assignments.length} group membership assignments to process\n`);

  if (options.dryRun) {
    console.log('🔍 DRY RUN — Preview:\n');
    const byGroup = new Map<string, string[]>();
    for (const a of assignments) {
      if (!byGroup.has(a.groupName)) byGroup.set(a.groupName, []);
      byGroup.get(a.groupName)!.push(a.userName);
    }
    for (const [group, members] of byGroup) {
      console.log(`  ${group} (${members.length} members):`);
      members.forEach(m => console.log(`    + ${m}`));
    }
    console.log(`\n💡 Run without --dry-run to apply changes`);
    return;
  }

  // Apply assignments
  console.log('🔄 Assigning group memberships...\n');
  const result = await graphClient.addGroupMembersBatch(assignments);

  // Summary
  console.log('\n' + '='.repeat(60));
  console.log('📊 Group Membership Update Summary');
  console.log('='.repeat(60));
  console.log(`  Users processed:     ${usersWithGroups.length}`);
  console.log(`  Users not found:     ${notFound}`);
  console.log(`  Groups resolved:     ${groupMap.size}/${allGroupNames.size}`);
  console.log(`  Members added:       ${result.added}`);
  console.log(`  Already members:     ${result.alreadyMember}`);
  if (result.failed.length > 0) {
    console.log(`  Errors:              ${result.failed.length}`);
    result.failed.forEach(f => console.log(`    ✗ ${f.userName} → ${f.groupName}: ${f.error}`));
  }
  console.log('='.repeat(60));
  console.log('\n✨ Group membership update complete!');
}

main().catch(error => {
  console.error('Fatal error:', error.message);
  process.exit(1);
});
