/**
 * Reset Password Tool
 *
 * Usage:
 *   Single user:  npm run reset-password -- user@domain.com NewPassword123!
 *   Bulk from CSV: npm run reset-password -- --csv config/users.csv --password NewPassword123!
 *
 * CSV format: Any CSV with an 'email' column (uses existing agent CSV files)
 */

import fs from 'fs';
import path from 'path';
import { parse } from 'csv-parse/sync';
import { BrowserAuthServer } from './auth/browser-auth-server.js';

interface ResetResult {
  email: string;
  success: boolean;
  error?: string;
}

class PasswordResetTool {
  private accessToken: string = '';

  async authenticate(): Promise<void> {
    const authServer = new BrowserAuthServer({
      tenantId: process.env.AZURE_TENANT_ID!,
      clientId: process.env.AZURE_CLIENT_ID!,
      port: parseInt(process.env.AUTH_SERVER_PORT || '5544'),
    });

    const authResult = await authServer.authenticate();
    this.accessToken = authResult.accessToken;
  }

  async resetPassword(email: string, newPassword: string): Promise<ResetResult> {
    try {
      // Find user
      const userResponse = await fetch(
        `https://graph.microsoft.com/v1.0/users/${encodeURIComponent(email)}`,
        {
          headers: { Authorization: `Bearer ${this.accessToken}` },
        }
      );

      if (!userResponse.ok) {
        const error = await userResponse.text();
        return { email, success: false, error: `User not found: ${error}` };
      }

      const user = await userResponse.json() as { id: string };

      // Reset password
      const resetResponse = await fetch(
        `https://graph.microsoft.com/v1.0/users/${user.id}`,
        {
          method: 'PATCH',
          headers: {
            Authorization: `Bearer ${this.accessToken}`,
            'Content-Type': 'application/json',
          },
          body: JSON.stringify({
            passwordProfile: {
              password: newPassword,
              forceChangePasswordNextSignIn: false,
            },
          }),
        }
      );

      if (!resetResponse.ok) {
        const error = await resetResponse.text();
        return { email, success: false, error: `Reset failed: ${error}` };
      }

      return { email, success: true };
    } catch (error: any) {
      return { email, success: false, error: error.message };
    }
  }

  async resetFromCsv(csvPath: string, newPassword: string): Promise<ResetResult[]> {
    const absolutePath = path.resolve(csvPath);
    if (!fs.existsSync(absolutePath)) {
      throw new Error(`CSV file not found: ${absolutePath}`);
    }

    const content = fs.readFileSync(absolutePath, 'utf8');
    const records = parse(content, {
      columns: true,
      skip_empty_lines: true,
      trim: true,
    });

    const results: ResetResult[] = [];
    console.log(`\n📋 Found ${records.length} users in CSV\n`);

    for (const record of records) {
      const email = record.email || record.Email || record.EMAIL;
      if (!email) {
        console.log(`⚠ Skipping row without email: ${JSON.stringify(record)}`);
        continue;
      }

      process.stdout.write(`  Resetting ${email}... `);
      const result = await this.resetPassword(email, newPassword);
      results.push(result);

      if (result.success) {
        console.log('✅');
      } else {
        console.log(`❌ ${result.error}`);
      }
    }

    return results;
  }
}

function printUsage(): void {
  console.log(`
🔑 Password Reset Tool

Usage:
  Single user:
    npm run reset-password -- <email> <password>

  Bulk from CSV:
    npm run reset-password -- --csv <file.csv> --password <password>

Examples:
  npm run reset-password -- user@domain.com "NewPass123!"
  npm run reset-password -- --csv config/agents.csv --password "NewPass123!"

Options:
  --csv       Path to CSV file with 'email' column
  --password  New password to set for all users in CSV

Note: Password must meet Microsoft's complexity requirements:
  - Minimum 8 characters
  - Contains uppercase, lowercase, number, and symbol
`);
}

async function main(): Promise<void> {
  const args = process.argv.slice(2);

  if (args.length === 0 || args.includes('--help') || args.includes('-h')) {
    printUsage();
    process.exit(0);
  }

  const tool = new PasswordResetTool();

  // Parse arguments
  const csvIndex = args.indexOf('--csv');
  const passwordIndex = args.indexOf('--password');

  if (csvIndex !== -1) {
    // Bulk mode: --csv <file> --password <pass>
    const csvPath = args[csvIndex + 1];
    const password = passwordIndex !== -1 ? args[passwordIndex + 1] : process.env.USER_PASSWORD_PROFILE;

    if (!csvPath) {
      console.error('❌ Missing CSV file path');
      printUsage();
      process.exit(1);
    }

    if (!password) {
      console.error('❌ Missing password. Use --password or set USER_PASSWORD_PROFILE in .env');
      printUsage();
      process.exit(1);
    }

    console.log(`\n🔑 Bulk Password Reset`);
    console.log(`   CSV: ${csvPath}`);
    console.log(`   Password: ${password}\n`);

    await tool.authenticate();
    const results = await tool.resetFromCsv(csvPath, password);

    // Summary
    const successful = results.filter((r) => r.success).length;
    const failed = results.filter((r) => !r.success).length;

    console.log(`\n📊 Summary:`);
    console.log(`   ✅ Successful: ${successful}`);
    console.log(`   ❌ Failed: ${failed}`);

    if (failed > 0) {
      console.log(`\n   Failed users:`);
      results
        .filter((r) => !r.success)
        .forEach((r) => console.log(`   - ${r.email}: ${r.error}`));
    }
  } else {
    // Single user mode: <email> <password>
    const email = args[0];
    const password = args[1] || process.env.USER_PASSWORD_PROFILE;

    if (!email) {
      console.error('❌ Missing email address');
      printUsage();
      process.exit(1);
    }

    if (!password) {
      console.error('❌ Missing password. Provide as argument or set USER_PASSWORD_PROFILE in .env');
      printUsage();
      process.exit(1);
    }

    console.log(`\n🔑 Password Reset`);
    console.log(`   Email: ${email}`);
    console.log(`   Password: ${password}\n`);

    await tool.authenticate();
    const result = await tool.resetPassword(email, password);

    if (result.success) {
      console.log(`✅ Password reset successfully!`);
      console.log(`\n   User can now sign in with:`);
      console.log(`   Email: ${email}`);
      console.log(`   Password: ${password}`);
    } else {
      console.error(`❌ Failed: ${result.error}`);
      process.exit(1);
    }
  }
}

main().catch((error) => {
  console.error(`\n❌ Error: ${error.message}`);
  process.exit(1);
});
