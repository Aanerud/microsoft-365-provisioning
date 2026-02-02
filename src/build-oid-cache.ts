#!/usr/bin/env node

import dotenv from 'dotenv';
import { ensureOidCacheWithAuth } from './oid-cache.js';

dotenv.config();

interface BuildOidCacheOptions {
  csvPath: string;
  force: boolean;
  auth: boolean;
}

function parseArgs(): BuildOidCacheOptions {
  const args = process.argv.slice(2);
  const options: BuildOidCacheOptions = {
    csvPath: 'config/agents-template.csv',
    force: false,
    auth: false,
  };

  for (let i = 0; i < args.length; i++) {
    switch (args[i]) {
      case '--csv':
        options.csvPath = args[++i];
        break;
      case '--force':
        options.force = true;
        break;
      case '--auth':
        options.auth = true;
        break;
      case '--help':
      case '-h':
        printHelp();
        process.exit(0);
      default:
        if (args[i].startsWith('--')) {
          console.error(`Unknown option: ${args[i]}`);
          process.exit(1);
        }
        break;
    }
  }

  return options;
}

function printHelp(): void {
  console.log(`
Build OID Cache (Delegated Auth)

Usage: npm run build-oid-cache -- [options]

Options:
  --csv <path>   CSV file to derive cache name (default: config/agents-template.csv)
  --force        Rebuild cache even if it exists
  --auth         Force re-authentication (ignore cached token)
  --help, -h     Show help

This tool builds a JSON cache mapping UPN -> externalDirectoryObjectId using
Microsoft Graph beta /users list. The cache file is created next to the CSV:
  <csv-name>_oid_cache.json
`);
}

async function main(): Promise<void> {
  const options = parseArgs();
  const tenantId = process.env.AZURE_TENANT_ID;
  const clientId = process.env.AZURE_CLIENT_ID;

  if (!tenantId || !clientId) {
    console.error('❌ Error: AZURE_TENANT_ID and AZURE_CLIENT_ID must be set in .env file');
    process.exit(1);
  }

  const result = await ensureOidCacheWithAuth({
    csvPath: options.csvPath,
    tenantId,
    clientId,
    authPort: parseInt(process.env.AUTH_SERVER_PORT || '5544', 10),
    force: options.force,
    forceRefresh: options.auth,
  });

  console.log(`\n✅ OID cache ready (${result.rebuilt ? 'rebuilt' : 'loaded'})`);
  console.log(`   Path: ${result.cachePath}`);
  console.log(`   Users: ${result.cache.userCount}`);
}

main().catch((error: any) => {
  console.error(`\n❌ Error: ${error.message}\n`);
  process.exit(1);
});
