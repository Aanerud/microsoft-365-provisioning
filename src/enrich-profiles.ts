#!/usr/bin/env node
/**
 * Option A: Profile API Enrichment
 *
 * Writes profile data that has NO people data labels via the Profile API.
 * Uses delegated authentication (browser login).
 *
 * Handles:
 *   - languages → POST /users/{id}/profile/languages
 *   - interests → POST /users/{id}/profile/interests
 *
 * These fields have no Graph Connector people data label, so they cannot
 * be made Copilot-searchable. They appear on profile cards only.
 */

import fs from 'fs/promises';
import { parse } from 'csv-parse/sync';
import dotenv from 'dotenv';
import { BrowserAuthServer } from './auth/browser-auth-server.js';
import {
  ProfileWriter,
  PersonInterest,
  LanguageProficiency,
} from './profile-writer.js';

dotenv.config();

interface EnrichOptions {
  csvPath: string;
  dryRun: boolean;
}

interface UserProfile {
  email: string;
  interests?: string[];
  languages?: LanguageProficiency[];
}

function parseArgs(): EnrichOptions {
  const args = process.argv.slice(2);
  const options: EnrichOptions = {
    csvPath: 'config/textcraft-europe.csv',
    dryRun: false,
  };

  for (let i = 0; i < args.length; i++) {
    switch (args[i]) {
      case '--csv':
        options.csvPath = args[++i];
        break;
      case '--dry-run':
        options.dryRun = true;
        break;
    }
  }

  return options;
}

function parseArrayValue(value: string): string[] {
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

async function loadProfiles(csvPath: string): Promise<UserProfile[]> {
  const content = await fs.readFile(csvPath, 'utf-8');
  const records = parse(content, {
    columns: true,
    skip_empty_lines: true,
    trim: true,
  });

  return records.map((row: any) => {
    const profile: UserProfile = {
      email: row.email,
    };

    if (row.interests) {
      profile.interests = parseArrayValue(row.interests);
    }
    if (row.languages) {
      profile.languages = ProfileWriter.parseLanguages(row.languages);
    }

    return profile;
  });
}

async function enrichViaProfileApi(
  profiles: UserProfile[],
  accessToken: string,
  dryRun: boolean
): Promise<{ successful: number; failed: number }> {
  console.log('\n' + '='.repeat(60));
  console.log('Option A: Profile API Enrichment');
  console.log('='.repeat(60));
  console.log('Writing: languages, interests (no people data labels available)\n');

  const profileWriter = new ProfileWriter(accessToken);
  let successful = 0;
  let failed = 0;

  for (let i = 0; i < profiles.length; i++) {
    const profile = profiles[i];
    const hasProfileData =
      (profile.interests && profile.interests.length > 0) ||
      (profile.languages && profile.languages.length > 0);

    if (!hasProfileData) continue;

    console.log(`[${i + 1}/${profiles.length}] ${profile.email}`);

    if (dryRun) {
      if (profile.interests?.length) console.log(`  [DRY RUN] Would POST ${profile.interests.length} interests`);
      if (profile.languages?.length) console.log(`  [DRY RUN] Would POST ${profile.languages.length} languages`);
      successful++;
      continue;
    }

    let userFailed = false;

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

async function run(): Promise<void> {
  const options = parseArgs();

  const tenantId = process.env.AZURE_TENANT_ID || '';
  const clientId = process.env.AZURE_CLIENT_ID || '';

  if (!tenantId || !clientId) {
    throw new Error('Missing Azure AD configuration in .env file');
  }

  console.log('Option A: Profile API Enrichment\n');
  console.log('Configuration:');
  console.log(`  CSV: ${options.csvPath}`);
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
  const profiles = await loadProfiles(options.csvPath);
  console.log(`Loaded ${profiles.length} profiles`);
  console.log('Field routing:');
  console.log('  Profile API: languages, interests (no people data labels)');
  console.log('');

  // Authenticate (delegated - browser login)
  console.log('Authenticating (delegated - browser login)...');
  const authServer = new BrowserAuthServer({
    tenantId,
    clientId,
    port: parseInt(process.env.AUTH_SERVER_PORT || '5544', 10),
    scopes: ['User.ReadWrite.All', 'People.Read.All', 'offline_access'],
  });
  const authResult = await authServer.authenticate();
  console.log('Authenticated\n');

  const results = await enrichViaProfileApi(
    profiles,
    authResult.accessToken,
    options.dryRun
  );

  // Summary
  console.log('\n' + '='.repeat(60));
  console.log('ENRICHMENT SUMMARY');
  console.log('='.repeat(60));
  console.log(`\nProfile API (languages, interests - profile cards only):`);
  console.log(`  Successful: ${results.successful}`);
  console.log(`  Failed: ${results.failed}`);
  console.log('\nNote: Languages/interests have no people data labels, so they remain');
  console.log('Profile API only (visible on cards, not Copilot-searchable).');
  console.log('='.repeat(60));

  if (results.failed > 0) {
    process.exit(1);
  }
}

run().catch(error => {
  console.error('Fatal error:', error.message);
  process.exit(1);
});
