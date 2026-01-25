/**
 * Direct Profile Enrichment via People Profile API
 *
 * Uses delegated auth to write to the People Profile API endpoints:
 * - /profile/skills (skills array)
 * - /profile/notes (aboutMe text)
 * - /profile/interests (interests array)
 * - /profile/languages (languages array)
 *
 * This is the CORRECT approach for Copilot searchability - the Profile API
 * stores data with People Data labels that Copilot can search.
 *
 * Based on cocogen retry patterns for reliability.
 */

import { parse } from 'csv-parse/sync';
import fs from 'fs/promises';
import dotenv from 'dotenv';
import { BrowserAuthServer } from './auth/browser-auth-server.js';
import {
  ProfileWriter,
  SkillProficiency,
  PersonInterest,
  PersonNote,
  LanguageProficiency,
} from './profile-writer.js';

dotenv.config();

interface UserProfile {
  email: string;
  skills?: string[];
  aboutMe?: string;
  interests?: string[];
  languages?: LanguageProficiency[];
}

interface EnrichmentResult {
  email: string;
  skills: 'success' | 'partial' | 'failed' | 'skipped';
  aboutMe: 'success' | 'failed' | 'skipped';
  interests: 'success' | 'partial' | 'failed' | 'skipped';
  languages: 'success' | 'partial' | 'failed' | 'skipped';
  errors: string[];
}

class DirectProfileEnricher {
  private profileWriter: ProfileWriter;
  private dryRun: boolean;

  constructor(accessToken: string, dryRun: boolean = false) {
    this.profileWriter = new ProfileWriter(accessToken);
    this.dryRun = dryRun;
  }

  /**
   * Parse CSV and extract profile data
   */
  async loadProfilesFromCsv(csvPath: string): Promise<UserProfile[]> {
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

      // Parse skills
      if (row.skills) {
        profile.skills = this.parseArrayValue(row.skills);
      }

      // Parse aboutMe
      if (row.aboutMe) {
        profile.aboutMe = row.aboutMe;
      }

      // Parse interests
      if (row.interests) {
        profile.interests = this.parseArrayValue(row.interests);
      }

      // Parse languages
      if (row.languages) {
        profile.languages = ProfileWriter.parseLanguages(row.languages);
      }

      return profile;
    });
  }

  /**
   * Parse array value from CSV (handles JSON arrays and comma-separated)
   */
  private parseArrayValue(value: string): string[] {
    if (!value || value.trim() === '') return [];

    try {
      // Try JSON array (with single quotes converted)
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
   * Enrich a single user's profile using the People Profile API
   */
  async enrichUser(profile: UserProfile): Promise<EnrichmentResult> {
    const result: EnrichmentResult = {
      email: profile.email,
      skills: 'skipped',
      aboutMe: 'skipped',
      interests: 'skipped',
      languages: 'skipped',
      errors: [],
    };

    if (this.dryRun) {
      if (profile.skills && profile.skills.length > 0) {
        console.log(`  [DRY RUN] Would POST ${profile.skills.length} skills to /profile/skills`);
        result.skills = 'success';
      }
      if (profile.aboutMe) {
        console.log(`  [DRY RUN] Would POST aboutMe to /profile/notes`);
        result.aboutMe = 'success';
      }
      if (profile.interests && profile.interests.length > 0) {
        console.log(`  [DRY RUN] Would POST ${profile.interests.length} interests to /profile/interests`);
        result.interests = 'success';
      }
      if (profile.languages && profile.languages.length > 0) {
        console.log(`  [DRY RUN] Would POST ${profile.languages.length} languages to /profile/languages`);
        result.languages = 'success';
      }
      return result;
    }

    // Write skills via Profile API
    if (profile.skills && profile.skills.length > 0) {
      try {
        // NOTE: We omit proficiency due to MS Graph API bug with certain values
        const skillObjects: SkillProficiency[] = profile.skills.map(s => ({
          displayName: s,
          allowedAudiences: 'organization',
        }));

        const skillResult = await this.profileWriter.writeSkills(profile.email, skillObjects);

        if (skillResult.failed === 0) {
          result.skills = 'success';
          console.log(`  ✓ skills (${skillResult.successful} items via /profile/skills)`);
        } else if (skillResult.successful > 0) {
          result.skills = 'partial';
          console.log(`  ⚠ skills (${skillResult.successful}/${profile.skills.length} succeeded)`);
          result.errors.push(...skillResult.errors);
        } else {
          result.skills = 'failed';
          result.errors.push(...skillResult.errors);
          console.log(`  ✗ skills: all failed`);
        }
      } catch (error: any) {
        result.skills = 'failed';
        result.errors.push(`skills: ${error.message}`);
        console.log(`  ✗ skills: ${error.message}`);
      }
    }

    // Write aboutMe via Profile API (notes endpoint)
    if (profile.aboutMe) {
      try {
        const noteObject: PersonNote[] = [{
          detail: profile.aboutMe,
          displayName: 'About Me',
          allowedAudiences: 'organization',
        }];

        const noteResult = await this.profileWriter.writeNotes(profile.email, noteObject);

        if (noteResult.failed === 0 && noteResult.successful > 0) {
          result.aboutMe = 'success';
          console.log(`  ✓ aboutMe (via /profile/notes)`);
        } else {
          result.aboutMe = 'failed';
          result.errors.push(...noteResult.errors);
          console.log(`  ✗ aboutMe: failed`);
        }
      } catch (error: any) {
        result.aboutMe = 'failed';
        result.errors.push(`aboutMe: ${error.message}`);
        console.log(`  ✗ aboutMe: ${error.message}`);
      }
    }

    // Write interests via Profile API
    if (profile.interests && profile.interests.length > 0) {
      try {
        const interestObjects: PersonInterest[] = profile.interests.map(i => ({
          displayName: i,
          allowedAudiences: 'organization',
        }));

        const interestResult = await this.profileWriter.writeInterests(profile.email, interestObjects);

        if (interestResult.failed === 0) {
          result.interests = 'success';
          console.log(`  ✓ interests (${interestResult.successful} items via /profile/interests)`);
        } else if (interestResult.successful > 0) {
          result.interests = 'partial';
          console.log(`  ⚠ interests (${interestResult.successful}/${profile.interests.length} succeeded)`);
          result.errors.push(...interestResult.errors);
        } else {
          result.interests = 'failed';
          result.errors.push(...interestResult.errors);
          console.log(`  ✗ interests: all failed`);
        }
      } catch (error: any) {
        result.interests = 'failed';
        result.errors.push(`interests: ${error.message}`);
        console.log(`  ✗ interests: ${error.message}`);
      }
    }

    // Write languages via Profile API
    if (profile.languages && profile.languages.length > 0) {
      try {
        const langResult = await this.profileWriter.writeLanguages(profile.email, profile.languages);

        if (langResult.failed === 0) {
          result.languages = 'success';
          console.log(`  ✓ languages (${langResult.successful} items via /profile/languages)`);
        } else if (langResult.successful > 0) {
          result.languages = 'partial';
          console.log(`  ⚠ languages (${langResult.successful}/${profile.languages.length} succeeded)`);
          result.errors.push(...langResult.errors);
        } else {
          result.languages = 'failed';
          result.errors.push(...langResult.errors);
          console.log(`  ✗ languages: all failed`);
        }
      } catch (error: any) {
        result.languages = 'failed';
        result.errors.push(`languages: ${error.message}`);
        console.log(`  ✗ languages: ${error.message}`);
      }
    }

    return result;
  }

  /**
   * Enrich all users from CSV
   */
  async enrichAll(profiles: UserProfile[]): Promise<{
    successful: number;
    partial: number;
    failed: number;
    results: EnrichmentResult[];
  }> {
    const results: EnrichmentResult[] = [];
    let successful = 0;
    let partial = 0;
    let failed = 0;

    for (let i = 0; i < profiles.length; i++) {
      const profile = profiles[i];
      console.log(`\n[${i + 1}/${profiles.length}] ${profile.email}`);

      const result = await this.enrichUser(profile);
      results.push(result);

      // Count results
      const statuses = [result.skills, result.aboutMe, result.interests, result.languages];
      const successCount = statuses.filter(s => s === 'success').length;
      const failCount = statuses.filter(s => s === 'failed').length;
      const partialCount = statuses.filter(s => s === 'partial').length;
      const attemptedCount = statuses.filter(s => s !== 'skipped').length;

      if (attemptedCount === 0) {
        // Nothing to do for this user
      } else if (failCount === 0 && partialCount === 0) {
        successful++;
      } else if (successCount > 0 || partialCount > 0) {
        partial++;
      } else {
        failed++;
      }

      // Small delay between users to avoid rate limiting
      if (i < profiles.length - 1) {
        await new Promise(resolve => setTimeout(resolve, 100));
      }
    }

    return { successful, partial, failed, results };
  }
}

async function main() {
  console.log('Profile Enrichment via People Profile API\n');

  // Parse arguments
  const args = process.argv.slice(2);
  let csvPath = 'config/textcraft-europe.csv';
  let dryRun = false;

  for (let i = 0; i < args.length; i++) {
    if (args[i] === '--csv' && args[i + 1]) {
      csvPath = args[++i];
    } else if (args[i] === '--dry-run') {
      dryRun = true;
    }
  }

  console.log('Configuration:');
  console.log(`  CSV: ${csvPath}`);
  console.log(`  Dry Run: ${dryRun ? 'Yes' : 'No'}`);
  console.log('');

  // Check CSV exists
  try {
    await fs.access(csvPath);
  } catch {
    console.error(`Error: CSV file not found: ${csvPath}`);
    process.exit(1);
  }

  // Get delegated auth token (browser login)
  console.log('Authenticating (delegated - browser login required)...');
  const authServer = new BrowserAuthServer({
    tenantId: process.env.AZURE_TENANT_ID!,
    clientId: process.env.AZURE_CLIENT_ID!,
    port: parseInt(process.env.AUTH_SERVER_PORT || '5544', 10),
    scopes: [
      'User.ReadWrite.All',
      'People.Read.All',
      'offline_access',
    ],
  });
  const authResult = await authServer.authenticate();
  const accessToken = authResult.accessToken;
  console.log('Authenticated\n');

  // Create enricher
  const enricher = new DirectProfileEnricher(accessToken, dryRun);

  // Load profiles
  console.log('Loading profiles from CSV...');
  const profiles = await enricher.loadProfilesFromCsv(csvPath);
  console.log(`Loaded ${profiles.length} profiles\n`);

  // Filter to profiles with enrichment data
  const profilesWithData = profiles.filter(p =>
    (p.skills && p.skills.length > 0) ||
    p.aboutMe ||
    (p.interests && p.interests.length > 0) ||
    (p.languages && p.languages.length > 0)
  );
  console.log(`Enriching ${profilesWithData.length} profiles with data...\n`);

  if (profilesWithData.length === 0) {
    console.log('No profiles have enrichment data (skills, aboutMe, interests, languages).');
    process.exit(0);
  }

  // Enrich all
  const { successful, partial, failed, results } = await enricher.enrichAll(profilesWithData);

  // Summary
  console.log('\n' + '='.repeat(60));
  console.log('Profile Enrichment Summary (via People Profile API)');
  console.log('='.repeat(60));
  console.log(`Fully successful: ${successful}`);
  console.log(`Partially successful: ${partial}`);
  console.log(`Failed: ${failed}`);
  console.log('='.repeat(60));

  // Show failures
  const failures = results.filter(r => r.errors.length > 0);
  if (failures.length > 0) {
    console.log('\nFailed enrichments:');
    for (const f of failures) {
      console.log(`  - ${f.email}:`);
      for (const err of f.errors) {
        console.log(`      ${err}`);
      }
    }
  }

  process.exit(failed > 0 ? 1 : 0);
}

main().catch(error => {
  console.error('Fatal error:', error.message);
  process.exit(1);
});
