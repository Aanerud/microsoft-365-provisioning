/**
 * Profile Writer - Writes data to People Profile API
 *
 * This module handles writing to the Microsoft Graph People Profile API endpoints
 * which require DELEGATED authentication (user sign-in).
 *
 * Supports writing:
 * - Languages (/profile/languages)
 * - Skills (/profile/skills)
 * - Notes/AboutMe (/profile/notes)
 * - Interests (/profile/interests)
 *
 * This is the correct approach for Copilot searchability - the Profile API
 * stores data with People Data labels that Copilot can search.
 */

import { Client } from '@microsoft/microsoft-graph-client';

/**
 * Language proficiency levels supported by Microsoft Graph Profile API
 */
export type ProficiencyLevel =
  | 'nativeOrBilingual'
  | 'fullProfessional'
  | 'professionalWorking'
  | 'conversational'
  | 'elementary';

/**
 * Language proficiency as expected by the Profile API
 */
export interface LanguageProficiency {
  /** Language name (e.g., "Polish", "English", "German") */
  displayName: string;
  /** Optional BCP-47 language tag (e.g., "pl-PL", "en-US") */
  tag?: string;
  /** Spoken proficiency level */
  spoken: ProficiencyLevel;
  /** Written proficiency level */
  written: ProficiencyLevel;
  /** Reading proficiency level */
  reading: ProficiencyLevel;
  /** Who can see this information */
  allowedAudiences?: 'organization' | 'everyone' | 'me';
}

/**
 * Skill proficiency levels supported by Microsoft Graph Profile API
 */
export type SkillProficiencyLevel =
  | 'elementary'
  | 'limitedWorking'
  | 'generalProfessional'
  | 'advancedProfessional'
  | 'expert';

/**
 * Skill as expected by the Profile API
 */
export interface SkillProficiency {
  /** Skill name (e.g., "TypeScript", "Project Management") */
  displayName: string;
  /** Proficiency level */
  proficiency?: SkillProficiencyLevel;
  /** Who can see this information */
  allowedAudiences?: 'organization' | 'everyone' | 'me';
}

/**
 * Person interest as expected by the Profile API
 */
export interface PersonInterest {
  /** Interest name (e.g., "Machine Learning", "Photography") */
  displayName: string;
  /** Who can see this information */
  allowedAudiences?: 'organization' | 'everyone' | 'me';
}

/**
 * Person note/about me as expected by the Profile API
 */
export interface PersonNote {
  /** Note content (aboutMe text) */
  detail: string;
  /** Display name for the note */
  displayName?: string;
  /** Who can see this information */
  allowedAudiences?: 'organization' | 'everyone' | 'me';
}

/**
 * Mapping from CSV proficiency text to Profile API values
 */
const PROFICIENCY_MAP: Record<string, ProficiencyLevel> = {
  // Native level
  native: 'nativeOrBilingual',
  bilingual: 'nativeOrBilingual',
  'native or bilingual': 'nativeOrBilingual',

  // Fluent/Full professional
  fluent: 'fullProfessional',
  'full professional': 'fullProfessional',

  // Professional working
  professional: 'professionalWorking',
  'professional working': 'professionalWorking',
  advanced: 'professionalWorking',

  // Conversational
  conversational: 'conversational',
  intermediate: 'conversational',

  // Elementary/Basic
  elementary: 'elementary',
  basic: 'elementary',
  beginner: 'elementary',
};

/**
 * Common language name to BCP-47 tag mapping
 */
const LANGUAGE_TAG_MAP: Record<string, string> = {
  polish: 'pl-PL',
  english: 'en-US',
  german: 'de-DE',
  french: 'fr-FR',
  spanish: 'es-ES',
  italian: 'it-IT',
  portuguese: 'pt-PT',
  dutch: 'nl-NL',
  russian: 'ru-RU',
  chinese: 'zh-CN',
  japanese: 'ja-JP',
  korean: 'ko-KR',
  arabic: 'ar-SA',
  hebrew: 'he-IL',
  hindi: 'hi-IN',
  turkish: 'tr-TR',
  greek: 'el-GR',
  czech: 'cs-CZ',
  swedish: 'sv-SE',
  norwegian: 'nb-NO',
  danish: 'da-DK',
  finnish: 'fi-FI',
  hungarian: 'hu-HU',
  romanian: 'ro-RO',
  ukrainian: 'uk-UA',
};

/**
 * Profile Writer for People Profile API
 *
 * Handles writing profile data that doesn't have Graph Connector labels.
 * Uses delegated authentication (user must be signed in).
 */
export class ProfileWriter {
  private client: Client;

  /**
   * Create a ProfileWriter instance
   * @param accessToken - Delegated access token from user sign-in
   */
  constructor(accessToken: string) {
    this.client = Client.init({
      authProvider: (done) => {
        done(null, accessToken);
      },
      defaultVersion: 'beta', // Profile API is beta-only
    });
  }

  /**
   * Write languages to a user's profile
   *
   * @param userId - User ID or UPN
   * @param languages - Array of language proficiency objects
   * @returns Count of successfully written languages
   */
  async writeLanguages(userId: string, languages: LanguageProficiency[]): Promise<{
    successful: number;
    failed: number;
    errors: string[];
  }> {
    const result = {
      successful: 0,
      failed: 0,
      errors: [] as string[],
    };

    if (!languages || languages.length === 0) {
      return result;
    }

    // First, get existing languages to avoid duplicates
    let existingLanguages: string[] = [];
    try {
      const existing = await this.client.api(`/users/${userId}/profile/languages`).get();
      existingLanguages = (existing.value || []).map((l: any) => l.displayName.toLowerCase());
    } catch (error: any) {
      // If profile doesn't exist yet, that's fine
      if (error.statusCode !== 404) {
        console.warn(`⚠ Could not fetch existing languages: ${error.message}`);
      }
    }

    // Write each language
    for (const lang of languages) {
      // Skip if language already exists
      if (existingLanguages.includes(lang.displayName.toLowerCase())) {
        console.log(`  ⏭ Language already exists: ${lang.displayName}`);
        result.successful++; // Count as success since it's already there
        continue;
      }

      try {
        const payload = {
          displayName: lang.displayName,
          tag: lang.tag,
          spoken: lang.spoken,
          written: lang.written,
          reading: lang.reading,
          allowedAudiences: lang.allowedAudiences || 'organization',
        };

        await this.client.api(`/users/${userId}/profile/languages`).post(payload);
        console.log(`  ✓ Added language: ${lang.displayName} (${lang.spoken})`);
        result.successful++;
      } catch (error: any) {
        const errorMsg = `Failed to add ${lang.displayName}: ${error.message}`;
        result.errors.push(errorMsg);
        result.failed++;
        console.warn(`  ✗ ${errorMsg}`);
      }
    }

    return result;
  }

  /**
   * Write skills to a user's profile
   *
   * @param userId - User ID or UPN
   * @param skills - Array of skill proficiency objects
   * @returns Count of successfully written skills
   */
  async writeSkills(userId: string, skills: SkillProficiency[]): Promise<{
    successful: number;
    failed: number;
    errors: string[];
  }> {
    const result = {
      successful: 0,
      failed: 0,
      errors: [] as string[],
    };

    if (!skills || skills.length === 0) {
      return result;
    }

    // First, get existing skills to avoid duplicates
    let existingSkills: string[] = [];
    try {
      const existing = await this.client.api(`/users/${userId}/profile/skills`).get();
      existingSkills = (existing.value || []).map((s: any) => s.displayName.toLowerCase());
    } catch (error: any) {
      // If profile doesn't exist yet, that's fine
      if (error.statusCode !== 404) {
        console.warn(`⚠ Could not fetch existing skills: ${error.message}`);
      }
    }

    // Write each skill
    for (const skill of skills) {
      // Skip if skill already exists
      if (existingSkills.includes(skill.displayName.toLowerCase())) {
        console.log(`  ⏭ Skill already exists: ${skill.displayName}`);
        result.successful++; // Count as success since it's already there
        continue;
      }

      try {
        // NOTE: Microsoft Graph API has a bug where certain proficiency values
        // (generalProfessional, advancedProfessional) cause deserialization errors.
        // Working values are: elementary, limitedWorking, expert
        // For now, we omit proficiency to avoid the bug - skills still work for search.
        const payload: any = {
          displayName: skill.displayName,
          allowedAudiences: skill.allowedAudiences || 'organization',
        };

        // Only include proficiency if it's a known working value
        const workingProficiencies = ['elementary', 'limitedWorking', 'expert'];
        if (skill.proficiency && workingProficiencies.includes(skill.proficiency)) {
          payload.proficiency = skill.proficiency;
        }

        await this.client.api(`/users/${userId}/profile/skills`).post(payload);
        console.log(`  ✓ Added skill: ${skill.displayName}`);
        result.successful++;
      } catch (error: any) {
        const errorMsg = `Failed to add skill ${skill.displayName}: ${error.message}`;
        result.errors.push(errorMsg);
        result.failed++;
        console.warn(`  ✗ ${errorMsg}`);
      }
    }

    return result;
  }

  /**
   * Write interests to a user's profile
   *
   * @param userId - User ID or UPN
   * @param interests - Array of interest objects
   * @returns Count of successfully written interests
   */
  async writeInterests(userId: string, interests: PersonInterest[]): Promise<{
    successful: number;
    failed: number;
    errors: string[];
  }> {
    const result = {
      successful: 0,
      failed: 0,
      errors: [] as string[],
    };

    if (!interests || interests.length === 0) {
      return result;
    }

    // First, get existing interests to avoid duplicates
    let existingInterests: string[] = [];
    try {
      const existing = await this.client.api(`/users/${userId}/profile/interests`).get();
      existingInterests = (existing.value || []).map((i: any) => i.displayName.toLowerCase());
    } catch (error: any) {
      // If profile doesn't exist yet, that's fine
      if (error.statusCode !== 404) {
        console.warn(`⚠ Could not fetch existing interests: ${error.message}`);
      }
    }

    // Write each interest
    for (const interest of interests) {
      // Skip if interest already exists
      if (existingInterests.includes(interest.displayName.toLowerCase())) {
        console.log(`  ⏭ Interest already exists: ${interest.displayName}`);
        result.successful++; // Count as success since it's already there
        continue;
      }

      try {
        const payload = {
          displayName: interest.displayName,
          allowedAudiences: interest.allowedAudiences || 'organization',
        };

        await this.client.api(`/users/${userId}/profile/interests`).post(payload);
        console.log(`  ✓ Added interest: ${interest.displayName}`);
        result.successful++;
      } catch (error: any) {
        const errorMsg = `Failed to add interest ${interest.displayName}: ${error.message}`;
        result.errors.push(errorMsg);
        result.failed++;
        console.warn(`  ✗ ${errorMsg}`);
      }
    }

    return result;
  }

  /**
   * Write notes/aboutMe to a user's profile
   *
   * @param userId - User ID or UPN
   * @param notes - Array of note objects (typically just one for aboutMe)
   * @returns Count of successfully written notes
   */
  async writeNotes(userId: string, notes: PersonNote[]): Promise<{
    successful: number;
    failed: number;
    errors: string[];
  }> {
    const result = {
      successful: 0,
      failed: 0,
      errors: [] as string[],
    };

    if (!notes || notes.length === 0) {
      return result;
    }

    // Write each note
    for (const note of notes) {
      try {
        // The Profile API requires detail to be an itemBody object, not a string
        const payload = {
          detail: {
            contentType: 'text',
            content: note.detail,
          },
          displayName: note.displayName || 'About Me',
          allowedAudiences: note.allowedAudiences || 'organization',
        };

        await this.client.api(`/users/${userId}/profile/notes`).post(payload);
        console.log(`  ✓ Added note: ${note.displayName || 'About Me'}`);
        result.successful++;
      } catch (error: any) {
        const errorMsg = `Failed to add note: ${error.message}`;
        result.errors.push(errorMsg);
        result.failed++;
        console.warn(`  ✗ ${errorMsg}`);
      }
    }

    return result;
  }

  /**
   * Parse skills from CSV format
   *
   * Supports formats:
   * - "['TypeScript','Python','Project Management']"
   * - "TypeScript, Python, Project Management"
   *
   * @param csvValue - Raw CSV value for skills column
   * @returns Array of parsed SkillProficiency objects
   */
  static parseSkills(csvValue: string): SkillProficiency[] {
    if (!csvValue || csvValue.trim() === '') {
      return [];
    }

    let rawEntries: string[] = [];

    // Try parsing as JSON array first
    try {
      const normalized = csvValue.replace(/'/g, '"');
      const parsed = JSON.parse(normalized);
      if (Array.isArray(parsed)) {
        rawEntries = parsed;
      }
    } catch {
      // Not JSON, try comma-separated
      rawEntries = csvValue.split(',').map(s => s.trim());
    }

    return rawEntries
      .filter(s => s && s.trim() !== '')
      .map(skillName => ({
        displayName: skillName.trim(),
        proficiency: 'generalProfessional' as SkillProficiencyLevel,
        allowedAudiences: 'organization' as const,
      }));
  }

  /**
   * Parse interests from CSV format
   *
   * Supports formats:
   * - "['AI','Photography','Music']"
   * - "AI, Photography, Music"
   *
   * @param csvValue - Raw CSV value for interests column
   * @returns Array of parsed PersonInterest objects
   */
  static parseInterests(csvValue: string): PersonInterest[] {
    if (!csvValue || csvValue.trim() === '') {
      return [];
    }

    let rawEntries: string[] = [];

    // Try parsing as JSON array first
    try {
      const normalized = csvValue.replace(/'/g, '"');
      const parsed = JSON.parse(normalized);
      if (Array.isArray(parsed)) {
        rawEntries = parsed;
      }
    } catch {
      // Not JSON, try comma-separated
      rawEntries = csvValue.split(',').map(s => s.trim());
    }

    return rawEntries
      .filter(s => s && s.trim() !== '')
      .map(interestName => ({
        displayName: interestName.trim(),
        allowedAudiences: 'organization' as const,
      }));
  }

  /**
   * Parse languages from CSV format
   *
   * Supports formats:
   * - "['Polish (Native)','English (Professional)','German (Conversational)']"
   * - "Polish (Native), English (Professional), German (Conversational)"
   * - "Polish:Native, English:Professional"
   *
   * @param csvValue - Raw CSV value for languages column
   * @returns Array of parsed LanguageProficiency objects
   */
  static parseLanguages(csvValue: string): LanguageProficiency[] {
    if (!csvValue || csvValue.trim() === '') {
      return [];
    }

    const languages: LanguageProficiency[] = [];
    let rawEntries: string[] = [];

    // Try parsing as JSON array first
    try {
      // Handle single-quoted JSON: ['Polish (Native)','English (Professional)']
      const normalized = csvValue.replace(/'/g, '"');
      const parsed = JSON.parse(normalized);
      if (Array.isArray(parsed)) {
        rawEntries = parsed;
      }
    } catch {
      // Not JSON, try comma-separated
      rawEntries = csvValue.split(',').map(s => s.trim());
    }

    // Parse each entry
    for (const entry of rawEntries) {
      const parsed = ProfileWriter.parseLanguageEntry(entry);
      if (parsed) {
        languages.push(parsed);
      }
    }

    return languages;
  }

  /**
   * Parse a single language entry
   *
   * Formats supported:
   * - "Polish (Native)"
   * - "Polish:Native"
   * - "Polish - Native"
   * - "Polish Native"
   *
   * @param entry - Single language entry string
   * @returns Parsed LanguageProficiency or null if invalid
   */
  private static parseLanguageEntry(entry: string): LanguageProficiency | null {
    if (!entry || entry.trim() === '') {
      return null;
    }

    // Try different patterns
    let languageName = '';
    let proficiencyText = '';

    // Pattern 1: "Polish (Native)" or "Polish (Native or Bilingual)"
    const parenMatch = entry.match(/^([^(]+)\s*\(([^)]+)\)$/);
    if (parenMatch) {
      languageName = parenMatch[1].trim();
      proficiencyText = parenMatch[2].trim();
    }
    // Pattern 2: "Polish:Native"
    else if (entry.includes(':')) {
      const parts = entry.split(':');
      languageName = parts[0].trim();
      proficiencyText = parts.slice(1).join(':').trim();
    }
    // Pattern 3: "Polish - Native"
    else if (entry.includes(' - ')) {
      const parts = entry.split(' - ');
      languageName = parts[0].trim();
      proficiencyText = parts.slice(1).join(' - ').trim();
    }
    // Pattern 4: Just language name, default to conversational
    else {
      languageName = entry.trim();
      proficiencyText = 'conversational';
    }

    if (!languageName) {
      return null;
    }

    // Map proficiency text to API value
    const proficiencyKey = proficiencyText.toLowerCase();
    const proficiency = PROFICIENCY_MAP[proficiencyKey] || 'conversational';

    // Get language tag if available
    const tag = LANGUAGE_TAG_MAP[languageName.toLowerCase()];

    return {
      displayName: languageName,
      tag,
      spoken: proficiency,
      written: proficiency,
      reading: proficiency,
      allowedAudiences: 'organization',
    };
  }

  /**
   * Get all supported proficiency level mappings (for documentation)
   */
  static getProficiencyMappings(): Record<string, ProficiencyLevel> {
    return { ...PROFICIENCY_MAP };
  }

  /**
   * Get all supported language tag mappings (for documentation)
   */
  static getLanguageTagMappings(): Record<string, string> {
    return { ...LANGUAGE_TAG_MAP };
  }
}
