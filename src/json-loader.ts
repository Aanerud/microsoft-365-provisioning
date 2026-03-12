/**
 * Shared JSON loader for both Option A (provisioning) and Option B (connector enrichment).
 *
 * Flat JSON format — each record has `email` + any combination of properties.
 * Option A extracts flat Entra fields (givenName, jobTitle, etc.).
 * Option B extracts rich profile entities (skills, educationalActivities, patents, etc.).
 */
import fs from 'fs/promises';

// Fields that must be arrays when present
const ARRAY_FIELDS = new Set([
  'skills', 'interests', 'certifications', 'awards', 'projects',
  'educationalActivities', 'languages', 'publications', 'patents',
  'responsibilities',
]);

// Fields that must be strings when present
const STRING_FIELDS = new Set([
  'aboutMe', 'mySite', 'birthday',
]);

export interface ValidationResult {
  valid: boolean;
  errors: string[];
  warnings: string[];
}

/**
 * Load rows from a flat-format JSON file.
 * Each record must have an `email` field.
 */
export async function loadRowsFromJson(jsonPath: string): Promise<any[]> {
  const content = await fs.readFile(jsonPath, 'utf-8');
  let records: any;

  try {
    records = JSON.parse(content);
  } catch (err: any) {
    throw new Error(`Failed to parse JSON file ${jsonPath}: ${err.message}`);
  }

  if (!Array.isArray(records)) {
    throw new Error(`JSON input must be an array of person records, got ${typeof records}`);
  }

  // Validate email on every record
  for (let i = 0; i < records.length; i++) {
    if (!records[i].email) {
      throw new Error(`JSON record at index ${i} missing required "email" field`);
    }
  }

  return records;
}

/**
 * Validate JSON input records for type correctness.
 */
export function validateJsonInput(records: any[]): ValidationResult {
  const errors: string[] = [];
  const warnings: string[] = [];

  const VALID_PROFICIENCY = new Set([
    'elementary', 'conversational', 'limitedWorking',
    'professionalWorking', 'fullProfessional', 'nativeOrBilingual',
  ]);

  const VALID_COLLAB_TAGS = new Set([
    'askMeAbout', 'ableToMentor', 'wantsToLearn', 'wantsToImprove',
  ]);

  for (let i = 0; i < records.length; i++) {
    const r = records[i];
    const prefix = `Record ${i} (${r.email || 'no email'})`;

    if (!r.email) {
      errors.push(`${prefix}: missing "email" field`);
      continue;
    }

    // Array field type checks
    for (const field of ARRAY_FIELDS) {
      if (r[field] !== undefined && !Array.isArray(r[field])) {
        errors.push(`${prefix}: "${field}" must be an array, got ${typeof r[field]}`);
      }
    }

    // String field type checks
    for (const field of STRING_FIELDS) {
      if (r[field] !== undefined && typeof r[field] !== 'string') {
        errors.push(`${prefix}: "${field}" must be a string, got ${typeof r[field]}`);
      }
    }

    // Language proficiency validation
    if (Array.isArray(r.languages)) {
      for (const lang of r.languages) {
        if (typeof lang === 'object' && lang !== null) {
          for (const key of ['reading', 'spoken', 'written']) {
            const val = lang[key];
            if (val && !VALID_PROFICIENCY.has(val)) {
              warnings.push(`${prefix}: unknown proficiency "${val}" for language "${lang.displayName || '?'}".${key}`);
            }
          }
        }
      }
    }

    // Interest collaborationTags validation
    if (Array.isArray(r.interests)) {
      for (const interest of r.interests) {
        if (typeof interest === 'object' && interest !== null && Array.isArray(interest.collaborationTags)) {
          for (const tag of interest.collaborationTags) {
            if (!VALID_COLLAB_TAGS.has(tag)) {
              warnings.push(`${prefix}: unknown collaborationTag "${tag}" for interest "${interest.displayName || '?'}"`);
            }
          }
        }
      }
    }

    // Patent isPending type check
    if (Array.isArray(r.patents)) {
      for (const p of r.patents) {
        if (typeof p === 'object' && p !== null && p.isPending !== undefined && typeof p.isPending !== 'boolean') {
          warnings.push(`${prefix}: patent "isPending" should be boolean, got ${typeof p.isPending}`);
        }
      }
    }
  }

  return { valid: errors.length === 0, errors, warnings };
}
