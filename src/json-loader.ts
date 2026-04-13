/**
 * Shared JSON loader for both Option A (provisioning) and Option B (connector enrichment).
 *
 * Supports two JSON formats:
 * - camelCase (original): flat records with `email` field
 * - PascalCase (data science team): records with `MailNickName`, `DisplayName`, etc.
 *
 * PascalCase format is auto-detected and normalized to camelCase before downstream code sees it.
 */
import fs from 'fs/promises';

// Fields that must be arrays when present (after normalization)
const ARRAY_FIELDS = new Set([
  'skills', 'interests', 'certifications', 'awards', 'projects',
  'educationalActivities', 'languages', 'publications', 'patents',
  'responsibilities',
]);

// Fields that must be strings when present (after normalization)
const STRING_FIELDS = new Set([
  'aboutMe', 'mySite', 'birthday',
]);

// PCP metadata fields to strip (not needed by our pipeline)
const STRIP_FIELDS = new Set([
  'allowedAudiences', 'inference', 'geoCoordinates', 'id',
  'isSearchable', 'createdDateTime', 'lastModifiedDateTime',
  'createdBy', 'lastModifiedBy', 'source', 'sources',
]);

export interface ValidationResult {
  valid: boolean;
  errors: string[];
  warnings: string[];
}

// ─── PascalCase Normalizer ──────────────────────────────────────────────────

/**
 * Detect if records use PascalCase format (data science team output).
 */
function detectPascalCaseFormat(records: any[]): boolean {
  if (records.length === 0) return false;
  const first = records[0];
  return 'MailNickName' in first || 'DisplayName' in first || 'FirstName' in first;
}

/**
 * Convert a single key from PascalCase to camelCase.
 * "DisplayName" → "displayName", "MailNickName" → "mailNickName"
 */
function pascalToCamel(key: string): string {
  if (!key || key.length === 0) return key;
  return key[0].toLowerCase() + key.slice(1);
}

/**
 * Recursively convert all object keys from PascalCase to camelCase.
 * Handles nested objects and arrays. Does NOT touch string values.
 */
function deepNormalizeKeys(obj: any): any {
  if (obj === null || obj === undefined) return obj;
  if (typeof obj !== 'object') return obj;

  if (Array.isArray(obj)) {
    return obj.map(item => deepNormalizeKeys(item));
  }

  const result: any = {};
  for (const [key, value] of Object.entries(obj)) {
    result[pascalToCamel(key)] = deepNormalizeKeys(value);
  }
  return result;
}

/**
 * Normalize a PascalCase record into the camelCase format the pipeline expects.
 * Runs AFTER deepNormalizeKeys() — all keys are already camelCase.
 */
function normalizePascalRecord(record: any): any {
  const domain = process.env.USER_DOMAIN || 'yourdomain.onmicrosoft.com';
  const r = { ...record };

  // ── Identity ──
  // Construct email from MailNickName + domain
  if (r.mailNickName && !r.email) {
    // Guard: if MailNickName already contains @, use it as-is (it's already a full UPN)
    r.email = r.mailNickName.includes('@')
      ? r.mailNickName
      : `${r.mailNickName}@${domain}`;
  }

  // Map firstName/lastName → givenName/surname
  if (r.firstName && !r.givenName) {
    r.givenName = r.firstName;
    delete r.firstName;
  }
  if (r.lastName && !r.surname) {
    r.surname = r.lastName;
    delete r.lastName;
  }

  // Provision.ts requires `name` and `role`
  if (r.displayName && !r.name) {
    r.name = r.displayName;
  }
  if (r.department && !r.role) {
    r.role = r.department.toLowerCase();
  }

  // ── Manager ──
  if ('manager' in r) {
    if (r.manager && typeof r.manager === 'string') {
      r.ManagerEmail = r.manager.includes('@') ? r.manager : `${r.manager}@${domain}`;
    } else {
      r.ManagerEmail = '';
    }
    delete r.manager;
  }

  // ── Address ──
  // Flatten Address object into Entra ID fields
  if (r.address && typeof r.address === 'object' && !Array.isArray(r.address)) {
    if (!r.city) r.city = r.address.city || '';
    if (!r.country) r.country = r.address.countryOrRegion || r.address.country || '';
    if (!r.streetAddress) r.streetAddress = r.address.street || '';
    if (!r.state) r.state = r.address.state || '';
    if (!r.postalCode) r.postalCode = r.address.postalCode || '';
    delete r.address;
  }

  // ── Phones ──
  if (r.phoneNumber && !r.mobilePhone) {
    r.mobilePhone = r.phoneNumber;
    delete r.phoneNumber;
  }
  if (Array.isArray(r.phones) && !r.businessPhones) {
    r.businessPhones = r.phones
      .filter((p: any) => p && p.number)
      .map((p: any) => p.number);
  }

  // ── Notes → aboutMe ──
  if (Array.isArray(r.notes) && r.notes.length > 0 && !r.aboutMe) {
    const first = r.notes[0];
    if (typeof first === 'object' && first?.detail?.content) {
      r.aboutMe = first.detail.content;
    } else if (typeof first === 'string') {
      r.aboutMe = first;
    }
  }

  // ── Anniversaries → employeeHireDate ──
  if (Array.isArray(r.anniversaries) && !r.employeeHireDate) {
    const workAnniversary = r.anniversaries.find(
      (a: any) => a && (a.type === 'work' || a.type === 'Work')
    );
    if (workAnniversary?.date) {
      r.employeeHireDate = workAnniversary.date;
    }
  }

  // ── Emails ──
  // Extract mail from Emails array if present
  if (Array.isArray(r.emails) && r.emails.length > 0 && !r.mail) {
    const primary = r.emails.find((e: any) => e?.address);
    if (primary) r.mail = primary.address;
  }

  // ── Education: fieldsOfStudy string → array ──
  if (Array.isArray(r.educationalActivities)) {
    for (const edu of r.educationalActivities) {
      if (edu?.program?.fieldsOfStudy && typeof edu.program.fieldsOfStudy === 'string') {
        edu.program.fieldsOfStudy = edu.program.fieldsOfStudy
          ? [edu.program.fieldsOfStudy]
          : [];
      }
    }
  }

  // ── Products → stringify for Path B custom property ──
  if (Array.isArray(r.products)) {
    r.products = r.products
      .filter((p: any) => p && p.name)
      .map((p: any) => p.name)
      .join(', ');
  }

  // ── Positions: pass through as-is ──
  // If the JSON has positions with relatedPerson (manager, colleagues) already structured,
  // they flow through to the connector. We do NOT auto-convert flat fields into relatedPerson —
  // that's the data science team's responsibility to structure correctly.
  // Flat top-level fields (DeploymentManager, Sponsor, etc.) stay as custom connector properties.
  if (Array.isArray(r.positions) && r.positions.length > 0) {
    r.positions[0].isCurrent = true;
  }

  // ── Strip PCP metadata from profile arrays ──
  const profileArrays = [
    'skills', 'interests', 'certifications', 'awards', 'projects',
    'educationalActivities', 'languages', 'publications', 'patents',
    'responsibilities', 'addresses', 'phones', 'emails', 'positions',
    'websites', 'webAccounts', 'anniversaries',
  ];
  for (const field of profileArrays) {
    if (Array.isArray(r[field])) {
      r[field] = r[field].map((item: any) => {
        if (typeof item !== 'object' || item === null) return item;
        const cleaned: any = {};
        for (const [k, v] of Object.entries(item)) {
          if (!STRIP_FIELDS.has(k)) cleaned[k] = v;
        }
        return cleaned;
      });
    }
  }

  // ── Cleanup: remove fields consumed by normalization ──
  // These were already mapped to their camelCase equivalents above.
  // Without removal they'd be detected as "custom connector properties".
  const consumedFields = [
    'mailNickName', 'firstName', 'lastName', 'address', 'phoneNumber',
    'addresses', 'phones', 'emails', 'notes',
    'anniversaries', 'websites', 'webAccounts',
    // 'positions' intentionally kept — used by item-ingester for personCurrentPosition with relatedPerson
  ];
  for (const f of consumedFields) {
    delete r[f];
  }
  // Keep `licenses` — consumed by license-resolver.ts during provisioning

  return r;
}

// ─── Public API ─────────────────────────────────────────────────────────────

/**
 * Load rows from a JSON file. Supports both camelCase and PascalCase formats.
 * PascalCase is auto-detected and normalized to camelCase.
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

  // Auto-detect PascalCase format and normalize
  if (detectPascalCaseFormat(records)) {
    console.log(`Detected PascalCase format (${records.length} records), normalizing to camelCase...`);

    records = records.map(r => normalizePascalRecord(deepNormalizeKeys(r)));
  }

  // Validate email on every record (after normalization)
  for (let i = 0; i < records.length; i++) {
    if (!records[i].email) {
      throw new Error(`JSON record at index ${i} missing required "email" field (or MailNickName + USER_DOMAIN)`);
    }
  }

  return records;
}

/**
 * Validate JSON input records for type correctness.
 * Runs after normalization — all fields are camelCase.
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
