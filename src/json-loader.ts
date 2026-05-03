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
import { readFileSync } from 'fs';

// Fields that must be arrays when present (after normalization)
const ARRAY_FIELDS = new Set([
  'skills', 'interests', 'certifications', 'awards', 'projects',
  'educationalActivities', 'languages', 'publications', 'patents',
  'responsibilities', 'groups',
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

export type PipelineMode = 'optionA' | 'optionB' | 'groups';

/**
 * Normalize a PascalCase record into the camelCase format the pipeline expects.
 * Runs AFTER deepNormalizeKeys() — all keys are already camelCase.
 *
 * When pipeline is 'optionB', skips mutations that destroy rich entity data
 * needed by the Graph Connector (anniversaries, emails, phones, webAccounts,
 * notes, metadata stripping, fieldsOfStudy coercion, isCurrent injection).
 * Option A and groups pipelines keep all existing behavior.
 */
function normalizePascalRecord(record: any, pipeline: PipelineMode = 'optionA'): any {
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
  if (pipeline !== 'optionB' && Array.isArray(r.phones) && !r.businessPhones) {
    r.businessPhones = r.phones
      .filter((p: any) => p && p.number)
      .map((p: any) => p.number);
  }

  // ── Notes → aboutMe ──
  if (pipeline !== 'optionB' && Array.isArray(r.notes) && r.notes.length > 0 && !r.aboutMe) {
    const first = r.notes[0];
    if (typeof first === 'object' && first?.detail?.content) {
      r.aboutMe = first.detail.content;
    } else if (typeof first === 'string') {
      r.aboutMe = first;
    }
  }

  // ── Anniversaries → employeeHireDate ──
  if (pipeline !== 'optionB' && Array.isArray(r.anniversaries) && !r.employeeHireDate) {
    const workAnniversary = r.anniversaries.find(
      (a: any) => a && (a.type === 'work' || a.type === 'Work')
    );
    if (workAnniversary?.date) {
      r.employeeHireDate = workAnniversary.date;
    }
  }

  // ── Emails ──
  // Extract mail from Emails array if present (Option A needs flat mail field)
  if (pipeline !== 'optionB' && Array.isArray(r.emails) && r.emails.length > 0 && !r.mail) {
    const primary = r.emails.find((e: any) => e?.address);
    if (primary) r.mail = primary.address;
  }

  // ── Projects: fix relatedPerson UPNs (same as positions) ──
  if (Array.isArray(r.projects)) {
    const fixUpn = (person: any) => {
      if (!person || typeof person !== 'object') return;
      const upn = person.userPrincipalName;
      if (upn && upn.includes('@') && !upn.includes(domain)) {
        person.userPrincipalName = upn.split('@')[0] + '@' + domain;
      }
    };
    for (const proj of r.projects) {
      if (Array.isArray(proj.colleagues)) proj.colleagues.forEach(fixUpn);
      if (Array.isArray(proj.sponsors)) proj.sponsors.forEach(fixUpn);
    }
  }

  // ── Education: fieldsOfStudy string → array ──
  if (pipeline !== 'optionB' && Array.isArray(r.educationalActivities)) {
    for (const edu of r.educationalActivities) {
      if (edu?.program?.fieldsOfStudy && typeof edu.program.fieldsOfStudy === 'string') {
        edu.program.fieldsOfStudy = edu.program.fieldsOfStudy
          ? [edu.program.fieldsOfStudy]
          : [];
      }
    }
  }

  // Collection rendering moved out of this function — see renderCollectionProperties().
  // Runs for both PascalCase and camelCase inputs via loadRowsFromJson().

  // ── Positions: pass through as-is ──
  // If the JSON has positions with relatedPerson (manager, colleagues) already structured,
  // they flow through to the connector. We do NOT auto-convert flat fields into relatedPerson —
  // that's the data science team's responsibility to structure correctly.
  // Flat top-level fields (DeploymentManager, Sponsor, etc.) stay as custom connector properties.
  if (Array.isArray(r.positions) && r.positions.length > 0) {
    if (pipeline !== 'optionB') r.positions[0].isCurrent = true;

    // Fix relatedPerson UPNs: rewrite non-tenant domains to USER_DOMAIN
    // The data science team may use company emails (fabrikam.com) but PCP needs tenant UPNs
    const fixUpn = (person: any) => {
      if (!person || typeof person !== 'object') return;
      const upn = person.userPrincipalName;
      if (upn && upn.includes('@') && !upn.includes(domain)) {
        person.userPrincipalName = upn.split('@')[0] + '@' + domain;
      }
    };

    for (const pos of r.positions) {
      if (pos.manager) fixUpn(pos.manager);
      if (Array.isArray(pos.colleagues)) pos.colleagues.forEach(fixUpn);
      if (Array.isArray(pos.sponsors)) pos.sponsors.forEach(fixUpn);
    }
  }

  // ── Strip PCP metadata from profile arrays (Option A only) ──
  // Option B passes through metadata fields — Graph ignores unknown nested fields
  // in JSON-serialized entity values. Stripping is only needed for Option A where
  // flat fields go to the Entra User API.
  if (pipeline !== 'optionB') {
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
  }

  // ── Cleanup: remove fields consumed by normalization ──
  // These were already mapped to their camelCase equivalents above.
  // Without removal they'd be detected as "custom connector properties".
  // Option B keeps rich arrays (emails, phones, anniversaries, etc.) intact.
  const consumedFieldsCommon = [
    'mailNickName', 'firstName', 'lastName', 'address', 'phoneNumber',
  ];
  const consumedFieldsOptionA = [
    'phones', 'emails', 'notes',
    'anniversaries', 'websites', 'webAccounts',
  ];
  for (const f of consumedFieldsCommon) {
    delete r[f];
  }
  if (pipeline !== 'optionB') {
    for (const f of consumedFieldsOptionA) {
      delete r[f];
    }
  }
  // Keep `licenses` — consumed by license-resolver.ts during provisioning

  return r;
}

// ─── Collection Property Renderer ───────────────────────────────────────────

/**
 * Render the `products` field (when present as array-of-objects) into a Path-B
 * safe single-string YAML representation. The schema `description` field tells
 * Copilot what the property means; the value here is pure data.
 *
 * Output (YAML — recommended by Microsoft docs for complex custom properties):
 *
 *   - name: Echo Vault
 *     model: Model A
 *     gtin: 96128715782278
 *   - name: Falcon Ridge
 *     model: Model Y
 *     gtin: 95441323966177
 *
 * Why a single string: connector custom properties without a people-data label
 * use Path B deserialization, which expects `string` only. Sending array or
 * stringCollection crashes the whole connection.
 *
 * Mutates the record in place. Runs for every record regardless of whether the
 * PascalCase normalizer ran — so camelCase inputs also get rendered. Safe to
 * call on already-rendered input (non-array products bypass).
 */
export interface PropertyDefinition {
  name: string;
  type: 'string' | 'collection';
  description?: string;
}

/**
 * Load property definitions from the sidecar properties config file.
 */
export function loadPropertyDefinitions(propertiesPath: string): PropertyDefinition[] {
  try {
    const content = readFileSync(propertiesPath, 'utf-8');
    const defs = JSON.parse(content);
    if (!Array.isArray(defs)) return [];
    return defs;
  } catch {
    return [];
  }
}

/**
 * Extract collection field names (camelCased) from property definitions.
 */
export function getCollectionFieldNames(defs: PropertyDefinition[]): string[] {
  return defs
    .filter(d => d.type === 'collection')
    .map(d => d.name[0].toLowerCase() + d.name.slice(1));
}

/**
 * Build a descriptions map from property definitions (camelCase name → description).
 */
export function buildDescriptionsMap(defs: PropertyDefinition[]): Record<string, string> {
  const map: Record<string, string> = {};
  for (const d of defs) {
    if (d.description) {
      map[d.name[0].toLowerCase() + d.name.slice(1)] = d.description;
    }
  }
  return map;
}

function sanitizeYaml(s: string): string {
  return s.replace(/[\r\n ]/g, ' ');
}

export function renderCollectionProperties(r: any, collectionFields: string[]): void {
  if (!r) return;
  for (const fieldName of collectionFields) {
    renderOneCollection(r, fieldName);
  }
}

function renderOneCollection(r: any, fieldName: string): void {
  const val = r[fieldName];
  if (!Array.isArray(val)) return;

  const items: string[] = [];
  for (const entry of val) {
    if (typeof entry !== 'object' || entry === null) continue;

    const pairs = Object.entries(entry)
      .filter(([, v]) => v !== undefined && v !== null && v !== '')
      .map(([k, v]) => ({ key: k.toLowerCase(), val: sanitizeYaml(String(v)) }));

    if (pairs.length === 0) continue;

    let item = `- ${pairs[0].key}: ${pairs[0].val}`;
    for (let i = 1; i < pairs.length; i++) {
      item += `\n  ${pairs[i].key}: ${pairs[i].val}`;
    }
    items.push(item);
  }

  if (items.length === 0) {
    r[fieldName] = '';
    return;
  }

  let out = items.join('\n');

  // Graph Connector string-property soft cap ~8KB.
  const MAX_BYTES = 8192;
  if (Buffer.byteLength(out, 'utf-8') > MAX_BYTES) {
    let truncated = out.slice(0, 8000);
    const lastNl = truncated.lastIndexOf('\n');
    if (lastNl > 0) truncated = truncated.slice(0, lastNl);
    out = truncated + '\n\n[TRUNCATED]';
    console.warn(
      `[renderCollection] ${fieldName} for ${r.mailNickName || r.email || 'unknown'} exceeded 8KB, truncated`
    );
  }

  r[fieldName] = out;
}

// ─── Public API ─────────────────────────────────────────────────────────────

/**
 * Load rows from a JSON file. Supports both camelCase and PascalCase formats.
 * PascalCase is auto-detected and normalized to camelCase.
 *
 * pipeline controls which mutations run:
 * - 'optionA' (default): full normalization for Entra User API
 * - 'optionB': preserves rich entity data for Graph Connector pass-through
 * - 'groups': same as optionA
 */
export async function loadRowsFromJson(jsonPath: string, pipeline: PipelineMode = 'optionA'): Promise<any[]> {
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

    records = records.map(r => normalizePascalRecord(deepNormalizeKeys(r), pipeline));
  }

  // NOTE: Collection properties (type=collection in properties config) are
  // rendered by the caller (enrich-connector.ts) using renderCollectionProperties().
  // This keeps loadRowsFromJson generic for both Option A and Option B callers.

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
