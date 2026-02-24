/**
 * Microsoft Graph User Property Schema
 *
 * Comprehensive schema for all writable user properties available in Microsoft Graph Beta API.
 * Based on: https://learn.microsoft.com/en-us/graph/api/resources/user?view=graph-rest-beta
 *
 * This schema enables the provisioning tool to:
 * 1. Support all 50+ standard Microsoft Graph user properties
 * 2. Automatically detect custom properties (not in this schema)
 * 3. Validate property types and constraints
 * 4. Map CSV columns to Graph API property names
 */

export type PropertyType = 'string' | 'boolean' | 'number' | 'date' | 'array' | 'object';

export type PropertyCategory =
  | 'basic'
  | 'contact'
  | 'address'
  | 'job'
  | 'identity'
  | 'security'
  | 'personal'
  | 'preferences'
  | 'onpremises'
  | 'legal';

export interface PropertyMetadata {
  /** Property name as it appears in CSV columns */
  name: string;
  /** Data type of the property */
  type: PropertyType;
  /** Whether this property can be written */
  writable: boolean;
  /** Whether this property is required for user creation */
  required?: boolean;
  /** Maximum length for string properties */
  maxLength?: number;
  /** Property category for organization */
  category: PropertyCategory;
  /** Actual property name in Microsoft Graph API */
  graphPath: string;
  /** Whether this property is only available in beta endpoint */
  betaOnly?: boolean;
  /** Property description */
  description?: string;
  /** Which option handles this property: Option A (standard Entra ID) or Option B (Graph Connectors) */
  handledBy: 'optionA' | 'optionB';
  /** For Option B properties: the official Microsoft Graph people data label (if available) */
  peopleDataLabel?: string | null;
  /** For Profile API properties: the endpoint path (e.g., '/profile/languages') */
  profileApiEndpoint?: string;
}

/**
 * Complete schema of all writable user properties in Microsoft Graph
 */
export const USER_PROPERTY_SCHEMA: PropertyMetadata[] = [
  // ===== BASIC INFO =====
  {
    name: 'displayName',
    type: 'string',
    writable: true,
    required: true,
    maxLength: 256,
    category: 'basic',
    graphPath: 'displayName',
    description: 'The name displayed in the address book',
    handledBy: 'optionA',
  },
  {
    name: 'givenName',
    type: 'string',
    writable: true,
    maxLength: 64,
    category: 'basic',
    graphPath: 'givenName',
    description: 'First name of the user',
    handledBy: 'optionA',
  },
  {
    name: 'surname',
    type: 'string',
    writable: true,
    maxLength: 64,
    category: 'basic',
    graphPath: 'surname',
    description: 'Last name of the user',
    handledBy: 'optionA',
  },
  {
    name: 'aboutMe',
    type: 'string',
    writable: true,
    category: 'personal',
    graphPath: 'aboutMe',
    description: 'A freeform text entry field for the user to describe themselves',
    handledBy: 'optionB',
    peopleDataLabel: 'personNote',
    profileApiEndpoint: '/profile/notes',
  },
  {
    name: 'accountEnabled',
    type: 'boolean',
    writable: true,
    required: true,
    category: 'basic',
    graphPath: 'accountEnabled',
    description: 'Whether the account is enabled',
    handledBy: 'optionA',
  },

  // ===== CONTACT =====
  {
    name: 'mail',
    type: 'string',
    writable: true,
    category: 'contact',
    graphPath: 'mail',
    description: 'The SMTP address for the user',
    handledBy: 'optionA',
  },
  {
    name: 'mailNickname',
    type: 'string',
    writable: true,
    required: true,
    maxLength: 64,
    category: 'contact',
    graphPath: 'mailNickname',
    description: 'The mail alias for the user',
    handledBy: 'optionA',
  },
  {
    name: 'mobilePhone',
    type: 'string',
    writable: true,
    category: 'contact',
    graphPath: 'mobilePhone',
    description: 'Primary cellular telephone number',
    handledBy: 'optionA',
  },
  {
    name: 'businessPhones',
    type: 'array',
    writable: true,
    category: 'contact',
    graphPath: 'businessPhones',
    description: 'Telephone numbers for the user',
    handledBy: 'optionA',
  },
  {
    name: 'otherMails',
    type: 'array',
    writable: true,
    category: 'contact',
    graphPath: 'otherMails',
    description: 'Additional email addresses for the user',
    handledBy: 'optionA',
  },
  {
    name: 'faxNumber',
    type: 'string',
    writable: true,
    category: 'contact',
    graphPath: 'faxNumber',
    description: 'Fax number of the user',
    handledBy: 'optionA',
  },

  // ===== ADDRESS =====
  {
    name: 'city',
    type: 'string',
    writable: true,
    maxLength: 128,
    category: 'address',
    graphPath: 'city',
    description: 'The city in which the user is located',
    handledBy: 'optionA',
  },
  {
    name: 'state',
    type: 'string',
    writable: true,
    maxLength: 128,
    category: 'address',
    graphPath: 'state',
    description: 'The state or province in which the user is located',
    handledBy: 'optionA',
  },
  {
    name: 'country',
    type: 'string',
    writable: true,
    maxLength: 128,
    category: 'address',
    graphPath: 'country',
    description: 'The country/region in which the user is located',
    handledBy: 'optionA',
  },
  {
    name: 'postalCode',
    type: 'string',
    writable: true,
    maxLength: 40,
    category: 'address',
    graphPath: 'postalCode',
    description: 'The postal code for the user',
    handledBy: 'optionA',
  },
  {
    name: 'streetAddress',
    type: 'string',
    writable: true,
    maxLength: 1024,
    category: 'address',
    graphPath: 'streetAddress',
    description: 'The street address of the user',
    handledBy: 'optionA',
  },
  {
    name: 'officeLocation',
    type: 'string',
    writable: true,
    maxLength: 128,
    category: 'address',
    graphPath: 'officeLocation',
    betaOnly: false,
    description: 'The office location in the user place of business',
    handledBy: 'optionA',
  },

  // ===== JOB INFO =====
  {
    name: 'jobTitle',
    type: 'string',
    writable: true,
    maxLength: 128,
    category: 'job',
    graphPath: 'jobTitle',
    description: "The user's job title",
    handledBy: 'optionA',
  },
  {
    name: 'department',
    type: 'string',
    writable: true,
    maxLength: 64,
    category: 'job',
    graphPath: 'department',
    description: 'The name for the department in which the user works',
    handledBy: 'optionA',
  },
  {
    name: 'companyName',
    type: 'string',
    writable: true,
    maxLength: 64,
    category: 'job',
    graphPath: 'companyName',
    betaOnly: false,
    description: 'The company name which the user is associated',
    handledBy: 'optionA',
  },
  {
    name: 'employeeId',
    type: 'string',
    writable: true,
    maxLength: 16,
    category: 'job',
    graphPath: 'employeeId',
    description: 'The employee identifier assigned to the user',
    handledBy: 'optionA',
  },
  {
    name: 'employeeType',
    type: 'string',
    writable: true,
    category: 'job',
    graphPath: 'employeeType',
    betaOnly: false,
    description: 'Captures enterprise worker type (Employee, Contractor, etc.)',
    handledBy: 'optionA',
  },
  {
    name: 'employeeHireDate',
    type: 'date',
    writable: true,
    category: 'job',
    graphPath: 'employeeHireDate',
    betaOnly: true,
    description: 'The hire date of the user',
    handledBy: 'optionA',
  },
  {
    name: 'employeeLeaveDateTime',
    type: 'date',
    writable: true,
    category: 'job',
    graphPath: 'employeeLeaveDateTime',
    betaOnly: true,
    description: 'The date and time when the user left or will leave the organization',
    handledBy: 'optionA',
  },
  {
    name: 'hireDate',
    type: 'date',
    writable: true,
    category: 'job',
    graphPath: 'hireDate',
    description: 'The hire date of the user',
    handledBy: 'optionA',
  },
  {
    name: 'employeeOrgData',
    type: 'object',
    writable: true,
    category: 'job',
    graphPath: 'employeeOrgData',
    betaOnly: true,
    description: 'Organizational data for employee (cost center, division)',
    handledBy: 'optionA',
  },

  // ===== IDENTITY =====
  {
    name: 'userPrincipalName',
    type: 'string',
    writable: true,
    required: true,
    category: 'identity',
    graphPath: 'userPrincipalName',
    description: 'The user principal name (UPN) of the user',
    handledBy: 'optionA',
  },
  {
    name: 'userType',
    type: 'string',
    writable: true,
    category: 'identity',
    graphPath: 'userType',
    description: 'String value that can be used to classify user types (Member, Guest)',
    handledBy: 'optionA',
  },
  {
    name: 'onPremisesImmutableId',
    type: 'string',
    writable: true,
    category: 'identity',
    graphPath: 'onPremisesImmutableId',
    description: 'Associates an on-premises Active Directory user account',
    handledBy: 'optionA',
  },

  // ===== PREFERENCES =====
  {
    name: 'usageLocation',
    type: 'string',
    writable: true,
    category: 'preferences',
    graphPath: 'usageLocation',
    description: 'Two-letter country code (ISO 3166) - required for license assignment',
    handledBy: 'optionA',
  },
  {
    name: 'preferredLanguage',
    type: 'string',
    writable: true,
    category: 'preferences',
    graphPath: 'preferredLanguage',
    description: 'Preferred language for the user (ISO 639-1 code)',
    handledBy: 'optionA',
  },
  {
    name: 'preferredDataLocation',
    type: 'string',
    writable: true,
    category: 'preferences',
    graphPath: 'preferredDataLocation',
    betaOnly: true,
    description: 'Preferred data location for multi-geo tenants',
    handledBy: 'optionA',
  },
  {
    name: 'mailboxSettings',
    type: 'object',
    writable: true,
    category: 'preferences',
    graphPath: 'mailboxSettings',
    description: 'Settings for the primary mailbox of the signed-in user',
    handledBy: 'optionA',
  },

  // ===== PERSONAL =====
  {
    name: 'birthday',
    type: 'date',
    writable: true,
    category: 'personal',
    graphPath: 'birthday',
    description: 'The birthday of the user',
    handledBy: 'optionB',
    peopleDataLabel: 'personAnniversaries',
    profileApiEndpoint: '/profile/anniversaries',
  },
  {
    name: 'interests',
    type: 'array',
    writable: true,
    category: 'personal',
    graphPath: 'interests',
    description: 'A list for the user to describe their interests',
    handledBy: 'optionB',
    peopleDataLabel: null,
    profileApiEndpoint: '/profile/interests',
  },
  {
    name: 'skills',
    type: 'array',
    writable: true,
    category: 'personal',
    graphPath: 'skills',
    description: 'A list for the user to enumerate their skills',
    handledBy: 'optionB',
    peopleDataLabel: 'personSkills',
    profileApiEndpoint: '/profile/skills',
  },
  {
    name: 'schools',
    type: 'array',
    writable: true,
    category: 'personal',
    graphPath: 'schools',
    description: 'A list of schools that the user has attended',
    handledBy: 'optionB',
    peopleDataLabel: null,
    profileApiEndpoint: '/profile/educationalActivities',
  },
  {
    name: 'projects',
    type: 'array',
    writable: true,
    category: 'personal',
    graphPath: 'pastProjects',
    description: 'A list for the user to enumerate their past projects',
    handledBy: 'optionB',
    peopleDataLabel: 'personProjects',
    profileApiEndpoint: '/profile/projects',
  },
  {
    name: 'responsibilities',
    type: 'array',
    writable: true,
    category: 'personal',
    graphPath: 'responsibilities',
    description: 'A list for the user to enumerate their responsibilities',
    handledBy: 'optionB',
    peopleDataLabel: null,
    profileApiEndpoint: '/profile/responsibilities',
  },
  {
    name: 'mySite',
    type: 'string',
    writable: true,
    category: 'personal',
    graphPath: 'mySite',
    description: 'The URL for the user personal site',
    handledBy: 'optionB',
    peopleDataLabel: 'personWebSite',
    profileApiEndpoint: '/profile/websites',
  },
  {
    name: 'certifications',
    type: 'array',
    writable: true,
    category: 'personal',
    graphPath: 'certifications',
    description: 'A list of certifications held by the user',
    handledBy: 'optionB',
    peopleDataLabel: 'personCertifications',
    profileApiEndpoint: '/profile/certifications',
  },
  {
    name: 'awards',
    type: 'array',
    writable: true,
    category: 'personal',
    graphPath: 'awards',
    description: 'A list of awards received by the user',
    handledBy: 'optionB',
    peopleDataLabel: 'personAwards',
    profileApiEndpoint: '/profile/awards',
  },
  {
    name: 'languages',
    type: 'array',
    writable: true,
    category: 'personal',
    graphPath: 'languages',
    betaOnly: false,
    description: 'Languages the user speaks with proficiency levels. Format: [{"language":"Norwegian","proficiency":"native"}]. Proficiency values: elementary, limitedWorking, professionalWorking, fullProfessional, nativeOrBilingual',
    handledBy: 'optionB',
    peopleDataLabel: null,
    profileApiEndpoint: '/profile/languages',
  },

  // ===== SECURITY =====
  {
    name: 'passwordPolicies',
    type: 'string',
    writable: true,
    category: 'security',
    graphPath: 'passwordPolicies',
    description: 'Specifies password policies for the user',
    handledBy: 'optionA',
  },
  {
    name: 'passwordProfile',
    type: 'object',
    writable: true,
    category: 'security',
    graphPath: 'passwordProfile',
    description: 'Specifies the password profile for the user',
    handledBy: 'optionA',
  },

  // ===== ON-PREMISES (Hybrid environments) =====
  {
    name: 'onPremisesExtensionAttributes',
    type: 'object',
    writable: true,
    category: 'onpremises',
    graphPath: 'onPremisesExtensionAttributes',
    description: 'Contains 15 custom extension attributes (extensionAttribute1-15)',
    handledBy: 'optionA',
  },

  // ===== LEGAL/COMPLIANCE =====
  {
    name: 'ageGroup',
    type: 'string',
    writable: true,
    category: 'legal',
    graphPath: 'ageGroup',
    description: 'Sets the age group of the user (null, Minor, NotAdult, Adult)',
    handledBy: 'optionA',
  },
  {
    name: 'consentProvidedForMinor',
    type: 'string',
    writable: true,
    category: 'legal',
    graphPath: 'consentProvidedForMinor',
    description: 'Sets whether consent has been obtained for minors',
    handledBy: 'optionA',
  },
];

/**
 * Map of property names to metadata for quick lookup
 */
export const PROPERTY_MAP = new Map<string, PropertyMetadata>(
  USER_PROPERTY_SCHEMA.map(prop => [prop.name, prop])
);

/**
 * Check if a column name is a standard Microsoft Graph property
 */
export function isStandardProperty(columnName: string): boolean {
  return PROPERTY_MAP.has(columnName);
}

/**
 * Get property metadata by column name
 */
export function getPropertyMetadata(columnName: string): PropertyMetadata | undefined {
  return PROPERTY_MAP.get(columnName);
}

/**
 * Internal CSV columns that are NOT enrichment properties
 * These are used by the tool for mapping/tracking but shouldn't go to Graph Connectors
 */
const INTERNAL_CSV_COLUMNS = new Set([
  'name',        // Maps to displayName (Option A)
  'email',       // Maps to userPrincipalName (Option A)
  'role',        // Internal tracking only
  'ManagerEmail', // Used for manager assignment (Option A)
]);

/**
 * Extract custom properties from CSV columns
 * (any column not in the standard schema AND not an internal column)
 * These custom properties will be handled by Option B (Graph Connectors)
 */
export function getCustomProperties(csvColumns: string[]): string[] {
  return csvColumns.filter(col =>
    !isStandardProperty(col) && !INTERNAL_CSV_COLUMNS.has(col)
  );
}

/**
 * Get properties handled by Option A (standard Entra ID properties)
 */
export function getOptionAProperties(): PropertyMetadata[] {
  return USER_PROPERTY_SCHEMA.filter(prop => prop.handledBy === 'optionA');
}

/**
 * Get properties handled by Option B (Graph Connector enrichment)
 */
export function getOptionBProperties(): PropertyMetadata[] {
  return USER_PROPERTY_SCHEMA.filter(prop => prop.handledBy === 'optionB');
}

/**
 * Get mapping of CSV property names to official people data labels
 * Properties without labels (null) become custom searchable properties
 */
export function getPeopleDataMapping(): Map<string, string | null> {
  const mapping = new Map<string, string | null>();
  USER_PROPERTY_SCHEMA
    .filter(p => p.handledBy === 'optionB')
    .forEach(p => mapping.set(p.name, p.peopleDataLabel ?? null));
  return mapping;
}

/**
 * Get all writable properties
 */
export function getWritableProperties(): PropertyMetadata[] {
  return USER_PROPERTY_SCHEMA.filter(prop => prop.writable);
}

/**
 * Get all required properties for user creation
 */
export function getRequiredProperties(): PropertyMetadata[] {
  return USER_PROPERTY_SCHEMA.filter(prop => prop.required);
}

/**
 * Get properties by category
 */
export function getPropertiesByCategory(category: PropertyCategory): PropertyMetadata[] {
  return USER_PROPERTY_SCHEMA.filter(prop => prop.category === category);
}

/**
 * Get all beta-only properties
 */
export function getBetaOnlyProperties(): PropertyMetadata[] {
  return USER_PROPERTY_SCHEMA.filter(prop => prop.betaOnly === true);
}

/**
 * Validate property value against schema
 */
export function validatePropertyValue(
  propertyName: string,
  value: any
): { valid: boolean; error?: string } {
  const metadata = getPropertyMetadata(propertyName);

  if (!metadata) {
    return { valid: true }; // Custom properties are always valid
  }

  // Type validation
  switch (metadata.type) {
    case 'string':
      if (typeof value !== 'string') {
        return { valid: false, error: `Expected string, got ${typeof value}` };
      }
      if (metadata.maxLength && value.length > metadata.maxLength) {
        return {
          valid: false,
          error: `Value exceeds maximum length of ${metadata.maxLength}`,
        };
      }
      break;

    case 'boolean':
      if (typeof value !== 'boolean') {
        return { valid: false, error: `Expected boolean, got ${typeof value}` };
      }
      break;

    case 'number':
      if (typeof value !== 'number') {
        return { valid: false, error: `Expected number, got ${typeof value}` };
      }
      break;

    case 'date':
      // Accept string or Date object
      if (!(value instanceof Date) && typeof value !== 'string') {
        return { valid: false, error: `Expected date string or Date object` };
      }
      // Validate date format if string
      if (typeof value === 'string' && isNaN(Date.parse(value))) {
        return { valid: false, error: `Invalid date format` };
      }
      break;

    case 'array':
      if (!Array.isArray(value)) {
        return { valid: false, error: `Expected array, got ${typeof value}` };
      }
      break;

    case 'object':
      if (typeof value !== 'object' || value === null || Array.isArray(value)) {
        return { valid: false, error: `Expected object, got ${typeof value}` };
      }
      break;
  }

  return { valid: true };
}

/**
 * Parse CSV value to appropriate type based on property metadata
 */
export function parsePropertyValue(propertyName: string, csvValue: string): any {
  const metadata = getPropertyMetadata(propertyName);

  if (!metadata) {
    // Custom property - return as-is
    return csvValue;
  }

  // Handle empty values
  if (csvValue === '' || csvValue === null || csvValue === undefined) {
    return undefined;
  }

  // Parse based on type
  switch (metadata.type) {
    case 'boolean':
      return csvValue.toLowerCase() === 'true' || csvValue === '1';

    case 'number':
      return Number(csvValue);

    case 'date':
      return csvValue; // Keep as string, Graph API expects ISO date strings

    case 'array':
      // Support JSON-encoded arrays like ['value1','value2'] or ["value1","value2"]
      // Also support comma-separated values as fallback
      try {
        // Try parsing as-is first (for double-quoted JSON)
        const parsed = JSON.parse(csvValue);
        if (Array.isArray(parsed)) {
          return parsed;
        }
      } catch {
        // Try converting single quotes to double quotes for JSON parsing
        try {
          const normalizedValue = csvValue.replace(/'/g, '"');
          const parsed = JSON.parse(normalizedValue);
          if (Array.isArray(parsed)) {
            return parsed;
          }
        } catch {
          // Not JSON, try comma-separated
        }
      }
      // Fallback to comma-separated values
      return csvValue.split(',').map(v => v.trim());

    case 'object':
      // Support JSON strings in CSV
      try {
        return JSON.parse(csvValue);
      } catch {
        return csvValue; // Return as string if not valid JSON
      }

    default:
      return csvValue;
  }
}

/**
 * Get properties that use the Profile API (have profileApiEndpoint defined)
 * These require delegated auth and go through /users/{id}/profile/* endpoints
 */
export function getProfileApiProperties(): PropertyMetadata[] {
  return USER_PROPERTY_SCHEMA.filter(prop => prop.profileApiEndpoint !== undefined);
}
