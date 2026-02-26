import { getOptionBProperties, getPeopleDataMapping } from '../schema/user-property-schema.js';

// All 13 official people data labels from Microsoft docs.
// Every property with an official label MUST use that label (Path A deserialization).
// Path A stores values as JsonElement — safely handles string and stringCollection.
// Reference: docs/MicrosoftDocs/build-connectors-with-people-data.md
const ENABLED_LABELS = new Set([
  // Group 1: Option B native (single CSV column → entity)
  'personSkills',         // stringCollection → skillProficiency
  'personNote',           // string           → personAnnotation
  'personCertifications', // stringCollection → personCertification
  'personProjects',       // stringCollection → projectParticipation
  'personAwards',         // stringCollection → personAward
  'personAnniversaries',  // stringCollection → personAnniversary
  'personWebSite',        // string           → webSite
  // Group 2: Composite (multiple Option A CSV columns → entity)
  'personName',            // string           → personName
  'personCurrentPosition', // string           → workPosition
  'personAddresses',       // stringCollection → itemAddress
  'personEmails',          // stringCollection → itemEmail
  'personPhones',          // stringCollection → itemPhone
  'personWebAccounts',     // stringCollection → webAccount (no CSV data yet)
]);

// Custom properties: searchable by Copilot/Search but not mapped to profile cards.
// CRITICAL: Must be type 'string' only (Path B deserialization).
const CUSTOM_PROPERTIES: Array<{ name: string; type: 'string' | 'stringCollection' }> = [
  { name: 'VTeam', type: 'string' },
];

const PROPERTY_NAME_REGEX = /^[A-Za-z0-9]+$/;
const MAX_PROPERTY_NAME_LENGTH = 32;
const LABEL_TYPE_OVERRIDES = new Map<string, 'string' | 'stringCollection'>([
  ['personAnniversaries', 'stringCollection'],
]);

function assertValidPropertyName(name: string): void {
  if (!PROPERTY_NAME_REGEX.test(name)) {
    throw new Error(
      `Invalid property name "${name}". People connector properties must be alphanumeric only.`
    );
  }
  if (name.length > MAX_PROPERTY_NAME_LENGTH) {
    throw new Error(
      `Invalid property name "${name}". People connector properties must be <= ${MAX_PROPERTY_NAME_LENGTH} characters.`
    );
  }
}

export class PeopleSchemaBuilder {
  /**
   * Build schema with people data labels + custom searchable properties.
   */
  static buildPeopleSchema(): any[] {
    const properties = [];
    const peopleDataMapping = getPeopleDataMapping();
    const optionBProps = getOptionBProperties();

    // Required: account mapping
    assertValidPropertyName('accountInformation');
    properties.push({
      name: 'accountInformation',
      type: 'string',
      labels: ['personAccount'],
    });

    // Labeled properties (Copilot-searchable via people data labels)
    for (const prop of optionBProps) {
      const label = peopleDataMapping.get(prop.name);
      if (!label || !ENABLED_LABELS.has(label)) {
        continue;
      }

      assertValidPropertyName(prop.name);
      const overrideType = LABEL_TYPE_OVERRIDES.get(label);
      const schemaType = overrideType ?? (prop.type === 'array' ? 'stringCollection' : 'string');

      properties.push({
        name: prop.name,
        type: schemaType,
        labels: [label],
      });
    }

    // Composite labeled properties (data from multiple Option A CSV columns)
    // Hardcoded like accountInformation since they don't come from getOptionBProperties()
    const compositeProperties: Array<{ name: string; type: string; labels: string[] }> = [
      { name: 'personNameInfo', type: 'string', labels: ['personName'] },
      { name: 'currentPosition', type: 'string', labels: ['personCurrentPosition'] },
      { name: 'addresses', type: 'stringCollection', labels: ['personAddresses'] },
      { name: 'emails', type: 'stringCollection', labels: ['personEmails'] },
      { name: 'phones', type: 'stringCollection', labels: ['personPhones'] },
      { name: 'webAccounts', type: 'stringCollection', labels: ['personWebAccounts'] },
    ];
    for (const comp of compositeProperties) {
      if (ENABLED_LABELS.has(comp.labels[0])) {
        assertValidPropertyName(comp.name);
        properties.push(comp);
      }
    }

    // Custom properties (searchable, no people data label)
    for (const custom of CUSTOM_PROPERTIES) {
      assertValidPropertyName(custom.name);
      properties.push({
        name: custom.name,
        type: custom.type,
        isSearchable: true,
        isQueryable: true,
        isRetrievable: true,
      });
    }

    return properties;
  }

  /**
   * Get the list of custom property names configured for the connector.
   */
  static getCustomPropertyNames(): string[] {
    return CUSTOM_PROPERTIES.map(p => p.name);
  }
}
