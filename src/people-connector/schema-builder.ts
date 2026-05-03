import { getOptionBProperties, getPeopleDataMapping, getCustomProperties } from '../schema/user-property-schema.js';

// All 18 people data labels from Microsoft docs.
// Every property with an official label MUST use that label (Path A deserialization).
// Path A stores values as JsonElement — safely handles string and stringCollection.
// MS docs: isQueryable, isRefinable, isRetrievable, isSearchable are ALL IGNORED
// for people connectors. All data about a person is indexed by default.
export const ENABLED_LABELS = new Set([
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
  // Group 3: Additional people data labels
  'personEducationalActivities', // stringCollection → educationalActivity
  'personInterests',             // stringCollection → personInterest
  'personLanguages',             // stringCollection → languageProficiency
  'personPublications',          // stringCollection → itemPublication
  'personPatents',               // stringCollection → itemPatent
]);

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
   * descriptions is an optional map of custom property name → description string.
   * When present, descriptions are included in the schema per Microsoft docs
   * (helps Copilot reason about custom fields).
   */
  static buildPeopleSchema(csvColumns: string[], descriptions?: Record<string, string>): any[] {
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

    // Custom properties (searchable, no people data label — string only, Path B)
    // Dynamically detected from config columns. We are a slave to the config —
    // whatever top-level columns aren't in the standard schema become customs.
    const customNames = getCustomProperties(csvColumns);

    for (const name of customNames) {
      assertValidPropertyName(name);
      const entry: any = {
        name,
        type: 'string',  // MUST be string only (Path B deserialization)
        isSearchable: true,
        isQueryable: true,
        isRetrievable: true,
      };
      if (descriptions?.[name]) {
        entry.description = descriptions[name];
      }
      properties.push(entry);
    }

    return properties;
  }

  /**
   * Get the list of custom property names configured for the connector.
   * Dynamically detected from config columns — no hardcoded fallback.
   */
  static getCustomPropertyNames(csvColumns: string[]): string[] {
    return getCustomProperties(csvColumns);
  }
}
