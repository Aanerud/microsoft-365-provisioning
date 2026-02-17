import { getOptionBProperties, getPeopleDataMapping } from '../schema/user-property-schema.js';

// People data labels to include in the connector schema.
// Each label maps to a Microsoft profile entity that Copilot can search.
// NOTE: Adding personNote here broke profile enrichment for personSkills too.
// Only personSkills is proven to work. Add labels back one at a time after verifying.
const ENABLED_LABELS = new Set(['personSkills']);

// Custom properties: searchable by Copilot/Search but not mapped to profile cards.
// NOTE: Custom properties without people data labels break profile enrichment
// for the entire connection. Keep this empty for people connectors.
// Use a separate non-people connector for custom searchable properties.
const CUSTOM_PROPERTIES: Array<{ name: string; type: 'string' | 'stringCollection' }> = [];

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
