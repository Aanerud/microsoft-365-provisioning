import { getOptionBProperties, getPeopleDataMapping } from '../schema/user-property-schema.js';

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
   * Build schema from Option B properties with people data labels only.
   */
  static buildPeopleSchema(): any[] {
    const properties = [];
    const peopleDataMapping = getPeopleDataMapping();
    const optionBProps = getOptionBProperties();

    assertValidPropertyName('accountInformation');
    properties.push({
      name: 'accountInformation',
      type: 'string',
      labels: ['personAccount'],
    });

    for (const prop of optionBProps) {
      const label = peopleDataMapping.get(prop.name);
      if (!label) {
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

    return properties;
  }
}
