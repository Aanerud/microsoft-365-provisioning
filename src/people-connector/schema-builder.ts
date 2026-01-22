import { getOptionBProperties, getPeopleDataMapping, getCustomProperties } from '../schema/user-property-schema.js';

export class PeopleSchemaBuilder {
  /**
   * Build schema from Option B properties + custom properties from CSV
   * @param csvColumns - All columns from CSV to detect custom properties
   */
  static buildPeopleSchema(csvColumns: string[]): any[] {
    const properties = [];
    const peopleDataMapping = getPeopleDataMapping();
    const optionBProps = getOptionBProperties();
    const customProps = getCustomProperties(csvColumns);

    // REQUIRED: Account information property
    properties.push({
      name: 'accountInformation',
      type: 'string',
      labels: ['personAccount'],
      searchable: false,
      queryable: false,
      retrievable: true
    });

    // Add Option B standard properties (with or without labels)
    for (const prop of optionBProps) {
      const label = peopleDataMapping.get(prop.name);
      const schemaProperty: any = {
        name: prop.name,
        type: prop.type === 'array' ? 'stringCollection' : 'string',
        searchable: true,
        retrievable: true,
        refinable: prop.type === 'array' // Arrays can be refined
      };

      // Add label if available (official people data)
      if (label) {
        schemaProperty.labels = [label];
      }
      // Otherwise it's a custom searchable property (no label)

      properties.push(schemaProperty);
    }

    // Add custom organization properties (VTeam, BenefitPlan, etc.)
    for (const customProp of customProps) {
      properties.push({
        name: customProp,
        type: 'string', // Custom properties default to string
        searchable: true,
        retrievable: true,
        queryable: true
      });
    }

    return properties;
  }
}
