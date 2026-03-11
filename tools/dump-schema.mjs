import { PeopleSchemaBuilder } from '../dist/people-connector/schema-builder.js';
const schema = PeopleSchemaBuilder.buildPeopleSchema();
console.log(JSON.stringify(schema, null, 2));
console.log(`\nTotal properties: ${schema.length}`);
const names = schema.map(p => p.name);
const dupes = names.filter((n, i) => names.indexOf(n) !== i);
if (dupes.length) console.log(`\nDUPLICATES: ${dupes.join(', ')}`);
