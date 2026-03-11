import fs from 'fs';
import { parse } from 'csv-parse/sync';
const data = parse(fs.readFileSync('config/textcraft-europe.csv', 'utf8'), { columns: true });
let withEdu = 0, withPub = 0, withPat = 0;
data.forEach(r => {
  if (r.educationalActivities && r.educationalActivities !== '[]') withEdu++;
  if (r.publications && r.publications !== '[]') withPub++;
  if (r.patents && r.patents !== '[]') withPat++;
});
console.log(`Rows with educationalActivities: ${withEdu}/${data.length}`);
console.log(`Rows with publications: ${withPub}/${data.length}`);
console.log(`Rows with patents: ${withPat}/${data.length}`);
