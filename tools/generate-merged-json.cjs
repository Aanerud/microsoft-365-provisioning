#!/usr/bin/env node
/**
 * Merges textcraft-europe.csv + profile-data.json → textcraft-europe.json
 */

const fs = require('fs');
const path = require('path');

const configDir = path.join(__dirname, '..', 'config');
const csvPath = path.join(configDir, 'textcraft-europe.csv');
const profilePath = path.join(configDir, 'profile-data.json');
const outputPath = path.join(configDir, 'textcraft-europe.json');

// ─── University location mapping ──────────────────────────────────────────────
const UNIVERSITY_LOCATIONS = {
  'Aarhus University': { city: 'Aarhus', countryOrRegion: 'Denmark' },
  'Accademia Nazionale d Arte Drammatica Rome': { city: 'Rome', countryOrRegion: 'Italy' },
  'BI Norwegian Business School': { city: 'Oslo', countryOrRegion: 'Norway' },
  'Bocconi University': { city: 'Milan', countryOrRegion: 'Italy' },
  'Charles University Prague': { city: 'Prague', countryOrRegion: 'Czech Republic' },
  'College of Europe Bruges': { city: 'Bruges', countryOrRegion: 'Belgium' },
  'Copenhagen Business School': { city: 'Copenhagen', countryOrRegion: 'Denmark' },
  'ESCP Business School': { city: 'Paris', countryOrRegion: 'France' },
  'ESSEC': { city: 'Cergy', countryOrRegion: 'France' },
  'ETH Zurich': { city: 'Zurich', countryOrRegion: 'Switzerland' },
  'Heidelberg University': { city: 'Heidelberg', countryOrRegion: 'Germany' },
  'Humboldt University Berlin': { city: 'Berlin', countryOrRegion: 'Germany' },
  'ISAE-SUPAERO Toulouse': { city: 'Toulouse', countryOrRegion: 'France' },
  'Istituto Marangoni Milan': { city: 'Milan', countryOrRegion: 'Italy' },
  'KTH Stockholm': { city: 'Stockholm', countryOrRegion: 'Sweden' },
  'Karolinska Institutet': { city: 'Stockholm', countryOrRegion: 'Sweden' },
  'LSE': { city: 'London', countryOrRegion: 'United Kingdom' },
  'London College of Communication': { city: 'London', countryOrRegion: 'United Kingdom' },
  'Ludwig Maximilian University Munich': { city: 'Munich', countryOrRegion: 'Germany' },
  'Lund University': { city: 'Lund', countryOrRegion: 'Sweden' },
  'Nielsen Norman Group': { city: 'Fremont', countryOrRegion: 'United States' },
  'Politecnico di Milano': { city: 'Milan', countryOrRegion: 'Italy' },
  'Sapienza University Rome': { city: 'Rome', countryOrRegion: 'Italy' },
  'Sciences Po Paris': { city: 'Paris', countryOrRegion: 'France' },
  'Sorbonne University': { city: 'Paris', countryOrRegion: 'France' },
  'Stockholm University': { city: 'Stockholm', countryOrRegion: 'Sweden' },
  'TU Berlin': { city: 'Berlin', countryOrRegion: 'Germany' },
  'Toulouse': { city: 'Toulouse', countryOrRegion: 'France' },
  'Trinity College Dublin': { city: 'Dublin', countryOrRegion: 'Ireland' },
  'UdK Berlin': { city: 'Berlin', countryOrRegion: 'Germany' },
  'Universidad Carlos III de Madrid': { city: 'Madrid', countryOrRegion: 'Spain' },
  'Universidad Complutense Madrid': { city: 'Madrid', countryOrRegion: 'Spain' },
  'Universidad Politécnica de Madrid': { city: 'Madrid', countryOrRegion: 'Spain' },
  'Universidad de Salamanca': { city: 'Salamanca', countryOrRegion: 'Spain' },
  'Universite Libre de Bruxelles': { city: 'Brussels', countryOrRegion: 'Belgium' },
  'University College Dublin': { city: 'Dublin', countryOrRegion: 'Ireland' },
  'University of Amsterdam': { city: 'Amsterdam', countryOrRegion: 'Netherlands' },
  'University of Bologna': { city: 'Bologna', countryOrRegion: 'Italy' },
  'University of Bristol': { city: 'Bristol', countryOrRegion: 'United Kingdom' },
  'University of Cambridge': { city: 'Cambridge', countryOrRegion: 'United Kingdom' },
  'University of Copenhagen': { city: 'Copenhagen', countryOrRegion: 'Denmark' },
  'University of East Anglia': { city: 'Norwich', countryOrRegion: 'United Kingdom' },
  'University of Edinburgh': { city: 'Edinburgh', countryOrRegion: 'United Kingdom' },
  'University of Lisbon': { city: 'Lisbon', countryOrRegion: 'Portugal' },
  'University of Manchester': { city: 'Manchester', countryOrRegion: 'United Kingdom' },
  'University of Mannheim': { city: 'Mannheim', countryOrRegion: 'Germany' },
  'University of Oxford': { city: 'Oxford', countryOrRegion: 'United Kingdom' },
  'University of Stuttgart': { city: 'Stuttgart', countryOrRegion: 'Germany' },
  'University of Vienna': { city: 'Vienna', countryOrRegion: 'Austria' },
  'University of Warsaw': { city: 'Warsaw', countryOrRegion: 'Poland' },
  'University of Zurich': { city: 'Zurich', countryOrRegion: 'Switzerland' },
  'Uppsala University': { city: 'Uppsala', countryOrRegion: 'Sweden' },
  'WHU Otto Beisheim School': { city: 'Vallendar', countryOrRegion: 'Germany' },
  'Warsaw University of Technology': { city: 'Warsaw', countryOrRegion: 'Poland' },
  // These appear as institution names but are cert bodies / acronyms — no campus location
  'CP Certification': null,
  'IAF': null,
  'IFCN': null,
  'MA Public Affairs': null,
  'OASIS': null,
  'Project Management Institute': null,
};

// ─── CSV parser (handles quoted fields with commas) ───────────────────────────
function parseCSV(text) {
  const lines = text.split('\n').filter(l => l.trim());
  const headers = parseCSVLine(lines[0]);
  return lines.slice(1).map(line => {
    const values = parseCSVLine(line);
    const obj = {};
    headers.forEach((h, i) => { obj[h] = values[i] || ''; });
    return obj;
  });
}

function parseCSVLine(line) {
  const fields = [];
  let current = '';
  let inQuotes = false;
  for (let i = 0; i < line.length; i++) {
    const ch = line[i];
    if (ch === '"' && !inQuotes) {
      inQuotes = true;
    } else if (ch === '"' && inQuotes) {
      if (i + 1 < line.length && line[i + 1] === '"') {
        current += '"';
        i++;
      } else {
        inQuotes = false;
      }
    } else if (ch === ',' && !inQuotes) {
      fields.push(current);
      current = '';
    } else {
      current += ch;
    }
  }
  fields.push(current);
  return fields;
}

// ─── Parse Python-style string arrays from CSV ───────────────────────────────
function parsePythonArray(str) {
  if (!str || str === '[]') return [];
  // e.g. ['val1','val2'] or ['+39 02 8765 4321']
  const inner = str.replace(/^\[/, '').replace(/\]$/, '');
  if (!inner.trim()) return [];
  const items = [];
  let current = '';
  let inQuote = false;
  for (let i = 0; i < inner.length; i++) {
    const ch = inner[i];
    if (ch === "'" && !inQuote) { inQuote = true; }
    else if (ch === "'" && inQuote) { inQuote = false; }
    else if (ch === ',' && !inQuote) {
      items.push(current.trim());
      current = '';
    } else {
      current += ch;
    }
  }
  if (current.trim()) items.push(current.trim());
  return items;
}

// ─── Strip @odata.type from any object ────────────────────────────────────────
function stripOdata(obj) {
  if (!obj || typeof obj !== 'object') return obj;
  if (Array.isArray(obj)) return obj.map(stripOdata);
  const result = {};
  for (const [k, v] of Object.entries(obj)) {
    if (k === '@odata.type') continue;
    result[k] = stripOdata(v);
  }
  return result;
}

// ─── Main ─────────────────────────────────────────────────────────────────────
const csvText = fs.readFileSync(csvPath, 'utf8');
const csvRows = parseCSV(csvText);
const profileData = JSON.parse(fs.readFileSync(profilePath, 'utf8'));

// Index profiles by UPN
const profileMap = {};
profileData.forEach(p => { profileMap[p.userPrincipalName] = p.profile; });

const ENTRA_FIELDS = [
  'name', 'givenName', 'surname', 'jobTitle', 'department', 'employeeType',
  'companyName', 'officeLocation', 'streetAddress', 'city', 'state', 'country',
  'postalCode', 'usageLocation', 'preferredLanguage', 'mobilePhone',
  'employeeId', 'employeeHireDate', 'ManagerEmail', 'role'
];

const CUSTOM_PROPS = [
  'VTeam', 'BenefitPlan', 'CostCenter', 'BuildingAccess', 'ProjectCode',
  'WritingStyle', 'Specialization'
];

const output = csvRows.map(row => {
  const record = {};

  // email
  record.email = row.email;

  // Entra fields
  for (const f of ENTRA_FIELDS) {
    record[f] = row[f] || '';
  }

  // businessPhones — parse from Python-style array
  record.businessPhones = parsePythonArray(row.businessPhones);

  // Custom properties
  for (const f of CUSTOM_PROPS) {
    record[f] = row[f] || '';
  }

  // Profile data
  const profile = profileMap[row.email];
  if (!profile) {
    console.warn(`WARNING: No profile data for ${row.email}`);
    return record;
  }

  // skills
  record.skills = (profile.skills || []).map(s => {
    const obj = { displayName: s.displayName };
    if (s.categories) obj.categories = s.categories;
    if (s.collaborationTags) obj.collaborationTags = s.collaborationTags;
    if (s.proficiency) obj.proficiency = s.proficiency;
    return obj;
  });

  // interests
  record.interests = (profile.interests || []).map(i => {
    const obj = { displayName: i.displayName };
    if (i.categories) obj.categories = i.categories;
    if (i.collaborationTags) obj.collaborationTags = i.collaborationTags;
    return obj;
  });

  // languages
  record.languages = (profile.languages || []).map(l => ({
    displayName: l.displayName,
    tag: l.tag,
    reading: l.reading,
    spoken: l.spoken,
    written: l.written,
  }));

  // aboutMe
  if (profile.notes && profile.notes.length > 0 && profile.notes[0].detail) {
    record.aboutMe = profile.notes[0].detail.content;
  }

  // certifications
  record.certifications = (profile.certifications || []).map(c => {
    const obj = { displayName: c.displayName };
    if (c.issuingCompany) obj.issuingCompany = c.issuingCompany;
    if (c.issuedDate) obj.issuedDate = c.issuedDate;
    return obj;
  });

  // awards
  record.awards = (profile.awards || []).map(a => ({ displayName: a.displayName }));

  // projects
  record.projects = (profile.projects || []).map(p => {
    const obj = { displayName: p.displayName };
    if (p.categories) obj.categories = p.categories;
    if (p.collaborationTags) obj.collaborationTags = p.collaborationTags;
    return obj;
  });

  // educationalActivities
  record.educationalActivities = (profile.educationalActivities || []).map(e => {
    const obj = {};

    // institution
    if (e.institution) {
      obj.institution = { displayName: e.institution.displayName };
      const loc = UNIVERSITY_LOCATIONS[e.institution.displayName];
      if (loc) {
        obj.institution.location = loc;
      }
    }

    // program
    if (e.program) {
      obj.program = { displayName: e.program.displayName };
      if (e.program.abbreviation) obj.program.abbreviation = e.program.abbreviation;
      if (e.program.fieldsOfStudy) {
        // Convert string to array
        obj.program.fieldsOfStudy = [e.program.fieldsOfStudy];
      }
    }

    if (e.startMonthYear) obj.startMonthYear = e.startMonthYear;
    if (e.endMonthYear) obj.endMonthYear = e.endMonthYear;
    if (e.completionMonthYear) obj.completionMonthYear = e.completionMonthYear;

    return obj;
  });

  // publications
  record.publications = (profile.publications || []).map(p => {
    const obj = { displayName: p.displayName };
    if (p.publisher) obj.publisher = p.publisher;
    if (p.publishedDate) obj.publishedDate = p.publishedDate;
    if (p.description) obj.description = p.description;
    if (p.webUrl) obj.webUrl = p.webUrl;
    return obj;
  });

  // patents
  record.patents = stripOdata(profile.patents || []);

  // responsibilities
  record.responsibilities = (profile.responsibilities || []).map(r => {
    const obj = { displayName: r.displayName };
    if (r.description) obj.description = r.description;
    return obj;
  });

  // mySite (from CSV if exists)
  if (row.mySite) record.mySite = row.mySite;

  // birthday (from CSV if exists)
  if (row.birthday) record.birthday = row.birthday;

  return record;
});

fs.writeFileSync(outputPath, JSON.stringify(output, null, 2), 'utf8');

// ─── Summary ──────────────────────────────────────────────────────────────────
console.log(`\nWrote ${output.length} records to ${outputPath}\n`);

const fieldCounts = {};
output.forEach(r => {
  for (const [k, v] of Object.entries(r)) {
    if (!fieldCounts[k]) fieldCounts[k] = 0;
    if (Array.isArray(v)) {
      if (v.length > 0) fieldCounts[k]++;
    } else if (v !== '' && v !== undefined && v !== null) {
      fieldCounts[k]++;
    }
  }
});

console.log('Field population summary:');
console.log('─'.repeat(50));
for (const [field, count] of Object.entries(fieldCounts)) {
  const pct = ((count / output.length) * 100).toFixed(0);
  console.log(`  ${field.padEnd(28)} ${String(count).padStart(3)} / ${output.length}  (${pct}%)`);
}
