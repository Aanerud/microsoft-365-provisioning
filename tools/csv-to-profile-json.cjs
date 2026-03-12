/**
 * Converts textcraft-europe.csv to MS Graph Profile API compliant JSON.
 * Maps each CSV column to its corresponding Graph beta resource type.
 * Enriches all entities with full property sets: dates, descriptions,
 * collaboration tags, issuing authorities, etc.
 */
const fs = require('fs');
const path = require('path');

const csvPath = path.join(__dirname, '..', 'config', 'textcraft-europe.csv');
const outPath = path.join(__dirname, '..', 'config', 'profile-data.json');

// ============================================================
// Deterministic seeded random (based on string hash)
// ============================================================

function hashStr(str) {
  let h = 0;
  for (let i = 0; i < str.length; i++) {
    h = ((h << 5) - h + str.charCodeAt(i)) | 0;
  }
  return Math.abs(h);
}

// Returns a deterministic pseudo-random [0,1) for a given seed string
function seededRand(seed) {
  const h = hashStr(seed);
  return (h % 10000) / 10000;
}

function pick(arr, seed) {
  return arr[hashStr(seed) % arr.length];
}

function pickN(arr, n, seed) {
  const shuffled = [...arr];
  const h = hashStr(seed);
  for (let i = shuffled.length - 1; i > 0; i--) {
    const j = (h + i * 31) % (i + 1);
    [shuffled[i], shuffled[j]] = [shuffled[j], shuffled[i]];
  }
  return shuffled.slice(0, n);
}

// ============================================================
// CSV Parsing
// ============================================================

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
  const result = [];
  let current = '';
  let inQuotes = false;
  let bracketDepth = 0;
  for (let i = 0; i < line.length; i++) {
    const ch = line[i];
    if (ch === '"' && bracketDepth === 0) {
      inQuotes = !inQuotes;
    } else if (ch === '[') {
      bracketDepth++;
      current += ch;
    } else if (ch === ']') {
      bracketDepth--;
      current += ch;
    } else if (ch === ',' && !inQuotes && bracketDepth === 0) {
      result.push(current.trim());
      current = '';
    } else {
      current += ch;
    }
  }
  result.push(current.trim());
  return result;
}

function parseList(val) {
  if (!val || val === '[]') return [];
  let s = val.trim();
  if (s.startsWith('[') && s.endsWith(']')) s = s.slice(1, -1);
  if (!s.trim()) return [];
  const items = [];
  let current = '';
  let inQuote = false;
  let quoteChar = '';
  for (let i = 0; i < s.length; i++) {
    const ch = s[i];
    if (!inQuote && (ch === "'" || ch === '"')) {
      inQuote = true;
      quoteChar = ch;
    } else if (inQuote && ch === quoteChar) {
      inQuote = false;
    } else if (!inQuote && ch === ',') {
      const trimmed = current.trim();
      if (trimmed) items.push(trimmed);
      current = '';
    } else {
      current += ch;
    }
  }
  const trimmed = current.trim();
  if (trimmed) items.push(trimmed);
  return items;
}

// ============================================================
// Language mapping
// ============================================================

const langProficiencyMap = {
  'native': 'nativeOrBilingual',
  'bilingual': 'nativeOrBilingual',
  'fluent': 'fullProfessional',
  'professional': 'professionalWorking',
  'conversational': 'conversational',
  'limited working': 'limitedWorking',
  'elementary': 'elementary',
  'basic': 'elementary',
};

const langTagMap = {
  'italian': 'it-IT', 'english': 'en-US', 'french': 'fr-FR',
  'german': 'de-DE', 'spanish': 'es-ES', 'swedish': 'sv-SE',
  'norwegian': 'nb-NO', 'danish': 'da-DK', 'dutch': 'nl-NL',
  'portuguese': 'pt-PT', 'finnish': 'fi-FI', 'polish': 'pl-PL',
  'czech': 'cs-CZ', 'hungarian': 'hu-HU', 'romanian': 'ro-RO',
  'greek': 'el-GR', 'turkish': 'tr-TR', 'russian': 'ru-RU',
  'arabic': 'ar-SA', 'japanese': 'ja-JP', 'chinese': 'zh-CN',
  'mandarin': 'zh-CN', 'korean': 'ko-KR', 'hindi': 'hi-IN',
  'catalan': 'ca-ES', 'basque': 'eu-ES', 'galician': 'gl-ES',
  'welsh': 'cy-GB', 'irish': 'ga-IE', 'scottish gaelic': 'gd-GB',
  'icelandic': 'is-IS', 'estonian': 'et-EE', 'latvian': 'lv-LV',
  'lithuanian': 'lt-LT', 'slovenian': 'sl-SI', 'croatian': 'hr-HR',
  'serbian': 'sr-RS', 'bulgarian': 'bg-BG', 'slovak': 'sk-SK',
  'ukrainian': 'uk-UA', 'maltese': 'mt-MT', 'albanian': 'sq-AL',
  'bosnian': 'bs-BA', 'macedonian': 'mk-MK', 'luxembourgish': 'lb-LU',
  'afrikaans': 'af-ZA', 'swahili': 'sw-KE', 'hebrew': 'he-IL',
  'persian': 'fa-IR', 'thai': 'th-TH', 'vietnamese': 'vi-VN',
  'indonesian': 'id-ID', 'malay': 'ms-MY', 'tagalog': 'tl-PH',
  'bengali': 'bn-IN', 'tamil': 'ta-IN', 'urdu': 'ur-PK',
  'sign language': 'sgn', 'british sign language': 'bfi',
};

function parseLanguage(langStr) {
  const match = langStr.match(/^(.+?)\s*\((.+?)\)\s*$/);
  let name, level;
  if (match) {
    name = match[1].trim();
    level = match[2].trim().toLowerCase();
  } else {
    name = langStr.trim();
    level = null;
  }
  const proficiency = level ? (langProficiencyMap[level] || 'professionalWorking') : 'professionalWorking';
  const tag = langTagMap[name.toLowerCase()] || null;
  const entry = {
    "@odata.type": "#microsoft.graph.languageProficiency",
    "displayName": name,
    "tag": tag || undefined,
    "spoken": proficiency,
    "written": proficiency,
    "reading": proficiency,
  };
  // Remove undefined tag
  if (!entry.tag) delete entry.tag;
  return entry;
}

// ============================================================
// Skill enrichment
// ============================================================

// Only safe proficiency values (see CLAUDE.md — generalProfessional & advancedProfessional fail)
const safeProficiencyLevels = ['elementary', 'limitedWorking', 'expert'];
const collabTags = ['askMeAbout', 'ableToMentor', 'wantsToLearn', 'wantsToImprove'];

function buildSkill(skillName, userSeed, idx) {
  const seed = `${userSeed}:skill:${idx}`;
  const entry = {
    "@odata.type": "#microsoft.graph.skillProficiency",
    "displayName": skillName,
    "categories": ["professional"],
    "proficiency": pick(safeProficiencyLevels, seed + ':prof'),
    "collaborationTags": pickN(collabTags, 1 + (hashStr(seed + ':ct') % 2), seed + ':ct'),
  };
  return entry;
}

// ============================================================
// Interest enrichment
// ============================================================

const interestCategories = ['professional', 'personal', 'hobby', 'volunteer'];

function buildInterest(interestName, userSeed, idx) {
  const seed = `${userSeed}:interest:${idx}`;
  const cat = pick(interestCategories, seed);
  const entry = {
    "@odata.type": "#microsoft.graph.personInterest",
    "displayName": interestName,
    "categories": [cat],
  };
  // ~40% get a collaborationTag
  if (seededRand(seed + ':ct') < 0.4) {
    entry.collaborationTags = [pick(collabTags, seed + ':ctv')];
  }
  return entry;
}

// ============================================================
// Project enrichment
// ============================================================

const projectCategories = [
  'technology', 'content', 'operations', 'quality', 'strategy',
  'communications', 'research', 'design', 'marketing', 'compliance',
];

function buildProject(projectName, userSeed, idx, hireDate) {
  const seed = `${userSeed}:project:${idx}`;
  const entry = {
    "@odata.type": "#microsoft.graph.projectParticipation",
    "displayName": projectName,
    "categories": [pick(projectCategories, seed)],
  };
  // ~50% get collaborationTags
  if (seededRand(seed + ':ct') < 0.5) {
    entry.collaborationTags = [pick(collabTags, seed + ':ctv')];
  }
  // Add client company for ~30%
  if (seededRand(seed + ':client') < 0.3) {
    const clients = ['TextCraft Europe', 'Internal', 'Pan-European Initiative', 'Cross-Office Collaboration'];
    entry.client = {
      "displayName": pick(clients, seed + ':cname'),
    };
  }
  return entry;
}

// ============================================================
// Responsibility enrichment
// ============================================================

function buildResponsibility(respName, userSeed, idx) {
  const seed = `${userSeed}:resp:${idx}`;
  const entry = {
    "@odata.type": "#microsoft.graph.personResponsibility",
    "displayName": respName,
    "description": respName, // displayName is the short form, description is the full text
  };
  // ~35% get collaborationTags
  if (seededRand(seed + ':ct') < 0.35) {
    entry.collaborationTags = [pick(collabTags, seed + ':ctv')];
  }
  return entry;
}

// ============================================================
// Certification enrichment — varied dates
// ============================================================

// Map well-known cert names to issuing authorities/companies
const certIssuerMap = {
  'MBA': { issuingAuthority: null, issuingCompany: null }, // institution-granted
  'MSc': { issuingAuthority: null, issuingCompany: null },
  'BSc': { issuingAuthority: null, issuingCompany: null },
  'BA':  { issuingAuthority: null, issuingCompany: null },
  'MA':  { issuingAuthority: null, issuingCompany: null },
  'PhD': { issuingAuthority: null, issuingCompany: null },
  'PMP': { issuingAuthority: 'Project Management Institute', issuingCompany: 'PMI' },
  'CFA': { issuingAuthority: 'CFA Institute', issuingCompany: 'CFA Institute' },
  'Certified Scrum Master': { issuingAuthority: 'Scrum Alliance', issuingCompany: 'Scrum Alliance' },
  'Six Sigma Black Belt': { issuingAuthority: 'American Society for Quality', issuingCompany: 'ASQ' },
  'Six Sigma Green Belt': { issuingAuthority: 'American Society for Quality', issuingCompany: 'ASQ' },
  'Google Analytics': { issuingAuthority: 'Google', issuingCompany: 'Google' },
  'AWS': { issuingAuthority: 'Amazon Web Services', issuingCompany: 'AWS' },
  'Azure': { issuingAuthority: 'Microsoft', issuingCompany: 'Microsoft' },
  'PRINCE2': { issuingAuthority: 'Axelos', issuingCompany: 'Axelos' },
  'ITIL': { issuingAuthority: 'Axelos', issuingCompany: 'Axelos' },
  'CISSP': { issuingAuthority: 'ISC2', issuingCompany: 'ISC2' },
  'TOGAF': { issuingAuthority: 'The Open Group', issuingCompany: 'The Open Group' },
};

function findCertIssuer(certName) {
  const lower = certName.toLowerCase();
  for (const [key, val] of Object.entries(certIssuerMap)) {
    if (lower.includes(key.toLowerCase())) return val;
  }
  return null;
}

// Extract institution from "Degree - Institution" format
function extractCertInstitution(certName) {
  const dashMatch = certName.match(/^.+?\s*-\s*(.+)$/);
  if (dashMatch) return dashMatch[1].trim();
  return null;
}

function buildCertification(certName, userSeed, idx, hireYear) {
  const seed = `${userSeed}:cert:${idx}`;
  const r = seededRand(seed + ':dateVariant');

  const entry = {
    "@odata.type": "#microsoft.graph.personCertification",
    "displayName": certName,
  };

  // Try to find issuer
  const issuer = findCertIssuer(certName);
  const institution = extractCertInstitution(certName);

  if (issuer && issuer.issuingAuthority) {
    entry.issuingAuthority = issuer.issuingAuthority;
    entry.issuingCompany = issuer.issuingCompany;
  } else if (institution) {
    entry.issuingCompany = institution;
  }

  // Date patterns — varied:
  // Pattern 1 (~25%): No dates at all (just displayName + issuer)
  // Pattern 2 (~25%): Only issuedDate
  // Pattern 3 (~25%): issuedDate + startDate (training period implied)
  // Pattern 4 (~25%): Full range: startDate, issuedDate, endDate (expiring cert)

  const baseYear = hireYear - 3 + (hashStr(seed + ':yr') % 8); // somewhere around hire date ± years
  const month = 1 + (hashStr(seed + ':mo') % 12);
  const day = 1 + (hashStr(seed + ':dy') % 28);
  const dateStr = (y, m, d) => `${y}-${String(m).padStart(2, '0')}-${String(d).padStart(2, '0')}`;

  if (r < 0.25) {
    // Pattern 1: no dates
  } else if (r < 0.50) {
    // Pattern 2: issuedDate only
    entry.issuedDate = dateStr(baseYear, month, day);
  } else if (r < 0.75) {
    // Pattern 3: startDate + issuedDate (completed after study period)
    const startMonth = 1 + (hashStr(seed + ':sm') % 12);
    entry.startDate = dateStr(baseYear - 1, startMonth, 1);
    entry.issuedDate = dateStr(baseYear, month, day);
  } else {
    // Pattern 4: full range with expiry (renewable certs like PMP, Scrum, CFA)
    const startMonth = 1 + (hashStr(seed + ':sm') % 12);
    entry.startDate = dateStr(baseYear - 1, startMonth, 1);
    entry.issuedDate = dateStr(baseYear, month, day);
    entry.endDate = dateStr(baseYear + 3, month, day);
  }

  return entry;
}

// ============================================================
// Award enrichment — varied dates and issuing authorities
// ============================================================

const awardIssuers = [
  'European Content Awards',
  'Industry Excellence Board',
  'International Writing Association',
  'European Publishing Council',
  'Digital Communications Forum',
  'Creative Industries Alliance',
  'Nordic Business Council',
  'Mediterranean Innovation Foundation',
  'Central European Arts Foundation',
  'Pan-European Quality Institute',
];

function buildAward(awardStr, userSeed, idx) {
  const seed = `${userSeed}:award:${idx}`;
  const r = seededRand(seed + ':dateVariant');

  const entry = {
    "@odata.type": "#microsoft.graph.personAward",
    "displayName": awardStr.trim(),
  };

  // Extract year from name if present
  const yearMatch = awardStr.match(/(\d{4})/);
  const year = yearMatch ? parseInt(yearMatch[1]) : null;

  // Date patterns — varied:
  // Pattern 1 (~20%): No date, no issuer — bare award
  // Pattern 2 (~30%): issuedDate only (extracted from name or generated)
  // Pattern 3 (~30%): issuedDate + issuingAuthority
  // Pattern 4 (~20%): issuedDate + issuingAuthority + description

  if (r < 0.20) {
    // Pattern 1: bare — no dates
  } else if (r < 0.50) {
    // Pattern 2: date only
    if (year) {
      const month = 1 + (hashStr(seed + ':mo') % 12);
      const day = 1 + (hashStr(seed + ':dy') % 28);
      entry.issuedDate = `${year}-${String(month).padStart(2, '0')}-${String(day).padStart(2, '0')}`;
    }
  } else if (r < 0.80) {
    // Pattern 3: date + authority
    if (year) {
      const month = 1 + (hashStr(seed + ':mo') % 12);
      const day = 1 + (hashStr(seed + ':dy') % 28);
      entry.issuedDate = `${year}-${String(month).padStart(2, '0')}-${String(day).padStart(2, '0')}`;
    }
    entry.issuingAuthority = pick(awardIssuers, seed + ':iss');
  } else {
    // Pattern 4: date + authority + description
    if (year) {
      const month = 1 + (hashStr(seed + ':mo') % 12);
      const day = 1 + (hashStr(seed + ':dy') % 28);
      entry.issuedDate = `${year}-${String(month).padStart(2, '0')}-${String(day).padStart(2, '0')}`;
    }
    entry.issuingAuthority = pick(awardIssuers, seed + ':iss');
    entry.description = `Awarded for outstanding contribution and excellence in the field.`;
  }

  return entry;
}

// ============================================================
// Education enrichment — dates, fields of study
// ============================================================

const fieldsOfStudyMap = {
  'MBA': 'Business Administration',
  'MSc': 'Science',
  'BSc': 'Science',
  'BA': 'Arts',
  'MA': 'Arts',
  'PhD': 'Doctoral Studies',
  'LLM': 'Law',
  'BEng': 'Engineering',
  'MEng': 'Engineering',
};

function buildEducation(eduStr, userSeed, idx, hireYear) {
  const seed = `${userSeed}:edu:${idx}`;
  const dashMatch = eduStr.match(/^(.+?)\s*-\s*(.+)$/);

  let degree, institution;
  if (dashMatch) {
    degree = dashMatch[1].trim();
    institution = dashMatch[2].trim();
  } else {
    degree = eduStr.trim();
    institution = null;
  }

  const abbrMatch = degree.match(/^(MBA|MSc|BSc|BA|MA|PhD|PMP|CFA|LLM|DBA|MPhil|BEng|MEng)\b/i);
  const abbreviation = abbrMatch ? abbrMatch[1].toUpperCase() : null;

  const entry = {
    "@odata.type": "#microsoft.graph.educationalActivity",
  };

  if (institution) {
    entry.institution = { "displayName": institution };
  }

  const program = {};
  if (abbreviation) {
    program.abbreviation = abbreviation;
    if (fieldsOfStudyMap[abbreviation]) {
      // Try to extract field from degree name after the abbreviation
      const afterAbbr = degree.replace(/^(MBA|MSc|BSc|BA|MA|PhD|LLM|BEng|MEng)\s*/i, '').trim();
      program.fieldsOfStudy = afterAbbr || fieldsOfStudyMap[abbreviation];
    }
  }
  program.displayName = degree;
  entry.program = program;

  // Add dates — graduated before or around hire year
  // Duration: 1-4 years depending on degree type
  const durationMap = { 'PhD': 4, 'MBA': 2, 'MSc': 2, 'MA': 2, 'BSc': 3, 'BA': 3, 'BEng': 4, 'MEng': 2, 'LLM': 1 };
  const duration = abbreviation ? (durationMap[abbreviation] || 1) : 1;
  const gradYear = hireYear - 2 - (hashStr(seed + ':grad') % 6); // graduated 2-7 years before hire
  const gradMonth = 1 + (hashStr(seed + ':gm') % 12);

  entry.completionMonthYear = `${gradYear}-${String(gradMonth).padStart(2, '0')}-01`;
  entry.startMonthYear = `${gradYear - duration}-${String(1 + (hashStr(seed + ':sm') % 12)).padStart(2, '0')}-01`;
  entry.endMonthYear = entry.completionMonthYear;

  return entry;
}

// ============================================================
// Publication enrichment
// ============================================================

const publishers = [
  'Harvard Business Review', 'European Publishing Quarterly', 'Content Strategy Journal',
  'Digital Communications Review', 'Nordic Business Press', 'Creative Industries Monthly',
  'Mediterranean Business Review', 'Central European Media Journal', 'Springer',
  'Oxford University Press', 'Routledge', 'De Gruyter', 'Elsevier',
];

function buildPublication(pubStr, userSeed, idx) {
  const seed = `${userSeed}:pub:${idx}`;
  const match = pubStr.match(/^(.+?)\s*\((.+?)\s+(\d{4})\)\s*$/);

  const entry = {
    "@odata.type": "#microsoft.graph.itemPublication",
  };

  if (match) {
    entry.displayName = match[1].trim();
    entry.publisher = match[2].trim();
    const year = parseInt(match[3]);
    const month = 1 + (hashStr(seed + ':mo') % 12);
    const day = 1 + (hashStr(seed + ':dy') % 28);
    entry.publishedDate = `${year}-${String(month).padStart(2, '0')}-${String(day).padStart(2, '0')}`;
  } else {
    entry.displayName = pubStr.trim();
    // ~60% still get a publisher and date
    if (seededRand(seed + ':hasPub') < 0.6) {
      entry.publisher = pick(publishers, seed + ':pub');
      const year = 2022 + (hashStr(seed + ':yr') % 4);
      const month = 1 + (hashStr(seed + ':mo') % 12);
      entry.publishedDate = `${year}-${String(month).padStart(2, '0')}-01`;
    }
  }

  // Add description for ~50%
  if (seededRand(seed + ':desc') < 0.5) {
    entry.description = `Published work on ${entry.displayName.toLowerCase().replace(/^the\s+/i, '')}.`;
  }

  return entry;
}

// ============================================================
// Patent enrichment — pending vs issued
// ============================================================

const patentAuthorities = [
  'European Patent Office',
  'EPO',
  'World Intellectual Property Organization',
  'WIPO',
  'UK Intellectual Property Office',
  'German Patent and Trade Mark Office',
  'French National Institute of Industrial Property',
  'Swedish Patent and Registration Office',
  'Spanish Patent and Trademark Office',
  'Italian Patent and Trademark Office',
];

function buildPatent(patStr, userSeed, idx) {
  const seed = `${userSeed}:patent:${idx}`;
  const name = patStr.trim();

  const entry = {
    "@odata.type": "#microsoft.graph.itemPatent",
    "displayName": name,
  };

  // Extract patent number from name if present, e.g. "(EP2024/007123)"
  const numMatch = name.match(/\(([A-Z]{2}\d{4}\/\d+)\)/);
  if (numMatch) {
    entry.number = numMatch[1];
    // Clean displayName to not repeat the number
    entry.displayName = name.replace(/\s*\([A-Z]{2}\d{4}\/\d+\)/, '').trim();
  }

  // Add description
  entry.description = `Patent covering innovations in ${entry.displayName.toLowerCase()}.`;

  // Determine pending vs issued — ~35% pending, ~65% issued
  const isPending = seededRand(seed + ':pending') < 0.35;
  entry.isPending = isPending;

  if (isPending) {
    // Pending patents: no issuedDate, but we know the filing context
    // Some pending patents have a number, some don't yet
    if (!entry.number && seededRand(seed + ':hasNum') < 0.5) {
      const yr = 2024 + (hashStr(seed + ':yr') % 2);
      const seq = 100000 + (hashStr(seed + ':seq') % 900000);
      entry.number = `EP${yr}/${String(seq).padStart(6, '0')}`;
    }
    entry.issuingAuthority = pick(patentAuthorities, seed + ':auth');
  } else {
    // Issued patents: have issuedDate + issuingAuthority
    const year = 2020 + (hashStr(seed + ':yr') % 6);
    const month = 1 + (hashStr(seed + ':mo') % 12);
    const day = 1 + (hashStr(seed + ':dy') % 28);
    entry.issuedDate = `${year}-${String(month).padStart(2, '0')}-${String(day).padStart(2, '0')}`;
    entry.issuingAuthority = pick(patentAuthorities, seed + ':auth');

    // Ensure number exists for issued patents
    if (!entry.number) {
      const seq = 100000 + (hashStr(seed + ':seq') % 900000);
      entry.number = `EP${year}/${String(seq).padStart(6, '0')}`;
    }
  }

  // ~40% get a webUrl
  if (seededRand(seed + ':url') < 0.4 && entry.number) {
    entry.webUrl = `https://worldwide.espacenet.com/patent/search?q=${entry.number.replace('/', '')}`;
  }

  return entry;
}

// ============================================================
// Main conversion
// ============================================================

const csv = fs.readFileSync(csvPath, 'utf-8');
const rows = parseCSV(csv);

console.log(`Parsed ${rows.length} users from CSV`);

const result = rows.map((row, rowIdx) => {
  const userSeed = row.email || `user${rowIdx}`;
  const hireYear = row.employeeHireDate ? parseInt(row.employeeHireDate.slice(0, 4)) : 2020;

  const skills = parseList(row.skills).map((s, i) => buildSkill(s, userSeed, i));
  const languages = parseList(row.languages).map(parseLanguage);

  const notes = [];
  if (row.aboutMe && row.aboutMe.trim()) {
    notes.push({
      "@odata.type": "#microsoft.graph.personAnnotation",
      "displayName": "About Me",
      "detail": {
        "contentType": "text",
        "content": row.aboutMe.trim(),
      },
    });
  }

  const interests = parseList(row.interests).map((s, i) => buildInterest(s, userSeed, i));
  const projects = parseList(row.projects).map((s, i) => buildProject(s, userSeed, i, hireYear));
  const responsibilities = parseList(row.responsibilities).map((s, i) => buildResponsibility(s, userSeed, i));
  const certifications = parseList(row.certifications).map((s, i) => buildCertification(s, userSeed, i, hireYear));
  const awards = parseList(row.awards).map((s, i) => buildAward(s, userSeed, i));
  const educationalActivities = parseList(row.educationalActivities).map((s, i) => buildEducation(s, userSeed, i, hireYear));
  const publications = parseList(row.publications).map((s, i) => buildPublication(s, userSeed, i));
  const patents = parseList(row.patents).map((s, i) => buildPatent(s, userSeed, i));

  return {
    userPrincipalName: row.email,
    displayName: row.name,
    profile: {
      skills,
      languages,
      notes,
      interests,
      projects,
      responsibilities,
      certifications,
      awards,
      educationalActivities,
      publications,
      patents,
    },
  };
});

fs.writeFileSync(outPath, JSON.stringify(result, null, 2), 'utf-8');
console.log(`Wrote ${result.length} users to ${outPath}`);

// ============================================================
// Summary + enrichment stats
// ============================================================

let totalEntities = 0;
const collections = ['skills', 'languages', 'notes', 'interests', 'projects', 'responsibilities', 'certifications', 'awards', 'educationalActivities', 'publications', 'patents'];
const counts = {};
collections.forEach(c => { counts[c] = 0; });
result.forEach(u => {
  collections.forEach(c => {
    counts[c] += u.profile[c].length;
    totalEntities += u.profile[c].length;
  });
});
console.log('\nCollection totals:');
collections.forEach(c => console.log(`  ${c}: ${counts[c]}`));
console.log(`\nTotal entities: ${totalEntities}`);

// Enrichment stats
let certsWithDates = 0, certsNoDates = 0, certsWithExpiry = 0;
let awardsWithDates = 0, awardsNoDates = 0, awardsWithIssuer = 0;
let patentsPending = 0, patentsIssued = 0;
let skillsWithProf = 0, skillsWithCollab = 0;
let pubsWithDate = 0, pubsWithDesc = 0;
let edusWithDates = 0;

result.forEach(u => {
  u.profile.certifications.forEach(c => {
    if (c.issuedDate) certsWithDates++;
    else certsNoDates++;
    if (c.endDate) certsWithExpiry++;
  });
  u.profile.awards.forEach(a => {
    if (a.issuedDate) awardsWithDates++;
    else awardsNoDates++;
    if (a.issuingAuthority) awardsWithIssuer++;
  });
  u.profile.patents.forEach(p => {
    if (p.isPending) patentsPending++;
    else patentsIssued++;
  });
  u.profile.skills.forEach(s => {
    if (s.proficiency) skillsWithProf++;
    if (s.collaborationTags && s.collaborationTags.length) skillsWithCollab++;
  });
  u.profile.publications.forEach(p => {
    if (p.publishedDate) pubsWithDate++;
    if (p.description) pubsWithDesc++;
  });
  u.profile.educationalActivities.forEach(e => {
    if (e.startMonthYear) edusWithDates++;
  });
});

console.log('\nEnrichment breakdown:');
console.log(`  Skills: ${skillsWithProf} with proficiency, ${skillsWithCollab} with collaborationTags`);
console.log(`  Certs: ${certsWithDates} with dates, ${certsNoDates} without dates, ${certsWithExpiry} with expiry`);
console.log(`  Awards: ${awardsWithDates} with dates, ${awardsNoDates} without dates, ${awardsWithIssuer} with issuingAuthority`);
console.log(`  Patents: ${patentsIssued} issued (with issuedDate), ${patentsPending} pending`);
console.log(`  Publications: ${pubsWithDate} with publishedDate, ${pubsWithDesc} with description`);
console.log(`  Education: ${edusWithDates} with start/end dates and fieldsOfStudy`);
