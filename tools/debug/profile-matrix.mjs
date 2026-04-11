#!/usr/bin/env node
/**
 * Fetch all user profiles and display a matrix of which collections have data.
 * Use this to track connector data propagation over time — run it daily after
 * Option B ingestion to see profiles gradually filling in (6-48h per cycle).
 *
 * Usage:
 *   node tools/debug/profile-matrix.mjs --json config/users.config.json
 *   node tools/debug/profile-matrix.mjs --json config/users.config.json --connection-id m365people03
 *
 * Requires a valid delegated token. Run: npm run test-connection (to refresh)
 */

import fs from 'fs/promises';

const TOKEN_CACHE_PATH = `${process.env.HOME}/.m365-provision/token-cache.json`;
const MAX_RETRIES = 3;
const RETRY_DELAY_MS = 2000;
const BATCH_PAUSE_MS = 500;
const BATCH_SIZE = 10;

const COLLECTIONS = [
  'skills', 'interests', 'certifications', 'awards', 'projects',
  'educationalActivities', 'languages', 'publications', 'patents',
  'names', 'positions', 'addresses', 'emails', 'phones', 'notes',
];

const SHORT = {
  skills: 'SKL', interests: 'INT', certifications: 'CRT', awards: 'AWD',
  projects: 'PRJ', educationalActivities: 'EDU', languages: 'LNG',
  publications: 'PUB', patents: 'PAT', names: 'NAM', positions: 'POS',
  addresses: 'ADR', emails: 'EML', phones: 'PHN', notes: 'NTE',
};

async function getToken() {
  const tokenData = JSON.parse(await fs.readFile(TOKEN_CACHE_PATH, 'utf-8'));
  if (new Date(tokenData.expiresOn * 1000) < new Date()) {
    console.error('Token expired. Run: npm run test-connection');
    process.exit(1);
  }
  return tokenData.accessToken;
}

async function sleep(ms) {
  return new Promise(r => setTimeout(r, ms));
}

async function fetchProfile(email, token) {
  for (let attempt = 1; attempt <= MAX_RETRIES; attempt++) {
    try {
      const res = await fetch(
        `https://graph.microsoft.com/beta/users/${encodeURIComponent(email)}/profile`,
        { headers: { Authorization: `Bearer ${token}` } }
      );

      // Throttled — back off and retry
      if (res.status === 429) {
        const retryAfter = parseInt(res.headers.get('Retry-After') || '5', 10);
        process.stderr.write(`  (throttled, waiting ${retryAfter}s...)\r`);
        await sleep(retryAfter * 1000);
        continue;
      }

      if (!res.ok) {
        if (attempt < MAX_RETRIES) {
          await sleep(RETRY_DELAY_MS * attempt);
          continue;
        }
        return null;
      }

      return await res.json();
    } catch (err) {
      if (attempt < MAX_RETRIES) {
        await sleep(RETRY_DELAY_MS * attempt);
        continue;
      }
      return null;
    }
  }
  return null;
}

async function getEmails(inputPath) {
  const raw = await fs.readFile(inputPath, 'utf-8');
  const data = JSON.parse(raw);

  if (data[0]?.MailNickName) {
    const { config } = await import('dotenv');
    config();
    const domain = process.env.USER_DOMAIN || 'yourdomain.onmicrosoft.com';
    return data.map(r => ({
      email: r.MailNickName.includes('@') ? r.MailNickName : `${r.MailNickName}@${domain}`,
      name: r.DisplayName,
    }));
  }
  return data.map(r => ({ email: r.email, name: r.name || r.displayName }));
}

// Parse args
const args = process.argv.slice(2);
let inputPath = null;
let connectionFilter = null;

for (let i = 0; i < args.length; i++) {
  if (args[i] === '--json' || args[i] === '--csv') inputPath = args[++i];
  if (args[i] === '--connection-id') connectionFilter = args[++i];
}

if (!inputPath) {
  console.log('Usage: node tools/debug/profile-matrix.mjs --json config/users.config.json [--connection-id m365people03]');
  console.log('');
  console.log('Displays a matrix of profile collections populated from the connector.');
  console.log('Run daily after Option B ingestion to track propagation progress.');
  process.exit(1);
}

const token = await getToken();
const users = await getEmails(inputPath);
const startTime = Date.now();

console.log(`Fetching profiles for ${users.length} users (with retry)...\n`);

// Header
const nameWidth = 30;
const header = 'Name'.padEnd(nameWidth) + COLLECTIONS.map(c => SHORT[c]).join(' ') + '  Source';
console.log(header);
console.log('─'.repeat(header.length));

let totals = {};
for (const c of COLLECTIONS) totals[c] = 0;
let connectorCount = 0;
let fetchedCount = 0;
let failedCount = 0;

for (const user of users) {
  const profile = await fetchProfile(user.email, token);
  fetchedCount++;

  if (!profile) {
    failedCount++;
    console.log(user.name.substring(0, nameWidth - 1).padEnd(nameWidth) + COLLECTIONS.map(() => ' · ').join('') + '  (failed)');
    continue;
  }

  let hasConnector = false;
  const cells = [];

  for (const c of COLLECTIONS) {
    const items = profile[c] || [];
    const fromConn = items.filter(i =>
      (i.sources || []).some(s => {
        if (connectionFilter) return (s.sourceId || '').includes(connectionFilter);
        return (s.sourceId || '').startsWith('m365people');
      })
    );
    const fromOther = items.length - fromConn.length;

    if (fromConn.length > 0) {
      cells.push(` ${fromConn.length} `);
      totals[c] += fromConn.length;
      hasConnector = true;
    } else if (fromOther > 0) {
      cells.push(' . ');
    } else {
      cells.push(' · ');
    }
  }

  if (hasConnector) connectorCount++;

  const src = hasConnector ? 'CONN' : 'aad';
  const name = user.name.substring(0, nameWidth - 1).padEnd(nameWidth);
  console.log(name + cells.join('') + '  ' + src);

  // Pause between batches to avoid throttling
  if (fetchedCount % BATCH_SIZE === 0) {
    process.stderr.write(`  (${fetchedCount}/${users.length})...\r`);
    await sleep(BATCH_PAUSE_MS);
  }
}

// Totals
const elapsed = ((Date.now() - startTime) / 1000).toFixed(0);
console.log('─'.repeat(header.length));
console.log('TOTALS'.padEnd(nameWidth) + COLLECTIONS.map(c => {
  const v = totals[c];
  return v > 0 ? String(v).padStart(2).padEnd(3) : ' · ';
}).join(''));

console.log('');
console.log(`${connectorCount}/${users.length} users have connector data (${Math.round(connectorCount / users.length * 100)}%)`);
if (failedCount > 0) {
  console.log(`${failedCount} fetch failures (retried ${MAX_RETRIES}x each — likely throttling or mailbox not provisioned)`);
}
console.log(`Completed in ${elapsed}s`);
console.log('');
console.log('Legend: N = items from connector, . = data from other source, · = empty');
console.log('Tip: re-run daily to track propagation. Expect 6-48h per CAPIv2 export cycle.');
