#!/usr/bin/env node
/**
 * Verify that Skills from users.config.json are searchable via Microsoft Search.
 *
 * For each skill in the config, queries Microsoft Search (person entity) and
 * compares the people who return against the people who were supposed to have
 * that skill based on the config.
 *
 * Usage:
 *   node tools/debug/verify-skills-searchable.mjs                  # top-10 most common skills
 *   node tools/debug/verify-skills-searchable.mjs --all            # every unique skill
 *   node tools/debug/verify-skills-searchable.mjs --skill Python   # single skill
 *   node tools/debug/verify-skills-searchable.mjs --sample 20      # first N skills
 *   node tools/debug/verify-skills-searchable.mjs --rare           # skills with exactly 1 owner
 *
 * Token is read from ~/.m365-provision/token-cache.json.
 * If expired, run: npm run test-connection
 */

import fs from 'fs/promises';
import { readFileSync } from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';
import dotenv from 'dotenv';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const ROOT = path.resolve(__dirname, '../..');
dotenv.config({ path: path.join(ROOT, '.env') });

const TOKEN_CACHE = `${process.env.HOME}/.m365-provision/token-cache.json`;
const CONFIG_PATH = path.join(ROOT, 'config/users.config.json');
const USER_DOMAIN = process.env.USER_DOMAIN;

async function getToken() {
  const raw = JSON.parse(await fs.readFile(TOKEN_CACHE, 'utf-8'));
  const expires = new Date(raw.expiresOn * 1000);
  if (expires < new Date()) {
    console.error(`Token expired ${expires.toLocaleString()} — run: npm run test-connection`);
    process.exit(1);
  }
  return raw.accessToken;
}

function parseArgs() {
  const args = process.argv.slice(2);
  const out = { mode: 'top', limit: 10, skill: null };
  for (let i = 0; i < args.length; i++) {
    if (args[i] === '--all') out.mode = 'all';
    else if (args[i] === '--rare') out.mode = 'rare';
    else if (args[i] === '--skill') { out.mode = 'single'; out.skill = args[++i]; }
    else if (args[i] === '--sample') { out.mode = 'sample'; out.limit = Number(args[++i]); }
    else if (args[i] === '--top') { out.mode = 'top'; out.limit = Number(args[++i] ?? 10); }
  }
  return out;
}

function loadSkillIndex() {
  const raw = JSON.parse(readFileSync(CONFIG_PATH, 'utf-8'));
  const users = Array.isArray(raw) ? raw : (raw.users || raw.Users || []);
  // skill DisplayName → Set<UPN>
  const index = new Map();
  for (const u of users) {
    const upn = (u.UserPrincipalName || `${u.MailNickName}@${USER_DOMAIN}`).toLowerCase();
    for (const s of (u.Skills || u.skills || [])) {
      const name = s.DisplayName || s.displayName;
      if (!name) continue;
      if (!index.has(name)) index.set(name, new Set());
      index.get(name).add(upn);
    }
  }
  return { index, totalUsers: users.length };
}

async function searchBySkill(token, skill) {
  const body = {
    requests: [{
      entityTypes: ['person'],
      query: { queryString: `skill:"${skill.replace(/"/g, '\\"')}"` },
      from: 0,
      size: 50,
    }],
  };
  const res = await fetch('https://graph.microsoft.com/v1.0/search/query', {
    method: 'POST',
    headers: {
      Authorization: `Bearer ${token}`,
      'Content-Type': 'application/json',
    },
    body: JSON.stringify(body),
  });
  if (!res.ok) {
    const err = await res.text();
    return { error: `${res.status} ${err.slice(0, 200)}`, hits: [] };
  }
  const json = await res.json();
  const hits = [];
  for (const r of json.value || []) {
    for (const c of r.hitsContainers || []) {
      for (const h of c.hits || []) {
        const upn = h.resource?.userPrincipalName || h.resource?.scoredEmailAddresses?.[0]?.address || '';
        if (upn) hits.push(upn.toLowerCase());
      }
    }
  }
  return { hits };
}

function pickSkills({ index }, args) {
  const entries = [...index.entries()].map(([name, owners]) => ({ name, owners }));
  entries.sort((a, b) => b.owners.size - a.owners.size);
  if (args.mode === 'all') return entries;
  if (args.mode === 'rare') return entries.filter(e => e.owners.size === 1);
  if (args.mode === 'single') {
    const hit = entries.find(e => e.name.toLowerCase() === args.skill.toLowerCase());
    return hit ? [hit] : [{ name: args.skill, owners: new Set() }];
  }
  return entries.slice(0, args.limit); // top / sample
}

function fmtPct(n, total) {
  if (total === 0) return '—';
  return `${Math.round((n / total) * 100)}%`;
}

async function main() {
  const args = parseArgs();
  const token = await getToken();
  const { index, totalUsers } = loadSkillIndex();

  console.log(`Config: ${CONFIG_PATH}`);
  console.log(`Users in config: ${totalUsers}`);
  console.log(`Unique skills: ${index.size}`);
  console.log(`Mode: ${args.mode}${args.limit ? ` (limit ${args.limit})` : ''}${args.skill ? ` (${args.skill})` : ''}`);
  console.log('─'.repeat(90));

  const skills = pickSkills({ index }, args);
  const summary = [];

  for (const { name, owners } of skills) {
    const expected = owners;
    const { hits, error } = await searchBySkill(token, name);

    if (error) {
      console.log(`${name.padEnd(35)}  ERROR: ${error}`);
      summary.push({ name, expected: expected.size, matched: 0, extra: 0, missing: expected.size, error });
      continue;
    }

    const hitSet = new Set(hits);
    const matched = [...expected].filter(u => hitSet.has(u)).length;
    const missing = [...expected].filter(u => !hitSet.has(u));
    const extra = [...hitSet].filter(u => !expected.has(u));

    const coverage = fmtPct(matched, expected.size);
    console.log(
      `${name.padEnd(35)}  expected=${String(expected.size).padStart(3)} ` +
      `hits=${String(hits.length).padStart(3)} matched=${String(matched).padStart(3)} (${coverage}) ` +
      `missing=${missing.length} extra=${extra.length}`
    );

    if (missing.length && missing.length <= 5) {
      console.log(`    missing: ${missing.map(u => u.split('@')[0]).join(', ')}`);
    }
    summary.push({ name, expected: expected.size, matched, extra: extra.length, missing: missing.length });

    // Gentle pacing — Search API is throttled
    await new Promise(r => setTimeout(r, 150));
  }

  // Totals
  const totalExpected = summary.reduce((a, r) => a + r.expected, 0);
  const totalMatched = summary.reduce((a, r) => a + r.matched, 0);
  console.log('─'.repeat(90));
  console.log(`Totals  skills=${summary.length}  expected=${totalExpected}  matched=${totalMatched}  coverage=${fmtPct(totalMatched, totalExpected)}`);

  const zero = summary.filter(r => r.matched === 0 && r.expected > 0);
  if (zero.length) {
    console.log(`\nSkills with 0 coverage (${zero.length}):`);
    for (const r of zero) console.log(`  - ${r.name} (expected ${r.expected})`);
  }
}

main().catch(err => {
  console.error(err);
  process.exit(1);
});
