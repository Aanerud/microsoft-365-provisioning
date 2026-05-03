#!/usr/bin/env node
/**
 * Manager coverage matrix — shows, per user:
 *   EXPECTED  — the manager named in users.config.json (Option A)
 *   ENTRA     — the manager link in Azure AD (/users/{id}/manager)
 *   PAPI      — the manager relatedPerson in the profile's first position
 *
 * Answers the questions: did Option A set the manager link? Is the connector
 * writing the manager relatedPerson so Copilot can reason about it?
 *
 * Usage:
 *   node tools/debug/manager-matrix.mjs
 *   node tools/debug/manager-matrix.mjs --config config/users.config.json
 *   node tools/debug/manager-matrix.mjs --user sofia.j
 *   node tools/debug/manager-matrix.mjs --concurrency 8
 *   node tools/debug/manager-matrix.mjs --help
 *
 * Requires a valid delegated token (npm run test-connection to refresh).
 */

import { readFileSync } from 'fs';
import fs from 'fs/promises';
import path from 'path';
import { fileURLToPath } from 'url';
import dotenv from 'dotenv';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const ROOT = path.resolve(__dirname, '../..');
dotenv.config({ path: path.join(ROOT, '.env') });

const TOKEN_CACHE = `${process.env.HOME}/.m365-provision/token-cache.json`;
const DEFAULT_CONFIG = path.join(ROOT, 'config/users.config.json');
const OID_CACHE_PATH = path.join(ROOT, 'config/users.connector.config_oid_cache.json');
const USER_DOMAIN = process.env.USER_DOMAIN;
const MIN_TTL_SECONDS = 300;

// Build reverse OID → UPN/nick map from the connector OID cache. PAPI returns
// the manager relatedPerson with only `userId` populated (displayName and UPN
// are wiped by PCP), so we have to resolve back through the cache.
function loadOidReverseMap() {
  try {
    const raw = JSON.parse(readFileSync(OID_CACHE_PATH, 'utf-8'));
    const users = raw.users || raw;
    const rev = new Map(); // oid → { upn, nick }
    for (const [upn, oid] of Object.entries(users)) {
      if (typeof oid !== 'string') continue;
      const nick = (upn.split('@')[0] || '').toLowerCase();
      rev.set(oid.toLowerCase(), { upn, nick });
    }
    return rev;
  } catch {
    return new Map();
  }
}

const HELP = `
manager-matrix — per-user manager coverage (config vs Entra vs PAPI)

Usage:
  node tools/debug/manager-matrix.mjs [options]

Options:
  --config <path>       Users config JSON (default: config/users.config.json)
  --user <mailNickname> Only check this user
  --concurrency <N>     Concurrent Graph fetches (default: 8)
  --help                Show this message
`;

function parseArgs() {
  const args = process.argv.slice(2);
  const out = { config: DEFAULT_CONFIG, user: null, concurrency: 8, help: false };
  for (let i = 0; i < args.length; i++) {
    const a = args[i];
    if (a === '--help' || a === '-h') out.help = true;
    else if (a === '--config') out.config = args[++i];
    else if (a === '--user') out.user = args[++i];
    else if (a === '--concurrency') out.concurrency = Number(args[++i]) || 8;
    else { console.error(`Unknown arg: ${a}`); console.error(HELP); process.exit(2); }
  }
  return out;
}

async function loadToken() {
  let raw;
  try { raw = JSON.parse(await fs.readFile(TOKEN_CACHE, 'utf-8')); }
  catch { console.error('No cached token. Run: npm run test-connection'); process.exit(1); }
  const ttl = Math.floor((new Date(raw.expiresOn * 1000) - new Date()) / 1000);
  if (ttl <= 0) { console.error('Token expired. Run: npm run test-connection'); process.exit(1); }
  if (ttl < MIN_TTL_SECONDS) {
    console.error(`Token expires in ${ttl}s (need ≥${MIN_TTL_SECONDS}s). Run: npm run test-connection`);
    process.exit(1);
  }
  return raw.accessToken;
}

function loadConfig(configPath) {
  const raw = JSON.parse(readFileSync(configPath, 'utf-8'));
  const users = Array.isArray(raw) ? raw : (raw.users || raw.Users || []);
  if (!users.length) { console.error(`Config ${configPath} has no users.`); process.exit(1); }
  return users;
}

function upnFor(user) {
  const nick = user.MailNickName || user.mailNickName || user.mailNickname;
  if (!nick) return null;
  if (nick.includes('@')) return nick;
  if (!USER_DOMAIN) { console.error('USER_DOMAIN not set in .env'); process.exit(1); }
  return `${nick}@${USER_DOMAIN}`;
}

async function sleep(ms) { return new Promise(r => setTimeout(r, ms)); }

async function graphGet(url, token) {
  for (let attempt = 0; attempt < 2; attempt++) {
    let res;
    try { res = await fetch(url, { headers: { Authorization: `Bearer ${token}` } }); }
    catch (err) { return { error: 'NETWORK', detail: err.message }; }
    if (res.status === 404) return { notFound: true };
    if (res.status === 401) { console.error('\nToken rejected. Run: npm run test-connection'); process.exit(1); }
    if (res.status === 429) {
      if (attempt === 1) return { error: 'THROTTLED' };
      const wait = parseInt(res.headers.get('Retry-After') || '5', 10) * 1000;
      await sleep(wait);
      continue;
    }
    if (!res.ok) return { error: `HTTP_${res.status}`, detail: (await res.text().catch(() => '')).slice(0, 120) };
    return { data: await res.json() };
  }
  return { error: 'EXHAUSTED_RETRIES' };
}

// Entra: GET /users/{upn}/manager → { id, displayName, userPrincipalName, ... }
async function fetchEntraManager(upn, token) {
  const url = `https://graph.microsoft.com/v1.0/users/${encodeURIComponent(upn)}/manager`;
  const res = await graphGet(url, token);
  if (res.notFound) return { present: false };
  if (res.error) return { present: false, error: res.error, detail: res.detail };
  return {
    present: true,
    displayName: res.data.displayName,
    userPrincipalName: res.data.userPrincipalName,
  };
}

// PAPI: GET /beta/users/{upn}/profile/positions → {value: [workPosition]}
// There are typically two position entries: one AAD-sourced (from Entra) with
// empty manager/colleagues, and one connector-sourced (from our ingest) with
// the relatedPerson data. We pick the connector-sourced one.
//
// Key observation: PCP stores the manager with ONLY `userId` populated —
// displayName is "" and userPrincipalName is null. Resolution must go
// through the OID reverse map.
async function fetchPapiPositions(upn, token, oidReverse) {
  const url = `https://graph.microsoft.com/beta/users/${encodeURIComponent(upn)}/profile/positions`;
  const res = await graphGet(url, token);
  if (res.notFound) return { present: false, reason: 'profile-or-positions-not-found' };
  if (res.error) return { present: false, error: res.error, detail: res.detail };
  const positions = res.data.value || [];
  if (!positions.length) return { present: false, reason: 'positions-empty' };

  // Prefer the position whose source is our connector. Otherwise fall back.
  const connector = positions.find(p => {
    const srcTypes = p.source?.type || [];
    return srcTypes.some(t => /agent provisioning|connector/i.test(t));
  });
  const current = connector || positions.find(p => p.isCurrent) || positions[0];

  // Manager appears directly as `current.manager` after PCP processing.
  const mgr = current.manager;
  if (!mgr || typeof mgr !== 'object') return { present: false, reason: 'manager-null-or-missing' };

  const rel = (mgr.relationship || '').toLowerCase();
  const userId = mgr.userId;
  const displayName = mgr.displayName || null;
  const userPrincipalName = mgr.userPrincipalName || null;

  // Resolve userId → UPN/nick through our OID cache since PCP typically
  // wipes display fields.
  const resolved = userId ? oidReverse.get(String(userId).toLowerCase()) : null;

  // Consider manager "present" if we have EITHER a usable identifier
  // (userId, UPN, or displayName) OR relationship === 'manager'.
  const hasIdentifier = !!(userId || userPrincipalName || displayName);
  if (!hasIdentifier && rel !== 'manager') return { present: false, reason: 'empty-manager-object' };

  return {
    present: true,
    relationship: rel || 'manager',
    userId,
    displayName,
    userPrincipalName,
    resolvedUpn: resolved?.upn || null,
    resolvedNick: resolved?.nick || null,
    source: current.source?.type?.[0] || 'unknown',
    connectorSourced: !!connector,
  };
}

function nickFromUpn(upn) {
  if (!upn || typeof upn !== 'string') return null;
  return upn.split('@')[0].toLowerCase();
}

// papi object from fetchPapiPositions — uses resolvedNick when PAPI wiped the UPN
function compare(expectedNick, entraUpn, papi) {
  const e = nickFromUpn(entraUpn);
  const p = papi?.resolvedNick || nickFromUpn(papi?.userPrincipalName);
  const exp = expectedNick ? expectedNick.toLowerCase() : null;
  return {
    configHas: !!exp,
    entraHas: !!e,
    papiHas: !!(papi && papi.present),
    papiResolved: !!p,
    entraMatch: !!(exp && e && e === exp),
    papiMatch: !!(exp && p && p === exp),
    entraConflict: !!(exp && e && e !== exp),
    papiConflict: !!(exp && p && p !== exp),
  };
}

async function runPool(items, concurrency, worker) {
  const results = new Array(items.length);
  let i = 0;
  async function next() {
    while (i < items.length) {
      const mine = i++;
      results[mine] = await worker(items[mine], mine);
    }
  }
  await Promise.all(Array.from({ length: Math.min(concurrency, items.length) }, next));
  return results;
}

// ───── main ─────

async function main() {
  const args = parseArgs();
  if (args.help) { console.log(HELP); process.exit(0); }

  const allUsers = loadConfig(args.config);
  let targets = allUsers;
  if (args.user) {
    targets = allUsers.filter(u => (u.MailNickName || u.mailNickName) === args.user);
    if (!targets.length) { console.error(`User "${args.user}" not in config.`); process.exit(1); }
  }

  const token = await loadToken();
  const oidReverse = loadOidReverseMap();
  console.log(`Config: ${args.config}`);
  console.log(`OID reverse-map entries: ${oidReverse.size}`);
  console.log(`Checking ${targets.length} user(s) with concurrency=${args.concurrency}`);
  console.log('');

  const results = await runPool(targets, args.concurrency, async (u) => {
    const upn = upnFor(u);
    const expected = u.Manager || u.manager || null;
    const [entra, papi] = await Promise.all([
      fetchEntraManager(upn, token),
      fetchPapiPositions(upn, token, oidReverse),
    ]);
    return { nick: u.MailNickName || u.mailNickName, upn, expected, entra, papi };
  });

  // ── per-user table ──
  console.log('─'.repeat(115));
  console.log(`${'USER'.padEnd(15)} ${'EXPECTED'.padEnd(14)} ${'ENTRA'.padEnd(14)} ${'PAPI'.padEnd(30)} STATUS`);
  console.log('─'.repeat(115));

  const counts = {
    total: results.length,
    configHas: 0, entraHas: 0, papiHas: 0,
    entraMatch: 0, papiMatch: 0,
    bothMatch: 0,
    entraMissing: 0, papiMissing: 0,
    entraConflict: 0, papiConflict: 0,
  };
  const issues = [];

  for (const r of results) {
    const cmp = compare(r.expected, r.entra.userPrincipalName, r.papi);
    if (cmp.configHas) counts.configHas++;
    if (cmp.entraHas) counts.entraHas++;
    if (cmp.papiHas) counts.papiHas++;
    if (cmp.entraMatch) counts.entraMatch++;
    if (cmp.papiMatch) counts.papiMatch++;
    if (cmp.entraMatch && cmp.papiMatch) counts.bothMatch++;
    if (cmp.configHas && !cmp.entraHas) counts.entraMissing++;
    if (cmp.configHas && !cmp.papiHas) counts.papiMissing++;
    if (cmp.entraConflict) counts.entraConflict++;
    if (cmp.papiConflict) counts.papiConflict++;

    const entraDisplay = r.entra.present ? nickFromUpn(r.entra.userPrincipalName) || '?' : (r.entra.error || '—');

    let papiDisplay;
    if (!r.papi.present) {
      papiDisplay = r.papi.error || r.papi.reason || '—';
    } else if (r.papi.resolvedNick) {
      papiDisplay = r.papi.resolvedNick + ' (via userId)';
    } else if (r.papi.userPrincipalName) {
      papiDisplay = nickFromUpn(r.papi.userPrincipalName);
    } else if (r.papi.userId) {
      papiDisplay = 'userId=' + r.papi.userId.slice(0, 8) + '… (unresolved)';
    } else {
      papiDisplay = 'empty-object';
    }

    let status;
    if (!cmp.configHas) status = 'no-manager-in-config';
    else if (cmp.entraMatch && cmp.papiMatch) status = '✓ both-match';
    else if (cmp.entraMatch && cmp.papiHas && !cmp.papiResolved) status = 'entra-ok, papi-unresolved';
    else if (cmp.entraMatch && !cmp.papiHas) status = 'entra-ok, papi-missing';
    else if (cmp.entraMatch && cmp.papiConflict) status = 'entra-ok, papi-WRONG';
    else if (!cmp.entraHas && cmp.papiMatch) status = 'entra-MISSING, papi-ok';
    else if (!cmp.entraHas && !cmp.papiHas) status = 'BOTH MISSING';
    else if (cmp.entraConflict && cmp.papiMatch) status = 'entra-WRONG, papi-ok';
    else if (cmp.entraConflict && cmp.papiConflict) status = 'BOTH WRONG';
    else status = 'mixed';

    if (status !== '✓ both-match' && status !== 'no-manager-in-config') issues.push({ ...r, cmp, status });

    console.log(
      `${String(r.nick).padEnd(15)} ${String(r.expected || '—').padEnd(14)} ${String(entraDisplay).padEnd(14)} ${String(papiDisplay).padEnd(30)} ${status}`
    );
  }

  // ── summary ──
  console.log('');
  console.log('─'.repeat(110));
  console.log('SUMMARY');
  console.log('─'.repeat(110));
  console.log(`Users in scope:            ${counts.total}`);
  console.log(`Config has manager:        ${counts.configHas}`);
  console.log(`Entra has manager:         ${counts.entraHas}  (matches config: ${counts.entraMatch})`);
  console.log(`PAPI position has manager: ${counts.papiHas}  (matches config: ${counts.papiMatch})`);
  console.log(`Both Entra+PAPI match:     ${counts.bothMatch}`);
  console.log(`Entra missing (config has, Entra doesn't): ${counts.entraMissing}`);
  console.log(`PAPI missing (config has, PAPI doesn't):   ${counts.papiMissing}`);
  console.log(`Entra conflict (wrong manager):            ${counts.entraConflict}`);
  console.log(`PAPI conflict (wrong manager):             ${counts.papiConflict}`);

  if (issues.length) {
    console.log('');
    console.log('─'.repeat(110));
    console.log(`USERS WITH ISSUES (${issues.length})`);
    console.log('─'.repeat(110));
    for (const u of issues.slice(0, 40)) {
      console.log(`  ${u.nick.padEnd(15)} expected=${u.expected || '—'}  entra=${u.entra.userPrincipalName ? nickFromUpn(u.entra.userPrincipalName) : (u.entra.error || 'none')}  papi=${u.papi.userPrincipalName ? nickFromUpn(u.papi.userPrincipalName) : (u.papi.reason || u.papi.error || 'none')}  → ${u.status}`);
    }
    if (issues.length > 40) console.log(`  ... and ${issues.length - 40} more`);
  }

  // Non-zero exit if Entra is missing (Option A responsibility) — PAPI missing is just propagation lag.
  const criticalMissing = counts.entraMissing + counts.entraConflict + counts.papiConflict;
  process.exit(criticalMissing ? 1 : 0);
}

main().catch(err => { console.error(err); process.exit(1); });
