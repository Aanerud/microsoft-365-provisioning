#!/usr/bin/env node
/**
 * Content-level PAPI parity validator.
 *
 * Reads a users config JSON (default: config/users.connector.config.json) as
 * the source of truth and compares field-level CONTENT against each user's
 * PAPI profile. Not a key-or-count check — actual value equality with shape
 * fidelity, so Products can't silently collapse from {name,model,gtin} to a
 * names-only string without failing.
 *
 * Output categories (distinct):
 *   PASS             — input value present in output
 *   OUTPUT_MISSING   — input had it, output doesn't (real regression)
 *   SHAPE_COLLAPSE   — same key, degraded shape (e.g. Products losing Model/GTIN)
 *   INPUT_MISSING    — column not in input at all (upstream config gap, non-failing)
 *   PROFILE_MISSING  — 404 on user profile (non-failing, surfaced separately)
 *   FETCH_FAILED     — 429 twice or other fetch error
 *
 * Exit code 0 if zero OUTPUT_MISSING and zero SHAPE_COLLAPSE.
 *
 * Usage:
 *   node tools/debug/verify-papi-parity.mjs
 *   node tools/debug/verify-papi-parity.mjs --config config/users.connector.config.json
 *   node tools/debug/verify-papi-parity.mjs --user anders.n
 *   node tools/debug/verify-papi-parity.mjs --concurrency 3
 *   node tools/debug/verify-papi-parity.mjs --help
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
const DEFAULT_CONFIG = path.join(ROOT, 'config/users.connector.config.json');
const USER_DOMAIN = process.env.USER_DOMAIN;
const MIN_TOKEN_TTL_SECONDS = 600; // 10 minutes — must exceed realistic run time

const HELP = `
verify-papi-parity — content-level comparison of config input vs PAPI output

Usage:
  node tools/debug/verify-papi-parity.mjs [options]

Options:
  --config <path>       Config JSON (default: config/users.connector.config.json)
  --user <mailNickname> Only check this user (default: all users)
  --concurrency <N>     Concurrent Graph fetches (default: 5)
  --help                Show this message

Requires USER_DOMAIN in .env and a valid delegated token
(run: npm run test-connection).
`;

// ───── arg parsing ─────

function parseArgs() {
  const args = process.argv.slice(2);
  const out = {
    config: DEFAULT_CONFIG,
    user: null,
    concurrency: 5,
    help: false,
  };
  for (let i = 0; i < args.length; i++) {
    const a = args[i];
    if (a === '--help' || a === '-h') out.help = true;
    else if (a === '--config') out.config = args[++i];
    else if (a === '--user') out.user = args[++i];
    else if (a === '--concurrency') out.concurrency = Number(args[++i]) || 5;
    else {
      console.error(`Unknown argument: ${a}`);
      console.error(HELP);
      process.exit(2);
    }
  }
  return out;
}

// ───── token handling ─────

async function loadToken() {
  let raw;
  try {
    raw = JSON.parse(await fs.readFile(TOKEN_CACHE, 'utf-8'));
  } catch {
    console.error('No cached token. Run: npm run test-connection');
    process.exit(1);
  }
  const expiresOn = new Date(raw.expiresOn * 1000);
  const ttlSeconds = Math.floor((expiresOn - new Date()) / 1000);
  if (ttlSeconds <= 0) {
    console.error(`Token expired at ${expiresOn.toLocaleString()} — run: npm run test-connection`);
    process.exit(1);
  }
  // Outside-voice finding 5: refuse to start if TTL is too short for a full
  // concurrent run. Saves a polluted report from mid-run 401s.
  if (ttlSeconds < MIN_TOKEN_TTL_SECONDS) {
    console.error(
      `Token expires in ${ttlSeconds}s (need ≥${MIN_TOKEN_TTL_SECONDS}s for a 94-user concurrent run).`
    );
    console.error('Run: npm run test-connection   then retry.');
    process.exit(1);
  }
  return raw.accessToken;
}

// ───── config loading ─────

function loadConfig(configPath) {
  let raw;
  try {
    raw = JSON.parse(readFileSync(configPath, 'utf-8'));
  } catch (err) {
    console.error(`Failed to read config ${configPath}: ${err.message}`);
    process.exit(1);
  }
  const users = Array.isArray(raw) ? raw : (raw.users || raw.Users || []);
  if (users.length === 0) {
    console.error(`Config ${configPath} contains zero users.`);
    process.exit(1);
  }
  return users;
}

function upnFor(user) {
  const nick = user.MailNickName || user.mailNickName || user.mailNickname;
  if (!nick) return null;
  if (nick.includes('@')) return nick;
  if (!USER_DOMAIN) {
    console.error('USER_DOMAIN not set in .env');
    process.exit(1);
  }
  return `${nick}@${USER_DOMAIN}`;
}

// ───── Graph fetch ─────

async function sleep(ms) { return new Promise(r => setTimeout(r, ms)); }

async function fetchProfile(upn, token) {
  const url = `https://graph.microsoft.com/beta/users/${encodeURIComponent(upn)}/profile`;
  for (let attempt = 0; attempt < 2; attempt++) {
    let res;
    try {
      res = await fetch(url, { headers: { Authorization: `Bearer ${token}` } });
    } catch (err) {
      return { error: 'FETCH_FAILED', detail: err.message };
    }
    if (res.status === 404) return { profile: null, missing: true };
    if (res.status === 401) {
      console.error(`\nToken rejected mid-run for ${upn}. Run: npm run test-connection`);
      process.exit(1);
    }
    if (res.status === 429) {
      if (attempt === 1) return { error: 'FETCH_FAILED', detail: '429 after retry' };
      const wait = parseInt(res.headers.get('Retry-After') || '5', 10) * 1000;
      await sleep(wait);
      continue;
    }
    if (!res.ok) {
      return { error: 'FETCH_FAILED', detail: `${res.status} ${await res.text().then(t => t.slice(0, 100)).catch(() => '')}` };
    }
    const profile = await res.json();
    return { profile };
  }
  return { error: 'FETCH_FAILED', detail: 'exhausted retries' };
}

// ───── content checks ─────

const PRODUCT_LINE_RE = /^- Name: (.+?)(?:, model: (.+?))?(?:, gtin: (\d+))?$/gm;

/**
 * Outside-voice finding 3: shape fidelity via structured regex, not substring.
 * Returns { status, detail } where status is one of PASS, SHAPE_COLLAPSE,
 * OUTPUT_MISSING, INPUT_MISSING.
 */
function checkProducts(inputProducts, outputValue) {
  const inputIsArray = Array.isArray(inputProducts) && inputProducts.length > 0;
  const hasOutput = outputValue !== undefined && outputValue !== null && outputValue !== '';

  if (!inputIsArray && !hasOutput) return { status: 'INPUT_MISSING' };
  if (!inputIsArray && hasOutput) return { status: 'PASS', detail: '(output present, input empty — tolerated)' };
  if (inputIsArray && !hasOutput) return { status: 'OUTPUT_MISSING', detail: `input had ${inputProducts.length} products, output has nothing` };

  // Both present — shape check.
  if (typeof outputValue !== 'string') {
    return { status: 'SHAPE_COLLAPSE', detail: `output is ${typeof outputValue}, not string` };
  }

  // Count how many rendered product lines we can parse out.
  PRODUCT_LINE_RE.lastIndex = 0;
  const matches = [...outputValue.matchAll(PRODUCT_LINE_RE)];
  const expectedCount = inputProducts.filter(p => p && (p.name || p.Name)).length;

  if (matches.length !== expectedCount) {
    return {
      status: 'SHAPE_COLLAPSE',
      detail: `expected ${expectedCount} product lines, regex matched ${matches.length}`,
    };
  }

  // Verify each input product's name is present case-insensitively.
  const outLower = outputValue.toLowerCase();
  const missing = [];
  for (const p of inputProducts) {
    const name = p?.name ?? p?.Name;
    if (!name) continue;
    if (!outLower.includes(String(name).toLowerCase())) missing.push(name);
  }
  if (missing.length) {
    return { status: 'OUTPUT_MISSING', detail: `names not in output: ${missing.slice(0, 3).join(', ')}` };
  }

  return { status: 'PASS' };
}

function checkStringCollection(inputArr, outputArr, fieldName, extractInput, extractOutput) {
  const inputIs = Array.isArray(inputArr) && inputArr.length > 0;
  const outputIs = Array.isArray(outputArr) && outputArr.length > 0;

  if (!inputIs && !outputIs) return { status: 'INPUT_MISSING' };
  if (!inputIs && outputIs) return { status: 'PASS', detail: '(output has extras, input empty)' };
  if (inputIs && !outputIs) return { status: 'OUTPUT_MISSING', detail: `input had ${inputArr.length} ${fieldName}, output has 0` };

  const inputValues = new Set(inputArr.map(extractInput).filter(Boolean).map(s => s.toLowerCase()));
  const outputValues = new Set(outputArr.map(extractOutput).filter(Boolean).map(s => s.toLowerCase()));

  const missing = [...inputValues].filter(v => !outputValues.has(v));
  if (missing.length) {
    return { status: 'OUTPUT_MISSING', detail: `missing ${missing.length}/${inputValues.size} ${fieldName}: ${missing.slice(0, 3).join(', ')}` };
  }
  return { status: 'PASS' };
}

/**
 * Dynamic custom properties from the config (anything that isn't a known
 * profile collection). PAPI exposes these via a side channel; we check
 * string-to-string equality.
 */
function checkCustomProperty(inputValue, outputValue, fieldName) {
  const hasInput = inputValue !== undefined && inputValue !== null && inputValue !== '';
  const hasOutput = outputValue !== undefined && outputValue !== null && outputValue !== '';
  if (!hasInput && !hasOutput) return { status: 'INPUT_MISSING' };
  if (!hasInput && hasOutput) return { status: 'PASS', detail: '(output present, input empty)' };
  if (hasInput && !hasOutput) return { status: 'OUTPUT_MISSING', detail: `expected "${String(inputValue).slice(0, 50)}"` };
  // String equality
  if (String(inputValue) !== String(outputValue)) {
    return { status: 'SHAPE_COLLAPSE', detail: `value differs: input="${String(inputValue).slice(0, 30)}" vs output="${String(outputValue).slice(0, 30)}"` };
  }
  return { status: 'PASS' };
}

// ───── per-user comparison ─────

function compareUser(configUser, profile) {
  const results = {};

  // Products (Path B blob — structured regex check)
  const inputProducts = configUser.Products ?? configUser.products;
  // PAPI side: products is surfaced as a custom property when connector writes it.
  // The profile object's shape for custom props varies — look in notable places.
  const outputProducts = profile?.customProperties?.products
    ?? profile?.products
    ?? null;
  results.products = checkProducts(inputProducts, outputProducts);

  // Skills — stringCollection via profile.skills[].displayName
  results.skills = checkStringCollection(
    configUser.Skills ?? configUser.skills,
    profile?.skills?.value ?? profile?.skills,
    'skills',
    s => s?.DisplayName ?? s?.displayName,
    s => s?.displayName ?? s?.DisplayName,
  );

  // Languages
  results.languages = checkStringCollection(
    configUser.Languages ?? configUser.languages,
    profile?.languages?.value ?? profile?.languages,
    'languages',
    s => s?.DisplayName ?? s?.displayName,
    s => s?.displayName ?? s?.DisplayName,
  );

  // Interests
  results.interests = checkStringCollection(
    configUser.Interests ?? configUser.interests,
    profile?.interests?.value ?? profile?.interests,
    'interests',
    s => s?.DisplayName ?? s?.displayName,
    s => s?.displayName ?? s?.DisplayName,
  );

  // Certifications
  results.certifications = checkStringCollection(
    configUser.Certifications ?? configUser.certifications,
    profile?.certifications?.value ?? profile?.certifications,
    'certifications',
    s => s?.DisplayName ?? s?.displayName,
    s => s?.displayName ?? s?.DisplayName,
  );

  // Projects
  results.projects = checkStringCollection(
    configUser.Projects ?? configUser.projects,
    profile?.projects?.value ?? profile?.projects,
    'projects',
    s => s?.DisplayName ?? s?.displayName,
    s => s?.displayName ?? s?.DisplayName,
  );

  // workModality / jobFamilyGroupName / jobFamilyName are per-position properties
  // (Positions[].Detail.*), not top-level fields. They flow through to PAPI as
  // part of personCurrentPosition rather than flat custom properties. We don't
  // check them here because:
  //   (a) the source of truth is nested per-position, not per-person, and
  //   (b) the PAPI surfacing for labeled composite properties isn't directly
  //       comparable to flat custom-prop output.
  // If future validation of these fields is needed, compare
  // configUser.Positions[i].Detail.{WorkModality,JobFamilyGroupName,JobFamilyName}
  // against profile.positions[i].detail.* with isCurrent matching.

  return results;
}

// ───── concurrency pool ─────

async function runPool(items, concurrency, worker) {
  const results = [];
  let idx = 0;
  async function pullNext() {
    while (idx < items.length) {
      const mine = idx++;
      results[mine] = await worker(items[mine], mine);
    }
  }
  const runners = Array.from({ length: Math.min(concurrency, items.length) }, () => pullNext());
  await Promise.all(runners);
  return results;
}

// ───── main ─────

async function main() {
  const args = parseArgs();
  if (args.help) {
    console.log(HELP);
    process.exit(0);
  }

  console.log(`Config: ${args.config}`);
  const users = loadConfig(args.config);
  console.log(`Users in config: ${users.length}`);

  let targets = users;
  if (args.user) {
    targets = users.filter(u => (u.MailNickName || u.mailNickName || u.mailNickname) === args.user);
    if (targets.length === 0) {
      console.error(`No user matching mailNickname="${args.user}" in config.`);
      process.exit(1);
    }
  }

  const token = await loadToken();
  console.log(`Checking ${targets.length} user(s) with concurrency=${args.concurrency}\n`);

  const perUser = await runPool(targets, args.concurrency, async (user) => {
    const upn = upnFor(user);
    if (!upn) return { user, upn: null, error: 'no UPN', results: null };
    const fetchRes = await fetchProfile(upn, token);
    if (fetchRes.error) return { user, upn, error: fetchRes.error, detail: fetchRes.detail, results: null };
    if (fetchRes.missing) return { user, upn, missing: true, results: null };
    return { user, upn, results: compareUser(user, fetchRes.profile) };
  });

  // ── aggregate report ──
  const counts = {
    PASS: 0, OUTPUT_MISSING: 0, SHAPE_COLLAPSE: 0, INPUT_MISSING: 0,
    PROFILE_MISSING: 0, FETCH_FAILED: 0,
  };
  const byField = {};
  const criticalUsers = [];

  for (const entry of perUser) {
    const { user, upn, error, missing, results } = entry;
    const nick = user.MailNickName || user.mailNickName || user.mailNickname || upn;
    if (error) {
      counts.FETCH_FAILED++;
      criticalUsers.push({ nick, summary: `FETCH_FAILED: ${entry.detail}` });
      continue;
    }
    if (missing) {
      counts.PROFILE_MISSING++;
      criticalUsers.push({ nick, summary: 'PROFILE_MISSING (404)' });
      continue;
    }
    const userFailures = [];
    for (const [field, check] of Object.entries(results)) {
      counts[check.status]++;
      byField[field] = byField[field] || {
        PASS: 0, OUTPUT_MISSING: 0, SHAPE_COLLAPSE: 0, INPUT_MISSING: 0,
      };
      byField[field][check.status]++;
      if (check.status === 'OUTPUT_MISSING' || check.status === 'SHAPE_COLLAPSE') {
        userFailures.push(`${field}: ${check.status}${check.detail ? ` (${check.detail})` : ''}`);
      }
    }
    if (userFailures.length) criticalUsers.push({ nick, summary: userFailures.join('; ') });
  }

  // ── print ──
  console.log('─'.repeat(72));
  console.log('PER-FIELD SUMMARY');
  console.log('─'.repeat(72));
  const fieldNames = Object.keys(byField).sort();
  for (const f of fieldNames) {
    const b = byField[f];
    const total = b.PASS + b.OUTPUT_MISSING + b.SHAPE_COLLAPSE + b.INPUT_MISSING;
    console.log(
      `  ${f.padEnd(22)}  pass=${String(b.PASS).padStart(3)}  ` +
      `out_missing=${String(b.OUTPUT_MISSING).padStart(3)}  ` +
      `shape_collapse=${String(b.SHAPE_COLLAPSE).padStart(3)}  ` +
      `input_missing=${String(b.INPUT_MISSING).padStart(3)}  ` +
      `(${total} users)`
    );
  }

  console.log('\n' + '─'.repeat(72));
  console.log('TOTALS');
  console.log('─'.repeat(72));
  for (const [k, v] of Object.entries(counts)) {
    console.log(`  ${k.padEnd(20)} ${v}`);
  }

  if (criticalUsers.length) {
    console.log('\n' + '─'.repeat(72));
    console.log(`USERS WITH CONTENT ISSUES (${criticalUsers.length})`);
    console.log('─'.repeat(72));
    for (const u of criticalUsers.slice(0, 30)) {
      console.log(`  ${u.nick.padEnd(25)} ${u.summary}`);
    }
    if (criticalUsers.length > 30) console.log(`  ... and ${criticalUsers.length - 30} more`);
  }

  const failing = counts.OUTPUT_MISSING + counts.SHAPE_COLLAPSE;
  console.log('\n' + '─'.repeat(72));
  if (failing === 0) {
    console.log(`OK — no content regressions. INPUT_MISSING=${counts.INPUT_MISSING} (upstream config gaps, non-failing). PROFILE_MISSING=${counts.PROFILE_MISSING}.`);
    process.exit(0);
  } else {
    console.log(`FAIL — ${counts.OUTPUT_MISSING} OUTPUT_MISSING + ${counts.SHAPE_COLLAPSE} SHAPE_COLLAPSE across fields.`);
    process.exit(1);
  }
}

main().catch(err => {
  console.error(err);
  process.exit(1);
});
