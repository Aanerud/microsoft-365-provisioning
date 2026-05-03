#!/usr/bin/env node
/**
 * Dump the exact item payload that would be PUT to the Graph Connector for
 * a single user. Replicates the full enrich-connector.ts pre-ingest pipeline:
 *   1. loadRowsFromJson → normalize + renderProducts
 *   2. applyOidCacheToRows → inject externalDirectoryObjectId
 *   3. inject userId into relatedPerson objects (manager, colleagues, sponsors)
 *   4. PeopleItemIngester.createExternalItem → final payload
 *
 * Use this to compare what we ingest vs what PAPI/PGS stores for any user.
 *
 * Usage:
 *   node tools/debug/dump-ingest-payload.mjs --user fatima.z
 *   node tools/debug/dump-ingest-payload.mjs --user fatima.z --config config/users.connector.config.json
 */

import { readFileSync } from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';
import dotenv from 'dotenv';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const ROOT = path.resolve(__dirname, '../..');
dotenv.config({ path: path.join(ROOT, '.env') });

const DEFAULT_CONFIG = path.join(ROOT, 'config/users.connector.config.json');
const OID_CACHE = path.join(ROOT, 'config/users.connector.config_oid_cache.json');

function parseArgs() {
  const args = process.argv.slice(2);
  const out = { user: null, config: DEFAULT_CONFIG };
  for (let i = 0; i < args.length; i++) {
    if (args[i] === '--user') out.user = args[++i];
    else if (args[i] === '--config') out.config = args[++i];
  }
  if (!out.user) {
    console.error('Usage: node tools/debug/dump-ingest-payload.mjs --user <mailNickname> [--config <path>]');
    process.exit(2);
  }
  return out;
}

async function main() {
  const args = parseArgs();

  const { loadRowsFromJson } = await import('../../dist/json-loader.js');
  const { PeopleItemIngester } = await import('../../dist/people-connector/item-ingester.js');

  // 1. Load + normalize + renderProducts
  const rows = await loadRowsFromJson(args.config, 'optionB');
  const row = rows.find(r => {
    const nick = r.email?.split('@')[0] || r.mailNickName;
    return nick === args.user;
  });
  if (!row) { console.error(`User "${args.user}" not in ${args.config}`); process.exit(1); }

  // 2. Apply OID cache — set externalDirectoryObjectId
  const cache = JSON.parse(readFileSync(OID_CACHE, 'utf-8'));
  const cacheUsers = cache.users || cache;
  const upnLower = row.email.toLowerCase();
  if (cacheUsers[upnLower]) row.externalDirectoryObjectId = cacheUsers[upnLower];

  // 3. Inject userId into relatedPerson objects
  const injectOid = (person) => {
    if (!person || typeof person !== 'object' || person.userId) return;
    const upn = person.userPrincipalName;
    if (!upn) return;
    const oid = cacheUsers[upn.toLowerCase()];
    if (oid) person.userId = oid;
  };
  if (Array.isArray(row.positions)) {
    for (const pos of row.positions) {
      if (pos.manager) injectOid(pos.manager);
      if (Array.isArray(pos.colleagues)) pos.colleagues.forEach(injectOid);
      if (Array.isArray(pos.sponsors)) pos.sponsors.forEach(injectOid);
    }
  }
  if (Array.isArray(row.projects)) {
    for (const proj of row.projects) {
      if (Array.isArray(proj.colleagues)) proj.colleagues.forEach(injectOid);
      if (Array.isArray(proj.sponsors)) proj.sponsors.forEach(injectOid);
    }
  }

  // 4. Build payload via PeopleItemIngester
  const allKeys = new Set();
  for (const r of rows) for (const k of Object.keys(r)) allKeys.add(k);
  const csvColumns = [...allKeys];

  // stub logger + betaClient — createExternalItem doesn't call them
  const stub = { info: () => {}, warn: () => {}, error: () => {} };
  const ingester = new PeopleItemIngester(null, 'dummy-conn', stub, csvColumns);
  const item = ingester.createExternalItem(row);

  // Pretty-print — parse each JSON-stringified property value back for readability
  const pretty = { ...item, properties: {} };
  for (const [k, v] of Object.entries(item.properties || {})) {
    if (typeof v === 'string' && (v.startsWith('{') || v.startsWith('['))) {
      try { pretty.properties[k] = JSON.parse(v); continue; } catch {}
    }
    if (Array.isArray(v)) {
      pretty.properties[k] = v.map(x => {
        if (typeof x === 'string' && (x.startsWith('{') || x.startsWith('['))) {
          try { return JSON.parse(x); } catch { return x; }
        }
        return x;
      });
      continue;
    }
    pretty.properties[k] = v;
  }

  console.log(JSON.stringify(pretty, null, 2));
}

main().catch(err => { console.error(err); process.exit(1); });
