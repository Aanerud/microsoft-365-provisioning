import fs from 'fs/promises';
import path from 'path';
import { Client } from '@microsoft/microsoft-graph-client';
import { BrowserAuthServer } from './auth/browser-auth-server.js';
import { GraphClient } from './graph-client.js';

export interface OidCacheFile {
  generatedAt: string;
  tenantId: string;
  source: 'graph-beta';
  userCount: number;
  users: Record<string, string>;
}

export interface OidCacheSummary {
  total: number;
  hits: number;
  misses: number;
  existing: number;
}

const DEFAULT_PAGE_SIZE = 999;

function normalizeUpn(upn: string): string {
  return upn.trim().toLowerCase();
}

export function getOidCachePath(csvPath: string): string {
  const directory = path.dirname(csvPath);
  const baseName = path.basename(csvPath, path.extname(csvPath));
  return path.join(directory, `${baseName}_oid_cache.json`);
}

export async function loadOidCache(cachePath: string): Promise<OidCacheFile | null> {
  try {
    const content = await fs.readFile(cachePath, 'utf-8');
    return JSON.parse(content) as OidCacheFile;
  } catch (error: any) {
    if (error?.code === 'ENOENT') {
      return null;
    }
    throw error;
  }
}

async function fetchAllUsers(betaClient: Client): Promise<Array<{ id: string; userPrincipalName?: string }>> {
  const users: Array<{ id: string; userPrincipalName?: string }> = [];
  const selectFields = 'id,userPrincipalName';

  let response = await betaClient
    .api('/users')
    .select(selectFields)
    .top(DEFAULT_PAGE_SIZE)
    .get();

  users.push(...(response.value || []));
  let nextLink: string | undefined = response['@odata.nextLink'];

  while (nextLink) {
    response = await betaClient.api(nextLink).get();
    users.push(...(response.value || []));
    nextLink = response['@odata.nextLink'];
  }

  return users;
}

export async function buildOidCacheWithClient(options: {
  csvPath: string;
  tenantId: string;
  graphClient: GraphClient;
  cachePath?: string;
}): Promise<OidCacheFile> {
  const cachePath = options.cachePath ?? getOidCachePath(options.csvPath);
  const { betaClient } = options.graphClient.getClients();

  console.log(`\n🔍 Building OID cache from Graph beta...`);
  const users = await fetchAllUsers(betaClient);
  console.log(`  Fetched ${users.length} users`);

  const mapping: Record<string, string> = {};
  let missingUpn = 0;
  let duplicateUpn = 0;

  for (const user of users) {
    const upn = user.userPrincipalName;
    if (!upn || !user.id) {
      missingUpn++;
      continue;
    }
    const normalizedUpn = normalizeUpn(upn);
    if (mapping[normalizedUpn]) {
      duplicateUpn++;
      continue;
    }
    mapping[normalizedUpn] = user.id;
  }

  if (missingUpn > 0) {
    console.log(`  Skipped ${missingUpn} users without UPN`);
  }
  if (duplicateUpn > 0) {
    console.log(`  Skipped ${duplicateUpn} duplicate UPNs`);
  }

  const cache: OidCacheFile = {
    generatedAt: new Date().toISOString(),
    tenantId: options.tenantId,
    source: 'graph-beta',
    userCount: Object.keys(mapping).length,
    users: mapping,
  };

  await fs.mkdir(path.dirname(cachePath), { recursive: true });
  await fs.writeFile(cachePath, JSON.stringify(cache, null, 2), 'utf-8');
  console.log(`✓ OID cache written: ${cachePath}`);

  return cache;
}

export async function ensureOidCacheWithClient(options: {
  csvPath: string;
  tenantId: string;
  graphClient: GraphClient;
  force?: boolean;
}): Promise<{ cache: OidCacheFile; cachePath: string; rebuilt: boolean }> {
  const cachePath = getOidCachePath(options.csvPath);
  if (!options.force) {
    const existing = await loadOidCache(cachePath);
    if (existing) {
      return { cache: existing, cachePath, rebuilt: false };
    }
  }

  const cache = await buildOidCacheWithClient({
    csvPath: options.csvPath,
    tenantId: options.tenantId,
    graphClient: options.graphClient,
    cachePath,
  });

  return { cache, cachePath, rebuilt: true };
}

export async function ensureOidCacheWithAuth(options: {
  csvPath: string;
  tenantId: string;
  clientId: string;
  authPort?: number;
  force?: boolean;
  forceRefresh?: boolean;
}): Promise<{ cache: OidCacheFile; cachePath: string; rebuilt: boolean }> {
  const cachePath = getOidCachePath(options.csvPath);
  if (!options.force) {
    const existing = await loadOidCache(cachePath);
    if (existing) {
      return { cache: existing, cachePath, rebuilt: false };
    }
  }

  console.log('\n🔐 Authentication required to build OID cache (delegated)...');
  const authServer = new BrowserAuthServer({
    tenantId: options.tenantId,
    clientId: options.clientId,
    port: options.authPort || 5544,
    forceRefresh: options.forceRefresh || false,
  });

  const authResult = await authServer.authenticate();
  const graphClient = new GraphClient({ accessToken: authResult.accessToken });

  const cache = await buildOidCacheWithClient({
    csvPath: options.csvPath,
    tenantId: options.tenantId,
    graphClient,
    cachePath,
  });

  return { cache, cachePath, rebuilt: true };
}

export function applyOidCacheToRows(
  rows: Array<Record<string, any>>,
  cache: OidCacheFile
): OidCacheSummary {
  const summary: OidCacheSummary = {
    total: rows.length,
    hits: 0,
    misses: 0,
    existing: 0,
  };

  for (const row of rows) {
    if (row.externalDirectoryObjectId) {
      summary.existing++;
      continue;
    }

    const email = row.email;
    if (!email) {
      summary.misses++;
      continue;
    }

    const oid = cache.users[normalizeUpn(email)];
    if (oid) {
      row.externalDirectoryObjectId = oid;
      summary.hits++;
    } else {
      summary.misses++;
    }
  }

  return summary;
}
