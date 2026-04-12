/**
 * Resolves license display names (from users.config.json) to SKU IDs
 * by querying the tenant's subscribed SKUs.
 */
import { GraphClient } from './graph-client.js';

// Static mapping: common display names → skuPartNumber patterns
const DISPLAY_NAME_PATTERNS: Array<{ pattern: RegExp; skuPattern: RegExp }> = [
  { pattern: /Office 365 E5.*no Teams/i, skuPattern: /Office_365_E5.*no.Teams|ENTERPRISEPREMIUM_NOPSTNCONF/i },
  { pattern: /Office 365 E5/i, skuPattern: /^(?!.*no.Teams).*(Office_365_E5|ENTERPRISEPREMIUM)/i },
  { pattern: /Office 365 E3/i, skuPattern: /ENTERPRISEPACK/i },
  { pattern: /Microsoft 365 E5/i, skuPattern: /SPE_E5/i },
  { pattern: /Microsoft 365 E3/i, skuPattern: /SPE_E3/i },
  { pattern: /Microsoft Teams Enterprise/i, skuPattern: /^Microsoft_Teams_Enterprise|^TEAMS_ENTERPRISE/i },
  { pattern: /Microsoft 365 Copilot/i, skuPattern: /Microsoft_365_Copilot|Copilot/i },
  { pattern: /Power BI Pro/i, skuPattern: /POWER_BI_PRO/i },
  { pattern: /Visio/i, skuPattern: /VISIO/i },
  { pattern: /Project/i, skuPattern: /PROJECT/i },
];

export interface LicenseResolution {
  resolved: Map<string, string>;    // displayName → skuId
  unresolved: string[];             // display names that couldn't be matched
}

/**
 * Resolve an array of license display names to their SKU IDs.
 * Queries the tenant once, then matches all names.
 */
export async function resolveLicenses(
  graphClient: GraphClient,
  displayNames: string[]
): Promise<LicenseResolution> {
  const tenantSkus = await graphClient.listLicenses();
  const resolved = new Map<string, string>();
  const unresolved: string[] = [];

  for (const name of displayNames) {
    if (resolved.has(name)) continue;

    // Try pattern matching
    let matched = false;
    for (const { pattern, skuPattern } of DISPLAY_NAME_PATTERNS) {
      if (pattern.test(name)) {
        const sku = tenantSkus.find(s => skuPattern.test(s.skuPartNumber));
        if (sku) {
          resolved.set(name, sku.skuId);
          matched = true;
          break;
        }
      }
    }

    // Fallback: substring match on skuPartNumber
    if (!matched) {
      const words = name.replace(/[^a-zA-Z0-9]/g, ' ').split(/\s+/).filter(w => w.length > 3);
      const sku = tenantSkus.find(s =>
        words.some(w => s.skuPartNumber.toLowerCase().includes(w.toLowerCase()))
      );
      if (sku) {
        resolved.set(name, sku.skuId);
        matched = true;
      }
    }

    if (!matched) {
      unresolved.push(name);
    }
  }

  return { resolved, unresolved };
}

/**
 * Resolve per-user license arrays from all records.
 * Returns a cached resolver function.
 */
export async function buildLicenseResolver(
  graphClient: GraphClient,
  records: any[]
): Promise<(userLicenses: string[]) => string[]> {
  // Collect all unique display names across all records
  const allNames = new Set<string>();
  for (const r of records) {
    if (Array.isArray(r.licenses)) {
      for (const name of r.licenses) {
        if (typeof name === 'string') allNames.add(name);
      }
    }
  }

  if (allNames.size === 0) {
    return () => [];
  }

  console.log(`Resolving ${allNames.size} license type(s) from JSON...`);
  const { resolved, unresolved } = await resolveLicenses(graphClient, [...allNames]);

  for (const [name, skuId] of resolved) {
    console.log(`  ✓ "${name}" → ${skuId}`);
  }
  for (const name of unresolved) {
    console.warn(`  ⚠ "${name}" → not found in tenant`);
  }

  return (userLicenses: string[]) =>
    userLicenses
      .map(name => resolved.get(name))
      .filter((id): id is string => !!id);
}
