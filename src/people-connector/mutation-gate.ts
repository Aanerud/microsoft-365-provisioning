import { getPeopleDataMapping } from '../schema/user-property-schema.js';
import { ENABLED_LABELS } from './schema-builder.js';

export type MutationKind = 'REMOVED' | 'CHANGED' | 'TYPE_CHANGED' | 'ADDED';

export interface FieldMutation {
  field: string;
  kind: MutationKind;
  rawValue: any;
  serializedValue: any;
  label?: string;
}

export interface MutationWarning {
  field: string;
  message: string;
  docReference?: string;
}

export interface MutationReport {
  email: string;
  mutations: FieldMutation[];
  warnings: MutationWarning[];
}

export interface GateResult {
  reports: MutationReport[];
  totalMutations: number;
  totalWarnings: number;
  clean: boolean;
}

const UNSUPPORTED_LABELS = new Set([
  'personManager',
  'personAssistants',
  'personColleagues',
  'personAlternateContacts',
  'personEmergencyContacts',
]);

const PEOPLE_DATA_DOC = 'https://learn.microsoft.com/en-us/microsoft-365/copilot/extensibility/build-connectors-with-people-data';

const PRESERVED_ARRAY_ALIASES = new Map<string, string>([
  ['anniversaries', 'birthday'],
  ['notes', 'aboutMe'],
  ['websites', 'mySite'],
]);

function getOptionBFieldNames(): Set<string> {
  const mapping = getPeopleDataMapping();
  const names = new Set<string>();
  for (const [name, label] of mapping.entries()) {
    if (label && ENABLED_LABELS.has(label)) {
      names.add(name);
    }
  }
  for (const alias of PRESERVED_ARRAY_ALIASES.keys()) {
    names.add(alias);
  }
  return names;
}

function deepNormalizeKeys(obj: any): any {
  if (obj === null || obj === undefined) return obj;
  if (typeof obj !== 'object') return obj;
  if (Array.isArray(obj)) return obj.map(deepNormalizeKeys);
  const result: any = {};
  for (const [key, value] of Object.entries(obj)) {
    const camelKey = key.charAt(0).toLowerCase() + key.slice(1);
    result[camelKey] = deepNormalizeKeys(value);
  }
  return result;
}

function valuesEqual(a: any, b: any): boolean {
  if (a === b) return true;
  if (a === undefined && b === undefined) return true;
  if (a === null && b === null) return true;
  if (typeof a !== typeof b) return false;
  return JSON.stringify(deepNormalizeKeys(a)) === JSON.stringify(deepNormalizeKeys(b));
}

export function detectMutations(
  rawRecord: any,
  normalizedRecord: any,
): FieldMutation[] {
  const mutations: FieldMutation[] = [];
  const optionBFields = getOptionBFieldNames();
  const mapping = getPeopleDataMapping();

  for (const field of optionBFields) {
    const rawVal = rawRecord[field] ?? rawRecord[field.charAt(0).toUpperCase() + field.slice(1)];
    const normalizedVal = normalizedRecord[field];

    if (rawVal === undefined || rawVal === null) continue;

    if (normalizedVal === undefined) {
      mutations.push({
        field,
        kind: 'REMOVED',
        rawValue: rawVal,
        serializedValue: undefined,
        label: mapping.get(field) ?? undefined,
      });
    } else if (!valuesEqual(rawVal, normalizedVal)) {
      const rawType = Array.isArray(rawVal) ? 'array' : typeof rawVal;
      const normType = Array.isArray(normalizedVal) ? 'array' : typeof normalizedVal;

      mutations.push({
        field,
        kind: rawType !== normType ? 'TYPE_CHANGED' : 'CHANGED',
        rawValue: rawVal,
        serializedValue: normalizedVal,
        label: mapping.get(field) ?? undefined,
      });
    }
  }

  return mutations;
}

export function checkFormatWarnings(record: any): MutationWarning[] {
  const warnings: MutationWarning[] = [];

  const VALID_ANNIVERSARY_TYPES = new Set(['birthday', 'wedding', 'unknownFutureValue']);
  const anniversaries = record.anniversaries || record.birthday;
  if (Array.isArray(anniversaries)) {
    for (const ann of anniversaries) {
      if (ann?.type && !VALID_ANNIVERSARY_TYPES.has(ann.type)) {
        warnings.push({
          field: 'anniversaries.type',
          message: `Type "${ann.type}" is not a documented anniversaryType. Valid values: birthday, wedding. May be silently ignored by PCP.`,
          docReference: 'https://learn.microsoft.com/en-us/graph/api/resources/personanniversary?view=graph-rest-beta',
        });
      }
    }
  }

  if (Array.isArray(record.addresses) && record.addresses.length > 3) {
    warnings.push({
      field: 'addresses',
      message: `${record.addresses.length} addresses found, MS docs limit is max 3 (one each of Home, Work, Other). Extras will be silently ignored by Graph.`,
      docReference: PEOPLE_DATA_DOC,
    });
  }

  if (Array.isArray(record.emails) && record.emails.length > 3) {
    warnings.push({
      field: 'emails',
      message: `${record.emails.length} emails found, MS docs limit is max 3. Extras will be silently ignored by Graph.`,
      docReference: PEOPLE_DATA_DOC,
    });
  }

  const mapping = getPeopleDataMapping();
  for (const [, label] of mapping.entries()) {
    if (label && UNSUPPORTED_LABELS.has(label)) {
      warnings.push({
        field: label,
        message: `Label "${label}" is not currently supported by Microsoft. Data will be ignored.`,
        docReference: PEOPLE_DATA_DOC,
      });
      break;
    }
  }

  return warnings;
}

export function runMutationGate(
  rawRecords: any[],
  normalizedRecords: any[],
): GateResult {
  const reports: MutationReport[] = [];
  let totalMutations = 0;
  let totalWarnings = 0;

  for (let i = 0; i < normalizedRecords.length; i++) {
    const raw = rawRecords[i] || {};
    const normalized = normalizedRecords[i] || {};
    const email = normalized.email || raw.email || raw.MailNickName || `record-${i}`;

    const mutations = detectMutations(raw, normalized);
    const warnings = checkFormatWarnings(normalized);

    if (mutations.length > 0 || warnings.length > 0) {
      reports.push({ email, mutations, warnings });
      totalMutations += mutations.length;
      totalWarnings += warnings.length;
    }
  }

  return {
    reports,
    totalMutations,
    totalWarnings,
    clean: totalMutations === 0 && totalWarnings === 0,
  };
}

export function formatMutationReport(result: GateResult): string {
  if (result.clean) {
    return '✓ No mutations or warnings detected. Config data passes through cleanly.';
  }

  const lines: string[] = [];
  lines.push(`\n${'='.repeat(70)}`);
  lines.push(`MUTATION GATE REPORT: ${result.totalMutations} mutations, ${result.totalWarnings} warnings`);
  lines.push('='.repeat(70));

  for (const report of result.reports) {
    lines.push(`\n── ${report.email} ──`);

    for (const m of report.mutations) {
      const label = m.label ? ` (${m.label})` : '';
      lines.push(`  [${m.kind}] ${m.field}${label}`);
      if (m.kind === 'REMOVED') {
        lines.push(`    raw: ${JSON.stringify(m.rawValue)?.slice(0, 120)}`);
        lines.push(`    out: (deleted)`);
      } else {
        lines.push(`    raw: ${JSON.stringify(m.rawValue)?.slice(0, 120)}`);
        lines.push(`    out: ${JSON.stringify(m.serializedValue)?.slice(0, 120)}`);
      }
    }

    for (const w of report.warnings) {
      const ref = w.docReference ? ` [${w.docReference}]` : '';
      lines.push(`  [WARN] ${w.field}: ${w.message}${ref}`);
    }
  }

  lines.push(`\n${'='.repeat(70)}`);
  return lines.join('\n');
}
