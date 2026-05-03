import { describe, it, expect } from 'vitest';
import { detectMutations, checkFormatWarnings, runMutationGate, formatMutationReport } from '../mutation-gate.js';

describe('detectMutations', () => {
  it('returns empty when raw and normalized match', () => {
    const raw = { skills: [{ displayName: 'TypeScript' }] };
    const normalized = { skills: [{ displayName: 'TypeScript' }] };
    const mutations = detectMutations(raw, normalized);
    expect(mutations).toHaveLength(0);
  });

  it('detects REMOVED when a field is deleted by normalization', () => {
    const raw = { birthday: '1990-05-15' };
    const normalized = {};
    const mutations = detectMutations(raw, normalized);
    const removed = mutations.find(m => m.field === 'birthday');
    expect(removed).toBeDefined();
    expect(removed!.kind).toBe('REMOVED');
  });

  it('detects CHANGED when field value differs', () => {
    const raw = { skills: [{ displayName: 'TS', proficiency: 'expert' }] };
    const normalized = { skills: [{ displayName: 'TS' }] };
    const mutations = detectMutations(raw, normalized);
    const changed = mutations.find(m => m.field === 'skills');
    expect(changed).toBeDefined();
    expect(changed!.kind).toBe('CHANGED');
  });

  it('detects TYPE_CHANGED when array becomes string', () => {
    const raw = { aboutMe: ['Note 1', 'Note 2'] };
    const normalized = { aboutMe: 'Note 1' };
    const mutations = detectMutations(raw, normalized);
    const typeChanged = mutations.find(m => m.field === 'aboutMe');
    expect(typeChanged).toBeDefined();
    expect(typeChanged!.kind).toBe('TYPE_CHANGED');
  });

  it('ignores fields not in Option B schema', () => {
    const raw = { jobTitle: 'PM', randomField: 'hello' };
    const normalized = { jobTitle: 'Product Manager', randomField: 'world' };
    const mutations = detectMutations(raw, normalized);
    expect(mutations).toHaveLength(0);
  });

  it('handles PascalCase raw keys by looking up capitalized variant', () => {
    const raw = { Skills: [{ displayName: 'Go' }] };
    const normalized = { skills: [{ displayName: 'Go' }] };
    const mutations = detectMutations(raw, normalized);
    expect(mutations).toHaveLength(0);
  });
});

describe('checkFormatWarnings', () => {
  it('warns when addresses exceed max 3', () => {
    const record = {
      addresses: [
        { type: 'Home' }, { type: 'Work' }, { type: 'Other' }, { type: 'Extra' },
      ],
    };
    const warnings = checkFormatWarnings(record);
    const addrWarn = warnings.find(w => w.field === 'addresses');
    expect(addrWarn).toBeDefined();
    expect(addrWarn!.message).toContain('max 3');
  });

  it('warns when emails exceed max 3', () => {
    const record = {
      emails: [{ address: 'a@b.com' }, { address: 'c@d.com' }, { address: 'e@f.com' }, { address: 'g@h.com' }],
    };
    const warnings = checkFormatWarnings(record);
    expect(warnings.find(w => w.field === 'emails')).toBeDefined();
  });

  it('does not warn on YYYY-MM-DD in startMonthYear (Date type accepts full dates)', () => {
    const record = {
      positions: [{ detail: { startMonthYear: '2020-03-15', jobTitle: 'PM' } }],
    };
    const warnings = checkFormatWarnings(record);
    expect(warnings.find(w => w.field === 'positions.detail.startMonthYear')).toBeUndefined();
  });

  it('does not warn on fieldsOfStudy as string (Graph API type is String)', () => {
    const record = {
      educationalActivities: [{ program: { fieldsOfStudy: 'Business' } }],
    };
    const warnings = checkFormatWarnings(record);
    expect(warnings.find(w => w.field === 'educationalActivities.program.fieldsOfStudy')).toBeUndefined();
  });

  it('returns empty for clean record', () => {
    const record = {
      skills: [{ displayName: 'Go' }],
      addresses: [{ type: 'Work', city: 'Oslo' }],
    };
    const warnings = checkFormatWarnings(record);
    expect(warnings).toHaveLength(0);
  });
});

describe('runMutationGate', () => {
  it('returns clean when no mutations or warnings', () => {
    const raw = [{ skills: [{ displayName: 'TS' }], email: 'a@b.com' }];
    const normalized = [{ skills: [{ displayName: 'TS' }], email: 'a@b.com' }];
    const result = runMutationGate(raw, normalized);
    expect(result.clean).toBe(true);
    expect(result.totalMutations).toBe(0);
    expect(result.totalWarnings).toBe(0);
  });

  it('reports mutations across multiple records', () => {
    const raw = [
      { skills: [{ displayName: 'TS' }], email: 'a@b.com' },
      { birthday: '1990-05-15', email: 'c@d.com' },
    ];
    const normalized = [
      { skills: [{ displayName: 'TS' }], email: 'a@b.com' },
      { email: 'c@d.com' },
    ];
    const result = runMutationGate(raw, normalized);
    expect(result.clean).toBe(false);
    expect(result.totalMutations).toBe(1);
  });
});

describe('formatMutationReport', () => {
  it('outputs clean message when no issues', () => {
    const result = { reports: [], totalMutations: 0, totalWarnings: 0, clean: true };
    const output = formatMutationReport(result);
    expect(output).toContain('No mutations');
  });

  it('outputs mutation details when issues found', () => {
    const result = {
      reports: [{
        email: 'test@example.com',
        mutations: [{
          field: 'anniversaries',
          kind: 'REMOVED' as const,
          rawValue: [{ type: 'work' }],
          serializedValue: undefined,
          label: 'personAnniversaries',
        }],
        warnings: [],
      }],
      totalMutations: 1,
      totalWarnings: 0,
      clean: false,
    };
    const output = formatMutationReport(result);
    expect(output).toContain('REMOVED');
    expect(output).toContain('anniversaries');
    expect(output).toContain('test@example.com');
  });
});
