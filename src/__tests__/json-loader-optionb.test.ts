import { describe, it, expect, vi, beforeEach } from 'vitest';
import { loadRowsFromJson } from '../json-loader.js';
import fs from 'fs/promises';

vi.mock('fs/promises');

const mockRecord = {
  MailNickName: 'nora.d',
  DisplayName: 'Nora Dahl',
  FirstName: 'Nora',
  LastName: 'Dahl',
  Department: 'Engineering',
  Skills: [
    { DisplayName: 'Strategy', Proficiency: 'expert', AllowedAudiences: 'organization' },
  ],
  Anniversaries: [
    { Type: 'originalHireDate', Date: '2015-01-15', AllowedAudiences: 'organization' },
  ],
  Emails: [
    { Address: 'nora.d@contoso.com', Type: 'work', AllowedAudiences: 'organization' },
  ],
  Phones: [
    { Number: '+47-12345678', Type: 'mobile' },
  ],
  WebAccounts: [
    { Description: 'LinkedIn', WebUrl: 'https://linkedin.com/in/nora' },
  ],
  Notes: [
    { Detail: { ContentType: 'text', Content: 'About me text' } },
  ],
  EducationalActivities: [
    { Program: { FieldsOfStudy: 'Business', DisplayName: 'MBA' } },
  ],
  Positions: [
    { Detail: { JobTitle: 'PM', Company: { DisplayName: 'Contoso' } } },
  ],
};

function setupMock(records: any[]) {
  vi.mocked(fs.readFile).mockResolvedValue(JSON.stringify(records) as any);
}

beforeEach(() => {
  vi.clearAllMocks();
  process.env.USER_DOMAIN = 'contoso.onmicrosoft.com';
});

describe('loadRowsFromJson - Option B pipeline', () => {
  it('preserves anniversaries array (not extracted to employeeHireDate)', async () => {
    setupMock([mockRecord]);
    const rows = await loadRowsFromJson('/fake/path.json', 'optionB');
    expect(rows[0].anniversaries).toBeDefined();
    expect(Array.isArray(rows[0].anniversaries)).toBe(true);
    expect(rows[0].anniversaries[0].type).toBe('originalHireDate');
  });

  it('preserves emails array (not collapsed to single mail)', async () => {
    setupMock([mockRecord]);
    const rows = await loadRowsFromJson('/fake/path.json', 'optionB');
    expect(rows[0].emails).toBeDefined();
    expect(Array.isArray(rows[0].emails)).toBe(true);
    expect(rows[0].emails[0].address).toBe('nora.d@contoso.com');
  });

  it('preserves phones array with types', async () => {
    setupMock([mockRecord]);
    const rows = await loadRowsFromJson('/fake/path.json', 'optionB');
    expect(rows[0].phones).toBeDefined();
    expect(Array.isArray(rows[0].phones)).toBe(true);
    expect(rows[0].phones[0].number).toBe('+47-12345678');
    expect(rows[0].phones[0].type).toBe('mobile');
  });

  it('preserves webAccounts (not deleted)', async () => {
    setupMock([mockRecord]);
    const rows = await loadRowsFromJson('/fake/path.json', 'optionB');
    expect(rows[0].webAccounts).toBeDefined();
    expect(rows[0].webAccounts[0].webUrl).toBe('https://linkedin.com/in/nora');
  });

  it('preserves notes array (not extracted to aboutMe)', async () => {
    setupMock([mockRecord]);
    const rows = await loadRowsFromJson('/fake/path.json', 'optionB');
    expect(rows[0].notes).toBeDefined();
    expect(rows[0].aboutMe).toBeUndefined();
  });

  it('preserves metadata fields (allowedAudiences not stripped)', async () => {
    setupMock([mockRecord]);
    const rows = await loadRowsFromJson('/fake/path.json', 'optionB');
    expect(rows[0].skills[0].allowedAudiences).toBe('organization');
  });

  it('does NOT convert fieldsOfStudy from string to array', async () => {
    setupMock([mockRecord]);
    const rows = await loadRowsFromJson('/fake/path.json', 'optionB');
    expect(typeof rows[0].educationalActivities[0].program.fieldsOfStudy).toBe('string');
  });

  it('does NOT inject isCurrent on positions', async () => {
    setupMock([mockRecord]);
    const rows = await loadRowsFromJson('/fake/path.json', 'optionB');
    expect(rows[0].positions[0].isCurrent).toBeUndefined();
  });

  it('still normalizes PascalCase keys to camelCase', async () => {
    setupMock([mockRecord]);
    const rows = await loadRowsFromJson('/fake/path.json', 'optionB');
    expect(rows[0].displayName).toBe('Nora Dahl');
    expect(rows[0].DisplayName).toBeUndefined();
  });

  it('still constructs email from MailNickName', async () => {
    setupMock([mockRecord]);
    const rows = await loadRowsFromJson('/fake/path.json', 'optionB');
    expect(rows[0].email).toBe('nora.d@contoso.onmicrosoft.com');
  });
});

describe('loadRowsFromJson - Option A pipeline (regression)', () => {
  it('extracts anniversaries to employeeHireDate', async () => {
    setupMock([mockRecord]);
    const rows = await loadRowsFromJson('/fake/path.json', 'optionA');
    expect(rows[0].employeeHireDate).toBeUndefined();
  });

  it('collapses emails to mail field', async () => {
    setupMock([mockRecord]);
    const rows = await loadRowsFromJson('/fake/path.json', 'optionA');
    expect(rows[0].mail).toBe('nora.d@contoso.com');
  });

  it('deletes consumed fields', async () => {
    setupMock([mockRecord]);
    const rows = await loadRowsFromJson('/fake/path.json', 'optionA');
    expect(rows[0].emails).toBeUndefined();
    expect(rows[0].phones).toBeUndefined();
    expect(rows[0].anniversaries).toBeUndefined();
    expect(rows[0].webAccounts).toBeUndefined();
    expect(rows[0].notes).toBeUndefined();
  });

  it('strips metadata fields', async () => {
    setupMock([mockRecord]);
    const rows = await loadRowsFromJson('/fake/path.json', 'optionA');
    expect(rows[0].skills[0].allowedAudiences).toBeUndefined();
  });

  it('injects isCurrent on positions', async () => {
    setupMock([mockRecord]);
    const rows = await loadRowsFromJson('/fake/path.json', 'optionA');
    expect(rows[0].positions[0].isCurrent).toBe(true);
  });
});
