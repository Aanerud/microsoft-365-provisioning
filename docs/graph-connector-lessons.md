# Graph Connector Lessons Learned

Validated through m365people14–m365people23 (Jan–Feb 2026).

## Path A vs Path B — The Most Important Rule

Confirmed by Microsoft engineer Morten (internal team):

**Properties WITH a people data label** → Path A (JsonElement deserialization)
- Safely handles `string`, `stringCollection`, arrays, nested objects
- All 13 official labels go through this path

**Properties WITHOUT a label** → Path B (Blob deserialization)
- Expects `string` values ONLY
- Sending a `stringCollection` (array) without a label throws:
  ```
  Cannot get the value of a token type 'StartArray' as a string. Path: $.skills
  ```
- **Rule**: Custom properties MUST be type `string` — never `stringCollection`

## Connection History

| Connection | Labels | Custom | Result | Root Cause |
|---|---|---|---|---|
| m365people14 | personSkills | — | Working | Baseline |
| m365people15 | personSkills, personNote | VTeam | Working | Added notes + custom prop |
| m365people16 | 7 labels | 8 (incl. stringCollection) | **FAILED** | `responsibilities` as stringCollection without label → Path B crash |
| m365people17 | 4 labels | VTeam | Working | Rolled back to safe config |
| m365people18 | 5 labels | VTeam, BenefitPlan | **FAILED** | Code refactoring broke it (changed CUSTOM_PROPERTIES type, added dead code) |
| m365people19 | 4 labels | VTeam, BenefitPlan | **FAILED** | Same refactored code |
| m365people20 | 4 labels | — | **FAILED** | Same refactored code (even with zero custom props) |
| m365people21 | 4 labels | VTeam | Working | Reverted to exact m365people17 code structure |
| m365people22 | 13 labels | VTeam | Working | All labels enabled, minimal code additions to m365people21 baseline |
| m365people23 | 13 labels | VTeam | Working | Clean baseline (deleted old, fresh deploy) |

## Key Lessons

### 1. Never Refactor Working Connector Code

m365people18–20 failed despite having correct configurations. The ONLY difference from working m365people17/21 was "harmless" refactoring:
- Changed `CUSTOM_PROPERTIES` type from `Array<{name, type}>` to `string[]`
- Removed `getCustomProperties()` method
- Added unused `personWebSite` handler

These changes produced functionally identical item payloads, yet the connector stopped working. Whether this was caused by compiled JS differences or Microsoft backend timing is unknown. **The lesson: don't touch working code structure.**

### 2. Build on Working Baselines

The path from m365people21 (4 labels, working) to m365people22 (13 labels, working) succeeded because we made ONLY additive changes:
- Added entries to `ENABLED_LABELS` sets
- Added new serialization handlers (personWebSite, composites)
- Added composite property declarations to schema
- Did NOT change existing code structure, types, or method signatures

### 3. Custom Properties (Path B) Are Fragile

- VTeam as `string` works reliably (validated across 5+ connections)
- `responsibilities` as `stringCollection` without label crashed the entire connection
- Even `BenefitPlan` as `string` may have contributed to failures (m365people19)
- Safe rule: keep custom properties to a minimum, always `string` type

### 4. Delete Old Connections Before Deploying New Ones

Multiple active connections with overlapping labels can cause priority conflicts in `prioritizedSourceUrls`. The highest-priority source takes precedence — if it hasn't synced yet, profiles show empty data.

### 5. Profile Source Propagation Timing Matters

Our `connection-manager.ts` waits 60 seconds after profile source registration and polls to confirm. Without this, early items get `ProfileSourceRegistrar: Failed to retrieve settings from TSS, statusCode=Unauthorized` and `profileSyncEnabled` never flips to True.

### 6. profileSyncEnabled=False on Initial Export is Normal

The internal debug CSV shows `profileSyncEnabled=False` on the first CAPIv2 export cycle. This does NOT mean failure — data typically appears on profiles after subsequent cycles (6–24 hours). True failure is when it stays False across multiple cycles with TSS errors.

## All 13 People Data Labels

### Group 1: Option B Native (single CSV column)

| Label | Schema Type | CSV Column | Entity | Serialization |
|---|---|---|---|---|
| personSkills | stringCollection | skills | skillProficiency | `{"displayName":"..."}` |
| personNote | string | aboutMe | personAnnotation | `{"detail":{"contentType":"text","content":"..."}}` |
| personCertifications | stringCollection | certifications | personCertification | `{"displayName":"..."}` |
| personProjects | stringCollection | projects | projectParticipation | `{"displayName":"..."}` |
| personAwards | stringCollection | awards | personAward | `{"displayName":"..."}` |
| personAnniversaries | stringCollection | birthday | personAnniversary | `{"type":"birthday","date":"..."}` |
| personWebSite | string | mySite | webSite | `{"webUrl":"..."}` |

### Group 2: Composite (multiple Option A CSV columns)

| Label | Schema Type | CSV Columns | Entity | Serialization |
|---|---|---|---|---|
| personName | string | givenName, surname, displayName | personName | `{"displayName":"...","first":"...","last":"..."}` |
| personCurrentPosition | string | jobTitle, companyName, department | workPosition | `{"detail":{"jobTitle":"...","company":{"displayName":"..."}},"isCurrent":true}` |
| personAddresses | stringCollection | streetAddress, city, state, country, postalCode | itemAddress | `{"type":"business","street":"...","city":"...","state":"...","countryOrRegion":"...","postalCode":"..."}` |
| personEmails | stringCollection | mail, email | itemEmail | `{"address":"...","type":"main"}` |
| personPhones | stringCollection | mobilePhone, businessPhones | itemPhone | `{"number":"...","type":"mobile"/"business"}` |
| personWebAccounts | stringCollection | (none) | webAccount | No CSV data — declared but empty |

### Open Issue: addresses

`personAddresses` label is declared and items are ingested, but `addresses: []` appears on profiles (m365people22, m365people23). The `itemAddress` entity may require a different JSON structure (possibly needs `detail` wrapper or different field names). To investigate.

## Schema Item Format Reference

From Microsoft docs (`docs/MicrosoftDocs/build-connectors-with-people-data.md`):

```json
{
  "id": "person-user-domain-com",
  "acl": [{ "type": "everyone", "value": "everyone", "accessType": "grant" }],
  "properties": {
    "accountInformation": "{\"userPrincipalName\":\"user@domain.com\"}",
    "skills@odata.type": "Collection(String)",
    "skills": [
      "{\"displayName\":\"TypeScript\"}",
      "{\"displayName\":\"Project Management\"}"
    ],
    "aboutMe": "{\"detail\":{\"contentType\":\"text\",\"content\":\"About me...\"}}"
  }
}
```

Key patterns:
- `accountInformation` with `personAccount` label is REQUIRED for user mapping
- `stringCollection` needs `@odata.type: "Collection(String)"` annotation
- Each array item is a JSON-stringified entity object
- `string` properties are single JSON-stringified entity objects
