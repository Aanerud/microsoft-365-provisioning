# Graph Connector Lessons Learned

Validated through m365people14–m365people23 (Jan–Feb 2026).

## Path A vs Path B — The Most Important Rule

Confirmed by Microsoft Graph Connector engineering team:

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
| m365people24 | 13 labels | 7 dynamic custom | **Working** | Dynamic custom properties from CSV (VTeam, BenefitPlan, CostCenter, BuildingAccess, ProjectCode, WritingStyle, Specialization) |

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

### 3. Custom Properties (Path B) — Dynamic from CSV

Custom properties are now **auto-detected from CSV columns**. Any column not in the standard schema (`user-property-schema.ts`) and not an internal column (`name`, `email`, `role`, `ManagerEmail`) becomes a custom connector property.

**How it works**:
- `getCustomProperties(csvColumns)` in `user-property-schema.ts` detects non-standard columns
- `schema-builder.ts` accepts optional `csvColumns` parameter; falls back to hardcoded list when not provided
- All custom properties are registered as `string` type only (Path B safe)

**Rules**:
- Custom properties MUST be `string` type — never `stringCollection` (Path B crashes on arrays)
- `responsibilities` has no label and is array type — handled by Profile API only, NOT the connector
- VTeam + 6 more custom properties validated with m365people24

**Schema limitation**: Connector schemas cannot be updated once registered. If you add new custom columns to your CSV, you must delete the old connector and create a new one:
```bash
npm run enrich:delete-connector -- --connection-id m365people23
npm run option-b:setup -- --csv config/updated.csv --connection-id m365people24
```

### 4. Delete Old Connections Before Deploying New Ones

Multiple active connections with overlapping labels can cause priority conflicts in `prioritizedSourceUrls`. The highest-priority source takes precedence — if it hasn't synced yet, profiles show empty data.

### 5. Profile Source Propagation Timing Matters

Our `connection-manager.ts` waits 60 seconds after profile source registration and polls to confirm. Without this, early items get `ProfileSourceRegistrar: Failed to retrieve settings from TSS, statusCode=Unauthorized` and `profileSyncEnabled` never flips to True.

### 6. profileSyncEnabled=False on Initial Export is Normal

The internal debug CSV shows `profileSyncEnabled=False` on the first CAPIv2 export cycle. This does NOT mean failure — data typically appears on profiles after subsequent cycles (6–24 hours). True failure is when it stays False across multiple cycles with TSS errors.

### 7. enabledContentExperiences Must Be Set for Copilot/Search

The connection must have `enabledContentExperiences: ['search']` for data to appear in Copilot Chat and Search Results. Without this, the admin portal shows "Data from this connection will not appear in Copilot Chat or Search Results." Added to `connection-manager.ts` in the connection creation payload. Can also be PATCHed onto existing connections via `enableSearchExperience()`.

### 8. Semantic Labels (title, url, etc.) — Not Needed for People Connectors

The admin portal recommends 4 semantic labels: `title`, `lastModifiedBy`, `lastModifiedDateTime`, `url`. These are general connector labels for document-style content (articles, tickets, files). For people data connectors:
- `title` → Person's name is already in `personName` label
- `url` → No direct URL to person's profile in our CSV data source
- `lastModifiedBy` / `lastModifiedDateTime` → Would just be app name + current timestamp

These labels become relevant if the data source changes from CSV to JSON or an API with richer metadata. For CSV-based people enrichment, the 13 people data labels are sufficient.

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
