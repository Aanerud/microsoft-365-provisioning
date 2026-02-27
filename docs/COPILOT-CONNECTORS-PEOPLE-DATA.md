# Deep Understanding: Microsoft 365 Copilot Connectors for People Data

## End-to-End Flow Overview

```
┌─────────────────────────────────────────────────────────────────────────────┐
│ 1. CREATE CONNECTION                                                        │
│    POST /external/connections                                               │
│    { id, name, description, contentCategory: 'people' }                     │
│                                                                             │
│    ⚠️ CRITICAL: contentCategory MUST be 'people' for people data connectors │
└─────────────────────────────────────────────────────────────────────────────┘
                                    │
                                    ▼
┌─────────────────────────────────────────────────────────────────────────────┐
│ 2. REGISTER SCHEMA                                                          │
│    POST /external/connections/{id}/schema                                   │
│    { baseType, properties: [...] }                                          │
│                                                                             │
│    Properties define what data fields exist and their types/labels          │
└─────────────────────────────────────────────────────────────────────────────┘
                                    │
                                    ▼
┌─────────────────────────────────────────────────────────────────────────────┐
│ 3. REGISTER AS PROFILE SOURCE (Beta API)                                    │
│    POST /admin/people/profileSources                                        │
│    { sourceId: connectionId, kind: 'Connector', webUrl: '...' }             │
│                                                                             │
│    ⚠️ CRITICAL: kind: 'Connector' is REQUIRED per Microsoft docs            │
│    Links connector to Microsoft 365 People experiences                      │
└─────────────────────────────────────────────────────────────────────────────┘
                                    │
                                    ▼
┌─────────────────────────────────────────────────────────────────────────────┐
│ 4. SET PRIORITY (Beta API)                                                  │
│    PATCH /admin/people/profilePropertySettings/{id}                         │
│    { prioritizedSourceUrls: [...] }                                         │
│                                                                             │
│    Determines which source takes precedence when data conflicts             │
└─────────────────────────────────────────────────────────────────────────────┘
                                    │
                                    ▼
┌─────────────────────────────────────────────────────────────────────────────┐
│ 5. INGEST ITEMS (External Items)                                            │
│    PUT /external/connections/{id}/items/{itemId}                            │
│    { id, properties, acl, content? }                                        │
│                                                                             │
│    Each item = one person's enrichment data, linked via personAccount       │
└─────────────────────────────────────────────────────────────────────────────┘
                                    │
                                    ▼
┌─────────────────────────────────────────────────────────────────────────────┐
│ 6. MICROSOFT 365 PROCESSES DATA (6+ hours)                                  │
│    - Indexes items for search                                               │
│    - Maps items to users via personAccount label                            │
│    - Surfaces in Copilot, Search, Profile Cards                             │
└─────────────────────────────────────────────────────────────────────────────┘
```

---

## 1. contentCategory - Where Is It Set?

### Location in Code
**File**: `src/people-connector/connection-manager.ts` (lines 32-44)

```typescript
const connectionRequest = {
  id: this.connectionId,
  name,
  description,
  // REQUIRED for people data connectors
  contentCategory: 'people',
  activitySettings: {
    urlToItemResolvers: []
  }
};

// CRITICAL: Must use beta API for connection creation
await this.betaClient.api('/external/connections').post(connectionRequest);
```

**IMPORTANT**: The `contentCategory: 'people'` parameter is only recognized by the **beta API**. Using v1.0 API will result in `uncategorized`.

### Why It Matters
From Microsoft docs:
> "Copilot connectors that contains people data must set the `contentCategory` property to have the value `people`."

Without `contentCategory: 'people'`:
- Connection is treated as a regular search connector
- People data labels are NOT recognized
- Data won't appear in profile cards or Copilot people experiences

**Status in our code**: ✅ Correctly implemented

---

## 2. string vs stringCollection - Key Differences

### Type: `string`
- **Single value** per property
- Used for properties with one piece of data
- JSON serialized as a single string

**Examples**:
| Label | Use Case |
|-------|----------|
| `personAccount` | One user mapping per item |
| `personNote` | One "About Me" note |
| `personWebSite` | One personal website |
| `personCurrentPosition` | Current job title |

**Format in properties**:
```json
{
  "accountInformation": "{\"userPrincipalName\":\"user@domain.com\"}",
  "aboutMe": "{\"detail\":{\"contentType\":\"text\",\"content\":\"About me text...\"}}"
}
```

### Type: `stringCollection`
- **Multiple values** per property (array)
- Used for properties where a person can have many
- Each array element is a JSON serialized entity
- **MUST include `@odata.type` annotation**

**Examples**:
| Label | Use Case |
|-------|----------|
| `personSkills` | Multiple skills per person |
| `personCertifications` | Multiple certifications |
| `personAwards` | Multiple awards |
| `personProjects` | Multiple projects |
| `personPhones` | Up to N phone numbers |
| `personEmails` | Up to 3 emails |
| `personAddresses` | Max 3 (Home, Work, Other) |

**Format in properties**:
```json
{
  "skills@odata.type": "Collection(String)",
  "skills": [
    "{\"displayName\":\"TypeScript\"}",
    "{\"displayName\":\"Azure\"}",
    "{\"displayName\":\"Project Management\"}"
  ]
}
```

### Critical: The @odata.type Annotation

For `stringCollection` properties, you MUST include:
```json
"propertyName@odata.type": "Collection(String)"
```

Without this annotation, Microsoft Graph may not correctly parse the array.

### Schema Definition Difference

```json
// string property
{
  "name": "aboutMe",
  "type": "string",
  "labels": ["personNote"]
}

// stringCollection property
{
  "name": "skills",
  "type": "stringCollection",
  "labels": ["personSkills"]
}
```

---

## 3. ACL (Access Control List) - Deep Dive

### What is ACL?
The ACL determines **who can see** the external item in Microsoft 365 experiences.

### Structure
```json
{
  "acl": [
    {
      "type": "everyone" | "user" | "group",
      "value": "identifier",
      "accessType": "grant" | "deny"
    }
  ]
}
```

### ACL Types

| Type | Value | Description |
|------|-------|-------------|
| `everyone` | `"everyone"` | All users in the tenant |
| `user` | Microsoft Entra user ID | Specific user |
| `group` | Microsoft Entra group ID | Specific group |

### Access Types

| accessType | Effect |
|------------|--------|
| `grant` | Allows access |
| `deny` | Blocks access (takes precedence over grant) |

### Example: Everyone Can See
```json
{
  "acl": [
    {
      "type": "everyone",
      "value": "everyone",
      "accessType": "grant"
    }
  ]
}
```

### Example: Grant Everyone, Deny Specific User
```json
{
  "acl": [
    {
      "type": "everyone",
      "value": "everyone",
      "accessType": "grant"
    },
    {
      "type": "user",
      "value": "12345678-1234-1234-1234-123456789012",
      "accessType": "deny"
    }
  ]
}
```
**Result**: Everyone can see EXCEPT that specific user (deny takes precedence)

### For People Data Connectors
From Microsoft docs:
> "You must set the access control list (ACL) on items ingested by the connector to grant access to everyone."

**Required for people connectors**:
```json
"acl": [{ "type": "everyone", "value": "everyone", "accessType": "grant" }]
```

### Why "Everyone" for People Data?
- People data (skills, notes) should be visible organization-wide
- Profile cards show data to all colleagues
- Copilot needs access to reason over people data
- If you want to restrict visibility, use Information Barriers at the tenant level instead

---

## 4. External Items - Complete Anatomy

### What is an External Item?
An `externalItem` is the fundamental unit of data in a Graph Connector. For people connectors, **one item = one person's enrichment data**.

### Item Structure

```json
{
  "id": "unique-item-id",
  "properties": {
    // Schema-defined properties with optional labels
  },
  "acl": [
    // Access control entries
  ],
  "content": {
    // Optional full-text searchable content
    "value": "text content",
    "type": "text" | "html"
  },
  "activities": [
    // Optional activity tracking
  ]
}
```

### Component Deep-Dive

#### 1. Item ID
- **Unique identifier** for the item within the connection
- Used to create, update, or delete the item
- Best practice: Derive from email address

```typescript
const itemId = `person-${email.replace(/@/g, '-').replace(/\./g, '-')}`;
// "person-john-smith-contoso-com"
```

#### 2. Properties
The structured data fields defined by your schema.

**For People Connectors**:
```json
{
  "properties": {
    // REQUIRED: Maps item to a Microsoft Entra user
    "accountInformation": "{\"userPrincipalName\":\"user@domain.com\"}",

    // LABELED: Recognized by Microsoft 365 People experiences
    "skills@odata.type": "Collection(String)",
    "skills": [
      "{\"displayName\":\"Skill 1\"}",
      "{\"displayName\":\"Skill 2\"}"
    ],
    "aboutMe": "{\"detail\":{\"contentType\":\"text\",\"content\":\"About text\"}}",

    // CUSTOM: No label = custom searchable property
    "VTeam": "Platform Team",
    "CostCenter": "ENG-001"
  }
}
```

**Key Rules**:
1. `accountInformation` with `personAccount` label is **REQUIRED** - maps item to user
2. Labeled properties use JSON-serialized profile entities
3. Custom properties (no label) are searchable as custom fields
4. `stringCollection` requires `@odata.type` annotation

#### 3. Content (Optional for People Data)
Full-text searchable content blob.

```json
{
  "content": {
    "value": "John is an expert in TypeScript and Azure...",
    "type": "text"
  }
}
```

**For people data connectors**: Content is optional because:
- User mapping relies on `personAccount` in properties
- Labeled properties provide structured searchability
- Microsoft may ignore content for people connectors

#### 4. ACL
As explained above - must be "everyone" for people data.

#### 5. Activities (Optional)
Track user interactions with the item:
- `created`, `modified`, `commented`, `viewed`, etc.
- Powers intelligent recommendations

---

## 5. JSON Serialization Format for Profile Entities

### personAccount (userAccountInformation)
```json
{
  "userPrincipalName": "user@domain.com"
}
// OR
{
  "externalDirectoryObjectId": "azure-ad-object-id"
}
```

### personSkills (skillProficiency)
```json
{
  "displayName": "TypeScript"
}
// Full entity (optional fields):
{
  "displayName": "TypeScript",
  "proficiency": "expert",
  "webUrl": "https://..."
}
```

### personNote (personAnnotation)
```json
{
  "detail": {
    "contentType": "text",
    "content": "About me text here..."
  }
}
// OR with displayName:
{
  "displayName": "About Me",
  "detail": {
    "contentType": "text",
    "content": "About me text..."
  }
}
```

### personCertifications (personCertification)
```json
{
  "displayName": "AWS Solutions Architect",
  "issuedDate": "2024-01-15",
  "issuingAuthority": "Amazon Web Services"
}
```

### personAwards (personAward)
```json
{
  "displayName": "Employee of the Year",
  "issuedDate": "2025-03-01",
  "issuingAuthority": "Contoso Inc"
}
```

### personProjects (projectParticipation)
```json
{
  "displayName": "Project Atlas",
  "detail": "Led development of core platform"
}
```

---

## 6. Complete External Item Example

```json
{
  "id": "person-sarah-chen-contoso-com",
  "properties": {
    "accountInformation": "{\"userPrincipalName\":\"sarah.chen@contoso.com\"}",

    "skills@odata.type": "Collection(String)",
    "skills": [
      "{\"displayName\":\"Strategic Planning\"}",
      "{\"displayName\":\"Team Leadership\"}",
      "{\"displayName\":\"Azure Architecture\"}"
    ],

    "aboutMe": "{\"detail\":{\"contentType\":\"text\",\"content\":\"CEO with 15+ years in technology leadership. Passionate about building inclusive teams.\"}}",

    "certifications@odata.type": "Collection(String)",
    "certifications": [
      "{\"displayName\":\"PMP\",\"issuingAuthority\":\"PMI\"}"
    ],

    "VTeam": "Executive Leadership",
    "CostCenter": "EXEC-001",
    "BuildingAccess": "Level-5"
  },
  "acl": [
    {
      "type": "everyone",
      "value": "everyone",
      "accessType": "grant"
    }
  ]
}
```

---

## 7. Available People Data Labels

All 13 official people data labels are enabled and validated (m365people23, 2026-02-27).

| Label | Type | Status | Entity | CSV Source |
|-------|------|--------|--------|-----------|
| `personAccount` | string | ✅ Required | userAccountInformation | email (auto) |
| `personSkills` | stringCollection | ✅ Implemented | skillProficiency | skills |
| `personNote` | string | ✅ Implemented | personAnnotation | aboutMe |
| `personCertifications` | stringCollection | ✅ Implemented | personCertification | certifications |
| `personProjects` | stringCollection | ✅ Implemented | projectParticipation | projects |
| `personAwards` | stringCollection | ✅ Implemented | personAward | awards |
| `personAnniversaries` | stringCollection | ✅ Implemented | personAnniversary | birthday |
| `personWebSite` | string | ✅ Implemented | webSite | mySite |
| `personName` | string | ✅ Implemented (composite) | personName | givenName + surname |
| `personCurrentPosition` | string | ✅ Implemented (composite) | workPosition | jobTitle + companyName |
| `personAddresses` | stringCollection | ✅ Implemented (composite) | itemAddress | streetAddress + city + state + country + postalCode |
| `personEmails` | stringCollection | ✅ Implemented (composite) | itemEmail | mail |
| `personPhones` | stringCollection | ✅ Implemented (composite) | itemPhone | mobilePhone + businessPhones |
| `personWebAccounts` | stringCollection | ✅ Declared (no CSV data) | webAccount | *(none)* |

**NOT available as labels**: `personLanguages`, `personInterests` — these have no people data labels and must use the Profile API (Option A).

### Path A vs Path B Deserialization

Confirmed by Microsoft engineering:
- **Path A (with label)**: `JsonElement` deserialization — handles any JSON type safely. All 13 labels above use this path.
- **Path B (without label)**: `Blob` deserialization — expects `string` only. Custom properties (VTeam, etc.) use this path.
- **CRITICAL**: Never use `stringCollection` without a label. It will crash the entire connection.

See `docs/graph-connector-lessons.md` for full history and lessons from m365people14–23.

---

## 8. Verification Checklist

Before testing, ensure:

- [ ] Connection has `contentCategory: 'people'`
- [ ] Schema has `personAccount` labeled property
- [ ] Schema uses correct types (`string` vs `stringCollection`)
- [ ] Items have `accountInformation` with valid `userPrincipalName`
- [ ] Items have ACL set to `everyone`/`grant`
- [ ] `stringCollection` properties include `@odata.type` annotation
- [ ] JSON entities are properly serialized strings
- [ ] Connection is registered as profile source
- [ ] Connection is in prioritized sources list
- [ ] Waited 6+ hours after ingestion for indexing

---

## 9. Implementation Verification Results

### Requirements Checklist

| Requirement | Status | Location | Notes |
|-----------|--------|----------|-------|
| `contentCategory: 'people'` | ✅ | connection-manager.ts:37 | Correct |
| `personAccount` label + string type | ✅ | enrich-connector.ts:381 | Correct |
| Schema types (string vs stringCollection) | ✅ | enrich-connector.ts:384-390 | Correct |
| `accountInformation` with `userPrincipalName` | ✅ | enrich-connector.ts:419-420 | Correct |
| ACL: `everyone`/`grant` | ✅ | enrich-connector.ts:454-458 | Correct |
| `@odata.type` on stringCollection | ✅ | enrich-connector.ts:426 | Correct |
| JSON serialization (skills) | ✅ | enrich-connector.ts:427-429 | `{displayName: ...}` |
| JSON serialization (aboutMe) | ✅ | enrich-connector.ts:434-436 | `{detail: {contentType, content}}` |
| Profile source registration | ✅ | connection-manager.ts:137-213 | Beta API |
| Profile source prioritization | ✅ | connection-manager.ts:161-191 | Beta API |

**All 10 requirements are correctly implemented!** ✅

### Verified: `aboutMe` Serialization Format is CORRECT ✅

**Current implementation** (`enrich-connector.ts` line 434-436):
```typescript
properties.aboutMe = JSON.stringify({
  detail: { contentType: 'text', content: profile.aboutMe }
});
```

**Produces**:
```json
"aboutMe": "{\"detail\":{\"contentType\":\"text\",\"content\":\"About me text...\"}}"
```

**Verification**:

The `personAnnotation` entity (Microsoft Graph) has:
```json
{
  "detail": {
    "@odata.type": "microsoft.graph.itemBody"
  },
  "displayName": "String"
}
```

And `itemBody` is defined as:
```json
{
  "content": "String",
  "contentType": "text" | "html"
}
```

**Conclusion**: Our format `{detail: {contentType: 'text', content: '...'}}` correctly represents the `personAnnotation.detail` (itemBody) structure.

---

## 10. Testing Plan

### Step 1: Check Current Connector Status
```bash
# List existing connections
node tools/debug/test-list-connections.mjs
```

### Step 2: Delete Existing Connector (if schema changed)
Schema changes require deletion and recreation:
```bash
# Option A: Via npm script
npm run enrich:delete-connector

# Option B: Via debug tool with custom connection ID
node tools/admin/delete-connection.mjs m365provisionpeople
```

### Step 3: Wait for Deletion (~5-15 minutes)
Microsoft takes time to fully remove the connection and schema.

### Step 4: Run Setup with New Schema
```bash
npm run option-b:ingest -- --csv config/textcraft-europe.csv --setup --connection-id m365people3
```

**CRITICAL**: If schema gets stuck in "draft" state (>60s), use the wait-and-ingest script:
```bash
node tools/admin/wait-and-ingest.mjs m365people3 config/textcraft-europe.csv
```

This will:
1. Create new connection with `contentCategory: 'people'`
2. Register schema with `personSkills` and `personNote` labels
3. Register as profile source
4. Add to prioritized sources
5. Ingest items for all users

### Step 5: Wait for Indexing (6+ hours)
From Microsoft docs: "Microsoft 365 might take up to 6 hours after the connection is created before it becomes available in search, people experiences, or Copilot."

### Step 6: Test Copilot Search
After indexing, test with queries like:
- "Find people with skills in Strategic Planning"
- "Who has Polish Proofreading skills?"
- "Find people on VTeam Alpha"
- "Who is on the Executive Leadership team?"

### Expected Results
- Copilot should return **confirmed** matches (not "not confirmed")
- Skills should be searchable
- AboutMe/notes should be searchable
- Custom properties (VTeam, CostCenter) should be searchable

---

## Summary

| Concept | Key Point |
|---------|-----------|
| **contentCategory** | Must be `'people'` - set in connection creation |
| **string** | Single JSON-serialized value |
| **stringCollection** | Array of JSON-serialized values + `@odata.type` |
| **ACL** | Must be `everyone`/`grant` for people data |
| **External Item** | One item per person, linked via `personAccount` |
| **Profile Source** | Must register connection as profile source |
| **Indexing** | Takes 6+ hours after ingestion |

---

## 11. Critical Implementation Notes

### Profile Source Registration: `kind` Property is REQUIRED

> ⚠️ **CRITICAL FIX**: The `kind: 'Connector'` property is **REQUIRED** in the profile source registration payload.

When registering a Graph Connector as a profile source, the payload MUST include:

```typescript
// WRONG - causes "Profile source exists but missing kind property" error loop
const profileSourcePayload = {
  sourceId: this.connectionId,
  displayName,
  webUrl,
};

// CORRECT - kind: 'Connector' is REQUIRED per Microsoft docs
const profileSourcePayload = {
  sourceId: this.connectionId,
  displayName,
  kind: 'Connector',  // <-- REQUIRED
  webUrl,
};
```

**Without `kind: 'Connector'`**:
- Profile source may be created but not properly linked
- Data won't appear in profile cards or Copilot
- API may accept the request but produce undefined behavior

**Location in code**: `src/people-connector/connection-manager.ts`

### Beta API Requirement
The `contentCategory: 'people'` parameter is **only supported by the beta API**. Using the v1.0 API will result in the connection being created with `contentCategory: 'uncategorized'`, which causes:
- `personAccount`, `personSkills`, `personNote` labels to show as `unknownFutureValue`
- User mapping to fail (Users: - in Admin Portal)
- Data not appearing in profile cards or Copilot

**Fix applied**: `connection-manager.ts` now uses `betaClient` for connection creation.

### Labels Display as unknownFutureValue
When querying schema labels via API, they may display as `unknownFutureValue`. This is normal behavior - Microsoft's internal systems recognize the labels correctly when `contentCategory: 'people'` is set.

### Schema Provisioning Time
Schema can take 1-5 minutes to provision. Use the `wait-and-ingest.mjs` script if the main command times out.

### Old Connection Cleanup
When recreating connections, ensure old connections are **fully deleted** (wait 5-15 minutes) before creating new ones with the same ID. Alternatively, use a new connection ID.

---

## 12. Critical Learnings: Profile Source Propagation and Prioritization

### Lesson 1: Separate Registration from Ingestion

Microsoft's internal pipeline (TSS — Tenant Settings Service) must discover the profile source **before** it can process ingested items. When registration and ingestion happen in the same flow with no delay:

**Internal debug log sequence (BROKEN)**:
```
ProfileSourceRegistrar: Failed to retrieve settings from TSS, statusCode=Unauthorized
Profile source registration failed for item, but transformation will continue
CAPIv2 export completed with status 'Created', profileSyncEnabled=False, profileSynced=False
```

**Internal debug log sequence (WORKING)**:
```
Profile source already exists with kind property, no action needed
Profile source registered successfully for item
CAPIv2 export completed with status 'Created', profileSyncEnabled=True, profileSynced=True
```

**Key flag**: `profileSyncEnabled=True` is the indicator that Microsoft's backend will actually sync connector data to user profiles. Without it, CAPIv2 export still happens (status `Created`) but data never reaches `/users/{id}/profile/skills`.

**Microsoft's recommendation**: Use separate commands or separate applications for registration vs ingestion. Their official sample (`msgraph-people-connector-sample-dotnet`) uses 3 separate CLI commands: `setup`, `register`, `sync`.

**Our implementation**: 60-second polling wait after registration, verifying the source is readable via `GET /admin/people/profileSources` before proceeding to ingestion.

### Lesson 2: Never Remove Non-Connector Sources from Prioritized List

The `prioritizedSourceUrls` array contains URLs for ALL profile sources, including Microsoft's **internal AAD default source**:

```
https://graph.microsoft.com/beta/admin/people/profileSources(sourceId='4ce763dd-9214-4eff-af7c-da491cc3782d')
```

This UUID-format source ID is **NOT** an external connection. Calling `GET /external/connections/4ce763dd-...` returns 404. If your stale-cleanup code removes it:

1. You PATCH `prioritizedSourceUrls` with just your connector URL
2. Microsoft silently rejects or reverts this (AAD source is required)
3. Your connector never appears in the prioritized list
4. All items get `profileSyncEnabled=False` — data never reaches profiles

**Rule**: Only attempt stale-cleanup on sourceIds that match the external connector pattern (starts with a letter, alphanumeric only like `m365people14`). UUID-format sourceIds are internal Microsoft sources and must ALWAYS be preserved.

### Lesson 3: The `positions` Collection Proves Connector Works

Even when skills/notes don't sync, the `positions` collection on a user profile may show `ConnectorSource` as a `lastModifiedBy` source contributing `employeeType` and `employeeId`. This proves the connector CAN write to profiles — the issue is specifically with skills/notes not being synced due to missing prioritization.

---

**Last Updated**: 2026-02-16
