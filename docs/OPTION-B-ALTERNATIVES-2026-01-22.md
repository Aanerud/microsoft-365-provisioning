# Option B Alternatives - Research Findings

**Date**: 2026-01-22 (Updated 2026-01-26)
**Status**: ✅ Solution Implemented - Hybrid Approach with People Data Labels

## Summary

Graph Connectors are now working reliably. We discovered a critical insight: **Profile API data is NOT Copilot-searchable** because it's stored as `source.type: "User"` with `isSearchable: false`.

**Solution**: Use Graph Connectors with **people data labels** for Copilot searchability:
- `personSkills` label for skills
- `personNote` label for aboutMe/notes
- Custom properties (no label) are also searchable

**Limitation**: Languages and interests have no people data labels - they remain Profile API only (visible on cards, NOT Copilot-searchable).

## Tested Approaches

### 1. Graph Connectors (Current Option B)
**Status**: 🔴 Partially Working / Unreliable

```
PUT /external/connections/{id}/items/{itemId}
```

**Results**:
- Some items ingest successfully, others fail with 500 errors
- GET requests to verify items also return 500 errors
- Appears to be Microsoft service-side issue
- When working: Data surfaces in Copilot via People Data labels

**Auth**: Application permissions (client secret)
- `ExternalConnection.ReadWrite.OwnedBy`
- `ExternalItem.ReadWrite.OwnedBy`

### 2. Direct User PATCH (NEW - Tested & Working)
**Status**: ✅ Working

```
PATCH /users/{email}
{ "skills": [...], "aboutMe": "...", "interests": [...] }
```

**Test Results**:
```
1. PATCH /users/{id} with skills:    Status: 204 Success!
2. PATCH /users/{id} with aboutMe:   Status: 204 Success!
3. PATCH /users/{id} with interests: Status: 204 Success!
```

**Verified Data on Marie Dubois**:
```json
{
  "displayName": "Marie Dubois",
  "aboutMe": "Financial executive with background in media and publishing...",
  "interests": ["European Finance", "International Taxation", "Media Industry"],
  "skills": ["Financial Planning", "Budget Management", "M&A", "Risk Assessment", "International Taxation"]
}
```

**Auth**: Delegated permissions only (browser sign-in)
- `User.ReadWrite.All`
- Same auth as Option A

**Requirements**:
- Each property type (skills, aboutMe, interests) must be in SEPARATE PATCH requests
- Cannot combine with other user properties in same request
- Application permissions NOT supported

### 3. Profile API
**Status**: ⚠️ Not Tested (delegated only)

```
POST /users/{id}/profile/skills
```

**Pros**:
- Rich structure (proficiency levels, categories, collaboration tags)
- Each skill is a separate entity with ID

**Cons**:
- Delegated permissions only
- Beta API only
- More complex than direct PATCH

### 4. Open Extensions
**Status**: ⚠️ Deprecated (superseded by Option B connector)

```
POST /users/{id}/extensions
```

**Pros**:
- Works with both delegated and application permissions
- Already implemented in `src/extensions/open-extension-manager.ts`
- Batch support

**Cons**:
- Data does NOT surface in Copilot (no People Data labels)
- Custom storage only

## Comparison Matrix

| Feature | Graph Connectors | Direct PATCH | Profile API | Open Extensions |
|---------|------------------|--------------|-------------|-----------------|
| **Status** | ✅ Working | ✅ Working | ✅ Working | ✅ Working |
| **Auth** | Application | Delegated | Delegated | Both |
| **Storage** | External index | User object | Profile resource | Extension |
| **Copilot searchable** | **Yes** (with labels) | No | **No** (isSearchable:false) | No |
| **Batch** | No | No | No | Yes |
| **Complexity** | High | Low | Medium | Low |

### Critical Finding: Profile API isSearchable Issue

Data written via Profile API with delegated auth is stored as:
```json
{
  "source": {
    "type": ["User"]
  },
  "isSearchable": false
}
```

This means **Profile API data is NOT Copilot-searchable**, even though it appears on profile cards. Only data from system sources (Graph Connectors with people data labels) has `isSearchable: true`.

## Recommended Path Forward

### Implemented Solution: Hybrid Approach (2026-01-26)

Use **both** Profile API and Graph Connectors for optimal coverage:

```typescript
// PHASE 1: Profile API (Delegated Auth)
// For languages, interests (no connector labels available)
// Also writes skills/aboutMe for profile card redundancy
await profileWriter.writeLanguages(userId, languages);
await profileWriter.writeInterests(userId, interests);
await profileWriter.writeSkills(userId, skills);
await profileWriter.writeNotes(userId, aboutMe);

// PHASE 2: Graph Connectors (App-Only Auth)
// For Copilot searchability with people data labels
await connector.ingestItem({
  properties: {
    accountInformation: JSON.stringify({ userPrincipalName: email }),
    skills: skills.map(s => JSON.stringify({ displayName: s })),  // personSkills label
    aboutMe: JSON.stringify({ detail: { contentType: 'text', content: aboutMe } }),  // personNote label
    VTeam: vteam,  // custom property (searchable without label)
  }
});
```

**Why Hybrid?**:
- Profile API: Visible on profile cards immediately
- Graph Connectors: Copilot-searchable with people data labels
- Languages/interests: No connector labels exist - Profile API only

### Available People Data Labels

From [Microsoft documentation](https://learn.microsoft.com/en-us/microsoft-365-copilot/extensibility/build-connectors-with-people-data):

| Label | Type | Used For |
|-------|------|----------|
| `personAccount` | string | Required - maps item to user |
| `personSkills` | stringCollection | Skills (Copilot-searchable) |
| `personNote` | string | AboutMe/notes (Copilot-searchable) |
| `personCertifications` | stringCollection | Certifications |
| `personAwards` | stringCollection | Awards |
| `personProjects` | stringCollection | Projects |

**NOT available**: `personLanguages`, `personInterests` (platform limitation)

## Properties That Can Be Written via Direct PATCH

| Property | Type | Notes |
|----------|------|-------|
| `skills` | string[] | Must be separate PATCH |
| `interests` | string[] | Must be separate PATCH |
| `aboutMe` | string | Must be separate PATCH |
| `responsibilities` | string[] | Must be separate PATCH |
| `pastProjects` | string[] | Must be separate PATCH |
| `schools` | string[] | Must be separate PATCH |
| `mySite` | string | Must be separate PATCH |
| `birthday` | date | Must be separate PATCH |

**Important**: These properties require delegated permissions - cannot be set with application-only auth.

## CSV Column Mapping for Direct PATCH

From `textcraft-europe.csv`:

| CSV Column | Maps To | Type |
|------------|---------|------|
| `skills` | `skills` | string[] |
| `languages` | Could map to `interests` or custom | string[] |
| `aboutMe` | `aboutMe` | string |

## Implementation Notes

### Current Auth Flow

Option A uses browser-based OAuth (delegated):
- Token cached in `~/.m365-provision/token-cache.json`
- Has `User.ReadWrite.All` permission
- Can write profile properties

### Code Location

Profile enrichment could be added to:
- `src/provision.ts` - After batch user creation
- Or new `src/enrich-profiles-direct.ts` - Separate command

### Rate Limiting

Since each property requires separate PATCH:
- 95 users × 3 properties = 285 API calls
- Add delay between calls (100-200ms)
- Use retry logic for throttling

## Test Command Used

```javascript
// Verified working with Marie Dubois
PATCH /users/marie.dubois@domain.onmicrosoft.com
{ "skills": ["Financial Planning", "Budget Management", ...] }
// Status: 204 Success

PATCH /users/marie.dubois@domain.onmicrosoft.com
{ "aboutMe": "Financial executive with background..." }
// Status: 204 Success

PATCH /users/marie.dubois@domain.onmicrosoft.com
{ "interests": ["European Finance", ...] }
// Status: 204 Success
```

## Implementation Status (2026-01-26)

### Completed
1. ✅ Hybrid enrichment implemented in `src/enrich-profiles-hybrid.ts`
2. ✅ Schema updated with `personSkills` and `personNote` labels
3. ✅ Item creation includes skills and aboutMe with proper JSON format
4. ✅ Profile API writes languages/interests (no connector labels available)

### To Verify Copilot Searchability
1. Delete existing connector: `npm run enrich:delete-connector` or via Azure Portal
2. Wait for deletion (~5-15 minutes)
3. Setup with new schema: `npm run enrich:connector-only -- --csv config/textcraft-europe.csv --setup`
4. Wait 6+ hours for indexing
5. Test with Copilot: "Find people with skills in Polish Proofreading"

### Searchability Summary

| Data Type | Method | Copilot Searchable |
|-----------|--------|-------------------|
| Skills | Connector (`personSkills`) | **Yes** |
| AboutMe/Notes | Connector (`personNote`) | **Yes** |
| Custom props (VTeam, etc.) | Connector (no label) | **Yes** |
| Languages | Profile API only | **No** (no label available) |
| Interests | Profile API only | **No** (no label available) |

---

**Key Finding**: Profile API data has `isSearchable: false`. Graph Connectors with people data labels are the ONLY way to make profile data Copilot-searchable. Languages/interests cannot be made searchable due to Microsoft platform limitation (no connector labels exist for these types).
