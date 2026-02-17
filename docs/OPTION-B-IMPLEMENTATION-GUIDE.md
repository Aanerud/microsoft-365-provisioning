# Option B: Profile Enrichment via Microsoft Graph Connectors

**Status**: ✅ Implemented and Working (2026-01-26)

## Overview

Option B enriches Microsoft 365 user profiles by ingesting additional data through **Microsoft Graph Connectors with People Data labels**. This surfaces enrichment data in:

- Microsoft 365 profile cards
- Microsoft Search
- **Microsoft 365 Copilot responses** (requires people data labels)
- Teams member discovery

> **Note**: Profile data propagation to the `/me/profile` API can take 1-24 hours after ingestion.

### Key Insight: Copilot Searchability

Data written via Profile API with delegated auth is stored as `source.type: "User"` with `isSearchable: false`. **Only data from system sources (connectors with people data labels) is Copilot-searchable.**

This is why we use Graph Connectors with people data labels for skills and notes - they become searchable by Copilot.

## Architecture

```
┌────────────────────────────────────────────────────────────┐
│ Option A: Core User Provisioning                           │
│ File: src/provision.ts                                     │
│                                                             │
│ - Creates Entra ID users                                   │
│ - Sets standard properties (jobTitle, department, etc.)    │
│ - Assigns manager relationships                            │
│ - Assigns licenses                                         │
│ - NO enrichment data (handled by Hybrid Enrichment)        │
│                                                             │
│ Output: Real Entra ID users with standard properties       │
└─────────────────┬───────────────────────────────────────────┘
                  │
                  │ Users exist in Entra ID (prerequisite)
                  ↓
┌────────────────────────────────────────────────────────────┐
│ Option A: Profile API Enrichment (Delegated Auth)          │
│ File: src/enrich-profiles.ts                               │
│                                                             │
│ - Languages, interests (no connector labels available)     │
│ - Visible on profile cards, NOT Copilot-searchable         │
└─────────────────┬───────────────────────────────────────────┘
                  │
                  │ Users enriched with languages/interests
                  ↓
┌────────────────────────────────────────────────────────────┐
│ Option B: Graph Connector Pipeline (App-Only Auth)         │
│ File: src/enrich-connector.ts                              │
│                                                             │
│ - skills (personSkills label) → Copilot-searchable!        │
│ - aboutMe (personNote label) → Copilot-searchable!         │
│ - projects/awards/certs (people labels) → searchable       │
│ - Links items to Entra ID users via accountInformation     │
│ - Extra custom columns ignored (strict-by-doc)             │
│                                                             │
│ Output: Copilot search integration                         │
└────────────────────────────────────────────────────────────┘
```

## Data Classification

### Standard Properties (Option A)
Handled by `provision.ts` - written directly to Entra ID user object:
- givenName, surname, displayName
- jobTitle, department, employeeType
- companyName, officeLocation
- usageLocation, preferredLanguage
- mail, mobilePhone, businessPhones

### Enrichment Properties
Handled by `enrich-profiles.ts` (Profile API) and `enrich-connector.ts` (Graph Connectors):

**Graph Connector with People Data Labels** (Copilot-searchable):
| Property | Label | Copilot Searchable |
|----------|-------|-------------------|
| `skills` | `personSkills` | **Yes** |
| `aboutMe` | `personNote` | **Yes** |
| `pastProjects` | `personProjects` | Yes |
| `certifications` | `personCertifications` | Yes |
| `awards` | `personAwards` | Yes |
| `mySite` | `personWebSite` | Yes |
| `birthday` | `personAnniversaries` | Yes |

**Custom Columns (Ignored by Connector)**:
Extra CSV columns without people data labels are ignored in strict-by-doc mode.

**Profile API Only** (no connector labels available - NOT Copilot-searchable):
| Property | Visible on Profile Card | Copilot Searchable |
|----------|------------------------|-------------------|
| `languages` | Yes | **No** (platform limitation) |
| `interests` | Yes | **No** (platform limitation) |

> **Important**: Languages and interests have no people data labels in Microsoft's connector framework, so they cannot be made Copilot-searchable. They remain Profile API only.

## Prerequisites

### 1. Option A Must Run First

Users must exist in Entra ID before enriching their profiles. Option B links external items to Entra ID users via `userPrincipalName`.

To avoid backend OID lookups, Option A now builds an **OID cache** that maps `userPrincipalName` → `externalDirectoryObjectId`. Option B loads this cache automatically (and will prompt for delegated sign-in to create it if missing).

```bash
# First: Create users
npm run provision -- --csv config/textcraft-europe.csv

# Then: Enrich profiles
npm run option-b:ingest -- --csv config/textcraft-europe.csv
```

**OID cache file**:
```
config/textcraft-europe_oid_cache.json
```

You can build it manually if needed:
```bash
npm run build-oid-cache -- --csv config/textcraft-europe.csv
```

### 2. Azure AD Configuration

**App Registration Requirements**:

**Application Permissions** (required):
- `ExternalConnection.ReadWrite.OwnedBy` - Create/manage Graph Connector connections
- `ExternalItem.ReadWrite.OwnedBy` - Create/manage external items
- `PeopleSettings.ReadWrite.All` - Register profile source and configure prioritization (beta API)

**Admin Consent**: Required (Global Admin or delegated admin)

**Client Secret**:
1. Azure Portal > Your App > Certificates & secrets
2. New client secret
3. Copy the **Value** (not Secret ID)
4. Add to `.env`:
   ```bash
   AZURE_CLIENT_SECRET=your-secret-value-here
   ```

### 3. CSV File

Use the same CSV file for both Option A and Option B:

```csv
name,email,givenName,surname,jobTitle,department,skills,interests,aboutMe,VTeam,BenefitPlan
Sarah Chen,sarah@domain.com,Sarah,Chen,CEO,Executive,"['Leadership','Strategy']","['Innovation']",Experienced executive...,Platform Team,Premium
```

## Implementation

### Step 1: First-Time Setup

Create the connection and register the schema (only needs to be done once):

```bash
npm run option-b:setup
```

**What this does**:
1. Creates Graph Connector connection (`m365provisionpeople`)
2. **Registers as profile source** (beta API: `/admin/people/profileSources`)
   - Links the connector to People Data in M365
   - Required for data to appear in user profiles
3. **Adds to prioritized profile sources** (beta API: `/admin/people/profilePropertySettings`)
   - Sets connector as highest priority (index 0)
   - Ensures connector data takes precedence over other sources
4. Registers schema with people data labeled properties only:
   - Required: `accountInformation` (links to Entra ID user)
   - People-labeled properties present in CSV (skills, aboutMe, projects, etc.)
   - Extra custom columns are ignored
5. Waits for schema to be ready (can take 5-10 minutes)

**Output**:
```
✓ Created connection: m365provisionpeople
📋 Registering connection as profile source (beta API)...
✓ Registered as profile source
✓ Added to prioritized profile sources (highest priority)
✓ Schema registration initiated
  Schema state: draft, waiting...
  Schema state: ready
✓ Schema is ready
```

> **Important**: Profile source registration requires the `PeopleSettings.ReadWrite.All` application permission. Without this, data will be ingested but may not appear in the `/me/profile` API or profile cards.

### Step 2: Ingest People Data

Ingest enrichment data from your CSV file:

```bash
npm run option-b:ingest -- --csv config/textcraft-europe.csv
```

**What this does**:
1. Authenticates with OAuth 2.0 Client Credentials Flow
2. Verifies profile source registration and prioritization
3. Loads CSV and parses enrichment properties
4. Creates external items for each person
5. Ingests items to Graph Connector (using beta endpoint)
6. Links items to Entra ID users via email address and cached OID (if available)
7. Automatically detects and deletes orphaned items (items not in CSV)

**Output**:
```
🚀 Profile Enrichment (Option B)

Configuration:
  CSV: config/agents-template.csv
  Connection ID: m365provisionpeople
  Dry Run: No

🔐 Using application-only authentication (OAuth 2.0 client credentials)...
✓ Authenticated with client credentials

📖 Loading people data from CSV...
✓ Loaded 20 people

🔨 Creating external items...

📤 Ingesting 20 items (batch size: 20)...

✅ Ingested: person-sarah-chen-domain-com
✅ Ingested: person-john-doe-domain-com
...

🔍 Checking for orphaned external items...
✓ No orphaned items found

============================================================
📊 Enrichment Summary
============================================================
✅ Successful: 20
❌ Failed:     0
🗑️  Deleted:    0
============================================================
```

### Step 3 (Optional): Dry Run

Preview what would be created without actually ingesting:

```bash
npm run option-b:dry-run
```

**Output**: Shows sample external item structure with all properties.

## External Item Structure

Each person is converted to an external item with this structure:

```json
{
  "id": "person-email-domain-com",
  "properties": {
    "accountInformation": "{\"userPrincipalName\":\"email@domain.com\"}",
    "skills@odata.type": "Collection(String)",
    "skills": [
      "{\"displayName\":\"TypeScript\"}",
      "{\"displayName\":\"Azure\"}"
    ],
    "pastProjects@odata.type": "Collection(String)",
    "pastProjects": [
      "{\"displayName\":\"Migration\"}"
    ],
    "aboutMe": "{\"detail\":{\"contentType\":\"text\",\"content\":\"Experienced engineer...\"}}"
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

### Key Points

1. **Item ID**: Generated from email address (e.g., `person-sarah-chen-domain-com`)
2. **accountInformation**: REQUIRED - links to Entra ID user via `userPrincipalName` (JSON serialized)
3. **skills**: JSON-serialized `skillProficiency` entities with `displayName` field (for `personSkills` label)
4. **aboutMe**: JSON-serialized `personAnnotation` entity with `detail.contentType` and `detail.content` (for `personNote` label)
5. **Custom columns**: Ignored by the connector (strict-by-doc)
6. **Array properties**: Marked with `@odata.type: Collection(String)`
7. **ACL**: Everyone in the organization can view
8. **No content field**: People data connectors rely on properties with labels, not content field

## Schema Details

The Graph Connector schema includes core properties with people data labels:

### Core Schema (People Data Labels for Copilot)

| Property | Type | Label | Copilot Searchable |
|----------|------|-------|-------------------|
| `accountInformation` | string | `personAccount` | Required for user mapping |
| `skills` | stringCollection | `personSkills` | **Yes** |
| `aboutMe` | string | `personNote` | **Yes** |

### Custom Columns (Ignored)

Extra CSV columns without people data labels are ignored by the connector in strict-by-doc mode.

**Dynamic Schema**: Only people-labeled properties are included; extra columns are ignored.

### Available People Data Labels (Microsoft)

From [Microsoft documentation](https://learn.microsoft.com/en-us/microsoft-365-copilot/extensibility/build-connectors-with-people-data):

| Label | Type | Profile Entity |
|-------|------|----------------|
| `personAccount` | string | userAccountInformation (required) |
| `personSkills` | stringCollection | skillProficiency |
| `personNote` | string | personAnnotation |
| `personCertifications` | stringCollection | personCertification |
| `personAwards` | stringCollection | personAward |
| `personProjects` | stringCollection | projectParticipation |
| `personAddresses` | stringCollection | itemAddress |
| `personEmails` | stringCollection | itemEmail |
| `personPhones` | stringCollection | itemPhone |

**NOT available**: `personLanguages`, `personInterests` (platform limitation - these cannot be made Copilot-searchable via connectors)

## Authentication Flow

### OAuth 2.0 Client Credentials Flow

```
┌─────────────────────────────────────────────────────────────┐
│  enrich-profiles.ts                                         │
│  1. Reads AZURE_CLIENT_SECRET from .env                     │
│  2. Creates ClientSecretCredential                          │
│     - tenantId                                              │
│     - clientId                                              │
│     - clientSecret                                          │
│  3. Acquires access token from Azure AD                     │
│  4. Token contains application permissions:                 │
│     - ExternalConnection.ReadWrite.OwnedBy                  │
│     - ExternalItem.ReadWrite.OwnedBy                        │
└────────────────┬────────────────────────────────────────────┘
                 │
                 │ Access Token (application permissions)
                 ▼
┌─────────────────────────────────────────────────────────────┐
│  Microsoft Graph API (beta endpoint)                        │
│  - Create/manage connections                                │
│  - Register schema with People Data labels                  │
│  - Ingest external items                                    │
│  - All operations use application context (no user)         │
└─────────────────────────────────────────────────────────────┘
```

**Why Client Credentials Flow?**

Graph Connectors operate as **background services** and require application permissions. Delegated permissions (user sign-in) don't work for item ingestion, even with correct scopes.

See [PEOPLE-DATA-AUTH-SOLUTION.md](./PEOPLE-DATA-AUTH-SOLUTION.md) for detailed explanation.

## Code Structure

### Core Files

**`src/enrich-profiles.ts`** (Main entry point)
- CLI argument parsing
- Authentication (client credentials)
- CSV loading and parsing
- Orchestrates connection/schema/ingestion

**`src/people-connector/connection-manager.ts`** (Connection lifecycle)
- `createConnection()` - Creates Graph Connector connection
- `registerSchema()` - Registers schema with properties
- `waitForSchemaReady()` - Polls until schema is ready
- `getConnection()` - Gets connection status

**`src/people-connector/schema-builder.ts`** (Schema generation)
- `buildPeopleSchema()` - Builds schema from CSV columns
- Maps properties to People Data labels
- Adds custom properties dynamically

**`src/people-connector/item-ingester.ts`** (Data ingestion)
- `createExternalItem()` - Converts CSV row to external item
- `ingestItem()` - Ingests single item (beta endpoint)
- `batchIngestItems()` - Ingests multiple items with error handling

**`src/schema/user-property-schema.ts`** (Property definitions)
- Property metadata with `handledBy: 'optionA' | 'optionB'`
- People Data label mappings
- Helper functions: `getOptionBProperties()`, `getPeopleDataMapping()`

### Key Functions

**CSV Parsing**:
```typescript
async loadPeopleData(csvPath: string): Promise<any[]> {
  const content = await fs.readFile(csvPath, 'utf-8');
  const records = parse(content, { columns: true, skip_empty_lines: true, trim: true });

  // Parse array values (skills, interests, etc.)
  return records.map(record => {
    if (record.skills) record.skills = parsePropertyValue('skills', record.skills);
    if (record.interests) record.interests = parsePropertyValue('interests', record.interests);
    // ... more array properties
    return record;
  });
}
```

**External Item Creation**:
```typescript
createExternalItem(csvRow: any, csvColumns: string[]): any {
  const email = csvRow.email;
  const itemId = `person-${email.replace(/@/g, '-').replace(/\./g, '-')}`;

  const properties = {
    accountInformation: JSON.stringify({ userPrincipalName: email })
  };

  // Add Option B properties (skills, interests, etc.)
  // Add custom properties (VTeam, BenefitPlan, etc.)

  return {
    id: itemId,
    content: { value: contentParts.join('. '), type: 'text' },
    properties,
    acl: [{ type: 'everyone', value: 'everyone', accessType: 'grant' }]
  };
}
```

**Item Ingestion** (using beta endpoint):
```typescript
async batchIngestItems(items: any[]): Promise<{...}> {
  for (const item of items) {
    await this.betaClient
      .api(`/external/connections/${this.connectionId}/items/${item.id}`)
      .put(item);

    // Delay to avoid rate limiting
    await new Promise(resolve => setTimeout(resolve, 100));
  }
}
```

## Troubleshooting

### 401 Unauthenticated Errors

**Symptom**: Items fail with "lacks valid authentication credentials"

**Cause**: Using delegated permissions instead of application permissions

**Solution**:
1. Verify client secret is in `.env`
2. Verify Application permissions granted (not Delegated)
3. Verify admin consent granted
4. See [PEOPLE-DATA-AUTH-SOLUTION.md](./PEOPLE-DATA-AUTH-SOLUTION.md)

### Invalid Client Secret

**Symptom**: "AADSTS7000215: Invalid client secret provided"

**Cause**: Using Secret ID instead of Secret Value, or secret expired

**Solution**:
1. Azure Portal > Your App > Certificates & secrets
2. Create NEW client secret
3. Copy the **Value** (appears only once!)
4. Update `.env` with the value

### Schema Stuck in "draft"

**Symptom**: Schema never reaches "ready" state

**Cause**: Schema registration error or timeout

**Solution**:
1. Wait 10 minutes (schema provisioning can be slow)
2. Check Azure Portal > Microsoft Search > Connectors
3. Delete and recreate connection if needed

### Items Not Appearing in Search

**Symptom**: Items ingested successfully but not searchable

**Cause**: Search indexing delay (expect 6+ hours) or using app-only search for `externalItem`

**Solution**:
1. Wait for indexing to complete (6+ hours typical)
2. Verify items exist: `node tools/debug/verify-items.mjs`
3. Check search totals with delegated auth: `node tools/debug/verify-ingestion-progress.mjs --search-auth delegated`
4. Check Microsoft Search admin center for indexing status

### Profile Data Not Appearing in /me/profile API

**Symptom**: Items ingested but `/me/profile` shows empty skills, notes, languages

**Cause**: Profile source not registered or not in prioritized sources

**Solution**:
1. Verify `PeopleSettings.ReadWrite.All` permission is granted with admin consent
2. Run `npm run option-b:setup` to re-register profile source
3. Check output for "✓ Registered as profile source" and "✓ Added to prioritized profile sources"
4. Allow 1-24 hours for profile data propagation

**Technical Details**:
- Profile source registration uses beta API: `POST /admin/people/profileSources`
- Prioritization uses: `PATCH /admin/people/profilePropertySettings/{id}`
- The API returns a collection with `value` array - must extract settings ID to PATCH

### Profile Data Not Searchable by Copilot

**Symptom**: Skills/notes written via Profile API appear on cards but Copilot can't find them. When asking "Find English speakers", Copilot shows people but marks them as "not confirmed".

**Cause**: Data written via Profile API with delegated auth is stored as `source.type: "User"` with `isSearchable: false`. Only data from system sources (connectors with people data labels) is Copilot-searchable.

**Solution** (implemented in Option B connector enrichment):
1. Write skills via Graph Connector with `personSkills` label
2. Write aboutMe via Graph Connector with `personNote` label
3. Delete and recreate connector (schema changes require this)
4. Wait 6+ hours for indexing
5. Test with Copilot: "Find people with skills in [skill name]"

**Technical Details**:
- Skills must be JSON-serialized `skillProficiency` entities: `{"displayName": "TypeScript"}`
- AboutMe must be JSON-serialized `personAnnotation` entity: `{"detail": {"contentType": "text", "content": "..."}}`
- Languages have NO connector label - platform limitation, cannot be made Copilot-searchable

### Profile Source Registration Failed

**Symptom**: "Resource not found for the segment 'profileSources'"

**Cause**: Using v1.0 API instead of beta, or missing permission

**Solution**:
1. Ensure using beta endpoint (code handles this automatically)
2. Verify `PeopleSettings.ReadWrite.All` application permission
3. Grant admin consent in Azure Portal
4. People data connectors are in preview - tenant may need opt-in

### Skills Empty Despite Successful Ingestion (profileSyncEnabled=False)

**Symptom**: Items ingested with `CAPIv2 export completed with status 'Created'`, but all exports show `profileSyncEnabled=False, profileSynced=False`. User profiles show `skills: []`.

**Root Cause**: The connection is NOT in the `prioritizedSourceUrls` list. This can happen if:

1. **Stale source cleanup removed the AAD default source**: The `prioritizedSourceUrls` array contains Microsoft's internal AAD source (UUID format like `4ce763dd-...`). If cleanup code tries to validate this by calling `GET /external/connections/{uuid}`, it gets 404 and removes it. Microsoft then silently reverts the entire PATCH, so our connector URL never gets added either.

2. **Profile source not propagated before ingestion**: Items ingested before TSS (Tenant Settings Service) propagates the profile source get `"Profile source registration failed for item"` in internal logs. Even though CAPIv2 export still happens, `profileSyncEnabled` stays `False`.

**Diagnosis**:
```bash
node tools/debug/check-profile-source.mjs <connection-id>
# Look for: "❌ Our connection is NOT in prioritized sources"
```

**Internal debug log indicators** (Admin Portal CSV export):
- `ProfileSourceRegistrar: Failed to retrieve settings from TSS, statusCode=Unauthorized` — profile source not yet propagated
- `CAPIv2 export completed ... profileSyncEnabled=False` — data will NOT sync to profiles
- `CAPIv2 export completed ... profileSyncEnabled=True` — data WILL sync to profiles (this is what you want)

**Solution**:
1. Fix prioritization: `node tools/admin/register-profile-source.mjs <connection-id>`
2. Re-ingest items: `npm run option-b:ingest -- --csv config/textcraft-europe.csv --connection-id <id>`
3. Wait 6+ hours for indexing

**Prevention** (implemented in code):
- Stale source cleanup now only validates alphanumeric connector IDs, preserving UUID-format internal sources
- 60-second propagation wait after profile source registration before ingestion starts

### Prioritized Sources — Stale Cleanup Destroying AAD Default

**Symptom**: After running `option-b:setup`, the diagnostic shows the connection is NOT in prioritized sources, even though the code logged `"Added to prioritized profile sources"`.

**Root Cause**: The stale source cleanup validated ALL source URLs by calling `GET /external/connections/{sourceId}`. The AAD default source (`4ce763dd-...`) is not an external connection, so it returned 404 and was removed. Microsoft then rejected or reverted the PATCH (you can't remove the AAD default source), and our connection was never added.

**Fix**: The cleanup code now uses `CONNECTOR_ID_PATTERN = /^[A-Za-z][A-Za-z0-9]*$/` to only validate external connector IDs. UUID-format sources are always preserved.

### Profile Source Missing `kind` Property

**Symptom**: Profile source appears to be registered, but data doesn't appear in profile cards or Copilot. Debug scripts may show "Profile source exists but missing kind property".

**Cause**: The `kind: 'Connector'` property was not included in the profile source registration payload.

**Solution**:
1. Delete and recreate the profile source with the correct payload:
   ```typescript
   const profileSourcePayload = {
     sourceId: connectionId,
     displayName: 'Your Display Name',
     kind: 'Connector',  // <-- REQUIRED
     webUrl: 'https://...',
   };
   ```
2. Alternatively, run `npm run option-b:setup` which now includes the `kind` property
3. See `docs/COPILOT-CONNECTORS-PEOPLE-DATA.md` Section 11 for details

### Labels Display as `unknownFutureValue`

**Symptom**: When querying schema via API, labels like `personAccount`, `personSkills`, `personNote` display as `unknownFutureValue`.

**Cause**: This is normal behavior when using beta API features. The Microsoft internal systems still recognize the labels correctly.

**Verification**:
1. Check that connection was created with `contentCategory: 'people'` (via beta API)
2. Check that schema was registered via beta API
3. Verify items are being ingested successfully
4. Wait 6+ hours for indexing, then test with Copilot

**If labels truly aren't working**:
1. Ensure you used beta API for connection creation (not v1.0)
2. Delete and recreate the connection with beta API
3. The v1.0 API ignores `contentCategory: 'people'` and sets it to `uncategorized`

### CSV Parsing Errors

**Symptom**: "Cannot parse array value" or similar

**Cause**: Invalid array format in CSV

**Solution**:
Arrays in CSV should be formatted as:
- JSON: `["value1","value2"]`
- Comma-separated: `value1,value2`
- Both work - parser handles both formats

## Verification

After ingestion, verify the data and indexing progress:

```bash
# Compare CSV totals vs indexed items (delegated auth)
npm run build
node tools/debug/verify-ingestion-progress.mjs \
  --search-auth delegated \
  --connection-id m365people \
  --csv config/textcraft-europe.csv \
  --query "*"
```

**Expected Output**:
```
✅ person-sarah-chen-domain-com
   Properties:
     - Email: sarah@domain.com
     - Skills: 3 items
     - Interests: 2 items
     - About: Experienced executive...
```

Notes:
- Search for `externalItem` requires delegated auth (default scope: `ExternalItem.Read.All`).
- The script falls back to `connection.itemCount` if search is unavailable.
- Use `--sample-email` to fetch a specific external item and verify properties/ACL.

**Manual Verification**:
1. **Microsoft Search**: Search for enrichment data (e.g., "skills:TypeScript")
2. **Profile Cards**: Open Teams, view user profile, check for enrichment data
3. **Copilot**: Ask Copilot about team members' skills or expertise

## CLI Reference

| Command | Description |
|---------|-------------|
| `npm run option-a:enrich` | Profile API enrichment (languages, interests) |
| `npm run option-a:enrich:dry-run` | Preview Profile API enrichment |
| `npm run option-b:setup` | Create connection, register schema, and ingest (first time) |
| `npm run option-b:ingest` | Ingest people data from default CSV |
| `npm run option-b:ingest -- --csv path/to/file.csv` | Ingest from specific CSV |
| `npm run option-b:dry-run` | Preview external items without ingesting |
| `npm run option-b:ingest -- --connection-id custom-id` | Use custom connection ID |
| `node tools/debug/verify-ingestion-progress.mjs --search-auth delegated ...` | Monitor indexing progress |

## Automatic Deletion (Orphan Cleanup)

### How It Works

Option B automatically tracks and cleans up orphaned external items:

**State Tracking**:
- Maintains state file: `state/external-items-state.json`
- Tracks all created external items by ID
- Updates after each ingestion

**Orphan Detection**:
- Compares previous state with current CSV
- Identifies items that exist in state but not in CSV
- Automatically deletes orphaned items

**When Deletion Occurs**:
- User removed from CSV
- User deleted in Option A
- Email address changed

**Example**:
```bash
# Previous CSV had 5 users, now has 3 users (2 removed)
npm run option-b:ingest -- --csv config/agents-template.csv

# Output:
🔍 Checking for orphaned external items...
🗑️  Deleted orphaned item: person-john-doe-domain-com
🗑️  Deleted orphaned item: person-jane-smith-domain-com
✓ Deleted 2 orphaned items
```

**State File Example** (`state/external-items-state.json`):
```json
{
  "items": [
    "person-ingrid-johansen-a830edad9050849coep9vqp9bog-onmicrosoft-com",
    "person-ola-nordmann-a830edad9050849coep9vqp9bog-onmicrosoft-com",
    "person-lars-hansen-a830edad9050849coep9vqp9bog-onmicrosoft-com"
  ],
  "lastUpdated": "2026-01-22T18:23:10.821Z"
}
```

## Maintenance

### Updating Data

To update enrichment data for existing users:

1. Update CSV file with new data
2. Run `npm run option-b:ingest` again
3. External items are **replaced** (PUT operation)
4. Orphaned items automatically deleted

### Adding New Users

Option B automatically handles new users:

1. Run Option A to create new users in Entra ID
2. Add new users to CSV
3. Run Option B - new items created, state updated

### Removing Users

Option B automatically handles user removal:

1. Remove users from CSV
2. Run Option B - orphaned items automatically deleted
3. State file updated to reflect current users

### Deleting Connection

To start fresh:

```bash
node delete-connection.mjs
```

Then run setup again:

```bash
npm run option-b:setup
```

## Performance

- **Batch Size**: 20 items (configured in code)
- **Rate Limiting**: 100ms delay between items
- **Throughput**: ~10 items/second
- **100 users**: ~10 seconds
- **1000 users**: ~2 minutes

## Security Best Practices

1. **Never commit `.env`**: Add to `.gitignore`
2. **Rotate secrets**: Set expiration, create new before expiry
3. **Least privilege**: Use `.OwnedBy` permissions, not `.All`
4. **Monitor access**: Review Azure AD sign-in logs regularly
5. **Consider certificates**: More secure than client secrets for production

## Benefits

✅ **Single enrichment system**: All non-standard data in one place
✅ **Copilot integration**: Data surfaces in AI responses
✅ **Profile cards**: Enrichment appears in M365 profile cards
✅ **Microsoft Search**: People-labeled connector data is searchable (after indexing delay)
✅ **Strict schema**: Only people-labeled properties ingested (extra columns ignored)
✅ **Same CSV file**: No duplicate data entry
✅ **Batch operations**: Efficient ingestion

---

**Last Updated**: 2026-02-16
**Status**: ✅ Production Ready (with Copilot searchability via people data labels)
**Authentication**: Option A: Delegated (Profile API), Option B: Client Credentials (Connectors)
**Main File**: `src/enrich-connector.ts` (Option B), `src/enrich-profiles.ts` (Option A enrichment)
**Sample Data**: `config/textcraft-europe.csv` (95 users)

## Copilot Searchability Summary

| Data Type | Method | Copilot Searchable |
|-----------|--------|-------------------|
| Skills | Connector (`personSkills` label) | **Yes** |
| Notes/AboutMe | Connector (`personNote` label) | **Yes** |
| Custom props (VTeam, etc.) | Connector (no label) | **Yes** |
| Languages | Profile API only | **No** (no label available) |
| Interests | Profile API only | **No** (no label available) |
