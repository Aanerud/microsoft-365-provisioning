# Option B: Profile Enrichment via Microsoft Graph Connectors

**Status**: ✅ Implemented and Working (2026-01-22)

## Overview

Option B enriches Microsoft 365 user profiles by ingesting additional data through **Microsoft Graph Connectors with People Data labels**. This surfaces enrichment data in:

- Microsoft 365 profile cards
- Microsoft Search
- Microsoft 365 Copilot responses
- Teams member discovery

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
│ - NO enrichment data (handled by Option B)                 │
│                                                             │
│ Output: Real Entra ID users with standard properties       │
└─────────────────┬───────────────────────────────────────────┘
                  │
                  │ Users exist in Entra ID (prerequisite)
                  ↓
┌────────────────────────────────────────────────────────────┐
│ Option B: Profile Enrichment                               │
│ File: src/enrich-profiles.ts                               │
│                                                             │
│ - Creates Graph Connector connection (once)                │
│ - Registers schema with People Data labels (once)          │
│ - Ingests enrichment data for all users:                   │
│   • Official People Data (skills, certifications, etc.)    │
│   • Custom properties (VTeam, BenefitPlan, etc.)           │
│ - Links items to Entra ID users                            │
│ - Automatic deletion of orphaned items (state-based)       │
│ - Uses OAuth 2.0 Client Credentials Flow (app-only)        │
│                                                             │
│ Output: External items in Microsoft Search                 │
│         Surfaces in profile cards & Copilot                │
│         State file tracking for automatic cleanup          │
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

### Enrichment Properties (Option B)
Handled by `enrich-profiles.ts` - stored in Graph Connector:

**Official People Data** (with Microsoft labels):
- `skills` → `personSkills`
- `pastProjects` → `personProjects`
- `certifications` → `personCertifications`
- `awards` → `personAwards`
- `aboutMe` → `personNote`
- `mySite` → `personWebSite`
- `birthday` → `personAnniversaries`

**Custom People Data** (searchable, no official labels):
- `interests`
- `responsibilities`
- `schools`

**Custom Organization Properties** (searchable):
- `VTeam`
- `BenefitPlan`
- `CostCenter`
- `BuildingAccess`
- `ProjectCode`
- *(Any additional columns in CSV not in schema)*

## Prerequisites

### 1. Option A Must Run First

Users must exist in Entra ID before enriching their profiles. Option B links external items to Entra ID users via `userPrincipalName`.

```bash
# First: Create users
npm run provision -- --csv config/agents-template.csv

# Then: Enrich profiles
npm run enrich-profiles -- --csv config/agents-template.csv
```

### 2. Azure AD Configuration

**App Registration Requirements**:

**Application Permissions** (required):
- `ExternalConnection.ReadWrite.OwnedBy`
- `ExternalItem.ReadWrite.OwnedBy`

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
npm run enrich-profiles:setup
```

**What this does**:
1. Creates Graph Connector connection (`m365provisionpeople`)
2. Registers schema with 16 properties:
   - 1 required: `accountInformation` (links to Entra ID user)
   - 7 with official People Data labels
   - 8 custom searchable properties
3. Waits for schema to be ready (can take 5-10 minutes)

**Output**:
```
✓ Created connection: m365provisionpeople
✓ Schema registration initiated
  Schema state: draft, waiting...
  Schema state: draft, waiting...
  Schema state: ready
✓ Schema is ready
```

### Step 2: Ingest People Data

Ingest enrichment data from your CSV file:

```bash
npm run enrich-profiles -- --csv config/agents-template.csv
```

**What this does**:
1. Authenticates with OAuth 2.0 Client Credentials Flow
2. Loads CSV and parses enrichment properties
3. Creates external items for each person
4. Ingests items to Graph Connector (using beta endpoint)
5. Links items to Entra ID users via email address
6. Automatically detects and deletes orphaned items (items not in CSV)

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
npm run enrich-profiles:dry-run
```

**Output**: Shows sample external item structure with all properties.

## External Item Structure

Each person is converted to an external item with this structure:

```json
{
  "id": "person-email-domain-com",
  "content": {
    "value": "About me text. Skills: TypeScript, Python, Azure. Interests: Open Source, Cloud. VTeam: Platform Team",
    "type": "text"
  },
  "properties": {
    "accountInformation": "{\"userPrincipalName\":\"email@domain.com\"}",
    "skills@odata.type": "Collection(String)",
    "skills": [
      "{\"displayName\":\"TypeScript\"}",
      "{\"displayName\":\"Python\"}",
      "{\"displayName\":\"Azure\"}"
    ],
    "interests@odata.type": "Collection(String)",
    "interests": ["Open Source", "Cloud"],
    "aboutMe": "{\"displayName\":\"Experienced engineer...\",\"detail\":{\"value\":\"Experienced engineer...\"}}",
    "VTeam": "Platform Team",
    "BenefitPlan": "Premium Plus",
    "CostCenter": "ENG-DEPT"
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
2. **accountInformation**: REQUIRED - links to Entra ID user via `userPrincipalName`
3. **Labeled properties**: JSON-serialized with `displayName` field (required by People Data)
4. **Custom properties**: Stored as plain strings or arrays
5. **Array properties**: Marked with `@odata.type: Collection(String)`
6. **Content**: Searchable text combining all enrichment data
7. **ACL**: Everyone in the organization can view

## Schema Details

The schema includes 16 properties:

| Property | Type | Label | Category |
|----------|------|-------|----------|
| accountInformation | string | personAccount | Required (link to Entra ID) |
| skills | stringCollection | personSkills | Official People Data |
| pastProjects | stringCollection | personProjects | Official People Data |
| certifications | stringCollection | personCertifications | Official People Data |
| awards | stringCollection | personAwards | Official People Data |
| aboutMe | string | personNote | Official People Data |
| mySite | string | personWebSite | Official People Data |
| birthday | string | personAnniversaries | Official People Data |
| interests | stringCollection | *(none)* | Custom searchable |
| responsibilities | stringCollection | *(none)* | Custom searchable |
| schools | stringCollection | *(none)* | Custom searchable |
| VTeam | string | *(none)* | Custom organization property |
| BenefitPlan | string | *(none)* | Custom organization property |
| CostCenter | string | *(none)* | Custom organization property |
| BuildingAccess | string | *(none)* | Custom organization property |
| ProjectCode | string | *(none)* | Custom organization property |

**Dynamic Schema**: If your CSV contains additional columns not in the standard schema, they're automatically added as custom searchable properties.

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

**Cause**: Search indexing delay (can take 1-2 hours)

**Solution**:
1. Wait for indexing to complete
2. Verify items exist: `node verify-items.mjs`
3. Check Microsoft Search admin center for indexing status

### CSV Parsing Errors

**Symptom**: "Cannot parse array value" or similar

**Cause**: Invalid array format in CSV

**Solution**:
Arrays in CSV should be formatted as:
- JSON: `["value1","value2"]`
- Comma-separated: `value1,value2`
- Both work - parser handles both formats

## Verification

After ingestion, verify the data:

```bash
# Check ingested items
node verify-items.mjs
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

**Manual Verification**:
1. **Microsoft Search**: Search for enrichment data (e.g., "skills:TypeScript")
2. **Profile Cards**: Open Teams, view user profile, check for enrichment data
3. **Copilot**: Ask Copilot about team members' skills or expertise

## CLI Reference

| Command | Description |
|---------|-------------|
| `npm run enrich-profiles:setup` | Create connection and register schema (first time only) |
| `npm run enrich-profiles` | Ingest people data from default CSV |
| `npm run enrich-profiles -- --csv path/to/file.csv` | Ingest from specific CSV |
| `npm run enrich-profiles:dry-run` | Preview external items without ingesting |
| `npm run enrich-profiles -- --connection-id custom-id` | Use custom connection ID |

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
npm run enrich-profiles -- --csv config/agents-template.csv

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
2. Run `npm run enrich-profiles` again
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
npm run enrich-profiles:setup
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
✅ **Microsoft Search**: All data is searchable
✅ **Flexible schema**: Add custom properties without code changes
✅ **Same CSV file**: No duplicate data entry
✅ **Batch operations**: Efficient ingestion

---

**Last Updated**: 2026-01-22
**Status**: ✅ Production Ready
**Authentication**: OAuth 2.0 Client Credentials Flow
