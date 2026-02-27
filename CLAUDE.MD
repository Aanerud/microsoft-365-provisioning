# M365 Agent Provisioning - Claude Code Documentation

## Project Overview

**M365-Agent-Provisioning** is a TypeScript CLI tool for bulk provisioning Microsoft 365 users with automatic license assignment and profile enrichment.

### Key Capabilities

- **Option A - User Provisioning**: Create M365 users with standard properties
- **Profile API Enrichment**: Write languages, interests directly to user profiles (no people data labels)
- **Graph Connectors**: Add people-labeled properties (skills, aboutMe, certifications, etc.) for Copilot search
- **License Management**: Automatically assign Microsoft 365 licenses
- **State Management**: CREATE/UPDATE/NOOP detection to avoid duplicate operations

## Architecture

### Option A / Option B Architecture

The tool has two independent pipelines for enrichment:

```
CSV Data
    │
    ├─► Option A: User Provisioning + Profile API Enrichment
    │   - Creates Entra ID accounts (src/provision.ts)
    │   - Sets standard properties (name, email, jobTitle, etc.)
    │   - Assigns M365 licenses
    │   - Profile API: languages, interests (src/enrich-profiles.ts)
    │   - Auth: Delegated (browser login)
    │
    └─► Option B: Graph Connector Pipeline (all 13 people data labels)
        - Native: skills, aboutMe, certs, projects, awards, birthday, mySite
        - Composite: name, position, addresses, emails, phones
        - Custom: VTeam (string only)
        - Auth: Application (client secret)
        - File: src/enrich-connector.ts
        - API: PUT /external/connections/{id}/items/{id}
```

### Data Routing

| Data Type | Script | Why |
|-----------|--------|-----|
| **Standard User Fields** | Option A (`provision.ts`) | Native identity management |
| **Profile Fields** (languages, interests) | Option A (`enrich-profiles.ts`) | No people data labels exist for Copilot |
| **Option B Native** (skills, aboutMe, certs, projects, awards, birthday, mySite) | Option B (`enrich-connector.ts`) | Have people data labels for Copilot search |
| **Composite Labels** (name, position, addresses, emails, phones) | Option B (`enrich-connector.ts`) | Composed from Option A CSV columns, sent via connector for Copilot discoverability |
| **Custom Org Fields** (VTeam) | Option B (`enrich-connector.ts`) | Custom searchable properties (string only, Path B) |

### Graph Connectors for Copilot Searchability

Graph Connectors with **people data labels** enable Copilot to search for people based on their skills, notes, and custom properties. All 13 official labels are enabled and validated (m365people23). Key requirements:
- Connection must have `contentCategory: 'people'`
- Schema must use `personAccount` label (required for user mapping)
- All 13 people data labels enabled (see full list below)
- Items must have ACL set to `everyone`/`grant`
- Wait 6–24 hours after ingestion for Microsoft 365 to index

See `docs/COPILOT-CONNECTORS-PEOPLE-DATA.md` for end-to-end flow.
See `docs/graph-connector-lessons.md` for **critical lessons learned** from m365people14–23.

#### Path A vs Path B Deserialization (CRITICAL)

Confirmed by Microsoft engineer — the connector backend has two deserialization paths:

- **Path A (with label)**: Stores values as `JsonElement` — safely handles any JSON type
- **Path B (without label)**: Blob deserialization — expects `string` values ONLY

**Rule**: Custom properties (no label) MUST be type `string`. Sending a `stringCollection` without a label crashes the entire connection with:
```
Cannot get the value of a token type 'StartArray' as a string. Path: $.skills
```

#### All 13 People Data Labels (validated m365people23)

**Group 1 — Option B native** (single CSV column → entity):

| Label | Type | CSV Column | Serialization |
|-------|------|-----------|---------------|
| `personSkills` | stringCollection | skills | `{"displayName":"..."}` |
| `personNote` | string | aboutMe | `{"detail":{"contentType":"text","content":"..."}}` |
| `personCertifications` | stringCollection | certifications | `{"displayName":"..."}` |
| `personProjects` | stringCollection | projects | `{"displayName":"..."}` |
| `personAwards` | stringCollection | awards | `{"displayName":"..."}` |
| `personAnniversaries` | stringCollection | birthday | `{"type":"birthday","date":"..."}` |
| `personWebSite` | string | mySite | `{"webUrl":"..."}` |

**Group 2 — Composite** (multiple Option A CSV columns → entity):

| Label | Type | CSV Columns | Serialization |
|-------|------|------------|---------------|
| `personName` | string | givenName, surname, displayName | `{"displayName":"...","first":"...","last":"..."}` |
| `personCurrentPosition` | string | jobTitle, companyName | `{"detail":{"jobTitle":"...","company":{"displayName":"..."}},"isCurrent":true}` |
| `personAddresses` | stringCollection | streetAddress, city, state, country, postalCode | `{"type":"business","street":"...","city":"..."}` |
| `personEmails` | stringCollection | mail | `{"address":"...","type":"main"}` |
| `personPhones` | stringCollection | mobilePhone, businessPhones | `{"number":"...","type":"mobile"/"business"}` |
| `personWebAccounts` | stringCollection | *(none)* | No CSV data — declared but empty |

#### Code Change Discipline

**NEVER refactor working connector code.** m365people18–20 failed because of "harmless" refactoring (changing type definitions, removing unused methods) that produced identical item payloads. Only make additive changes to the working baseline. See `docs/graph-connector-lessons.md` for full history.

#### Other Critical Requirements

1. **Profile Source Registration - `kind` Property**

   The `kind: 'Connector'` property is **REQUIRED** when registering a profile source.

2. **Beta API for Connection Creation**

   The `contentCategory: 'people'` parameter is only supported by the beta API.

3. **Profile Source Propagation — Separate Registration from Ingestion**

   `connection-manager.ts` waits 60 seconds after registration and polls `GET /admin/people/profileSources` to confirm propagation before ingestion begins.

4. **Prioritized Source Cleanup — NEVER Remove Non-Connector Sources**

   UUID-format sourceIds (e.g. `4ce763dd-...`) are internal Microsoft sources — never remove them from `prioritizedSourceUrls`.

5. **Delete Old Connections Before Deploying New Ones**

   Multiple active connections with overlapping labels cause priority conflicts. Always delete the previous connection first.

6. **`profileSyncEnabled` Flag**

   `profileSyncEnabled=False` on initial CAPIv2 export is NORMAL. Data appears on profiles after subsequent cycles (6–24h). True failure is when it stays False across multiple cycles with TSS Unauthorized errors.

### Profile API Collections (20 Total)

These can be written directly via `POST /users/{id}/profile/{collection}`:

| Collection | Endpoint | Example Data |
|------------|----------|--------------|
| `skills` | `/profile/skills` | "TypeScript", "Project Management" |
| `languages` | `/profile/languages` | "French (Native)", "English (Fluent)" |
| `notes` | `/profile/notes` | About Me text |
| `interests` | `/profile/interests` | "Machine Learning", "Photography" |
| `certifications` | `/profile/certifications` | "AWS Solutions Architect" |
| `awards` | `/profile/awards` | "Employee of the Year" |
| `positions` | `/profile/positions` | Work history |
| `projects` | `/profile/projects` | Project participation |
| `educationalActivities` | `/profile/educationalActivities` | University degrees |
| `patents` | `/profile/patents` | Patents held |
| `publications` | `/profile/publications` | Books, articles |
| `responsibilities` | `/profile/responsibilities` | Job duties |
| `addresses` | `/profile/addresses` | Physical addresses |
| `phones` | `/profile/phones` | Phone numbers |
| `emails` | `/profile/emails` | Email addresses |
| `websites` | `/profile/websites` | Personal websites |
| `webAccounts` | `/profile/webAccounts` | LinkedIn, GitHub |
| `anniversaries` | `/profile/anniversaries` | Birthday, work anniversary |
| `names` | `/profile/names` | Name variations |
| `accounts` | `/profile/accounts` | Account information |

**Note**: Profile API is beta-only and requires delegated authentication.

## Authentication

### Delegated Auth (Browser Login)
Used for: Option A (provisioning) and Profile API (enrichment)

```
1. CLI starts local HTTP server on port 5544
2. Browser opens Azure AD login page
3. User authenticates with admin account
4. Tokens cached in ~/.m365-provision/token-cache.json
```

### Application Auth (Client Secret)
Used for: Graph Connectors (custom properties)

```
1. CLI reads AZURE_CLIENT_SECRET from .env
2. Uses @azure/identity ClientSecretCredential
3. Token used for connector operations
```

## Environment Variables

```bash
# Azure AD Configuration
AZURE_TENANT_ID=your-tenant-id
AZURE_CLIENT_ID=your-client-id
AZURE_CLIENT_SECRET=your-client-secret  # Required for Graph Connectors

# Provisioning Configuration
LICENSE_SKU_IDS=sku-id-1,sku-id-2
USER_DOMAIN=yourdomain.onmicrosoft.com

# Auth Server
AUTH_SERVER_PORT=5544
```

## Azure AD Permissions

### Delegated Permissions (Profile API + Option A)
- `User.ReadWrite.All` - Create users and write to profiles
- `People.Read.All` - Read profile data
- `Directory.ReadWrite.All` - Full directory access
- `offline_access` - Refresh tokens

### Application Permissions (Graph Connectors)
- `ExternalConnection.ReadWrite.OwnedBy` - Manage owned connections
- `ExternalItem.ReadWrite.OwnedBy` - Manage items in owned connections
- `PeopleSettings.ReadWrite.All` - Profile source registration

## npm Scripts

### Option A: Provisioning + Profile API Enrichment
```bash
npm run provision                # Create users from CSV
npm run provision:beta           # Use beta Graph API
npm run option-a:enrich          # Profile API enrichment (languages, interests)
npm run option-a:enrich:dry-run  # Preview without making changes
npm run list-users               # List provisioned users
npm run update-licenses          # Add missing licenses
```

### Option B: Graph Connector Pipeline
```bash
npm run option-b:setup           # Create connection + schema + ingest
npm run option-b:ingest          # Ingest items (connection must exist)
npm run option-b:dry-run         # Preview without making changes
npm run enrich:wait              # Wait for schema provisioning
npm run enrich:delete-connector  # Delete connector connection
npm run build-oid-cache          # Build OID cache from Graph
```

### Utilities
```bash
npm run logout              # Clear cached tokens
npm run test-connection     # Test Graph API
npm run reset-tenant        # Preview tenant reset
npm run reset-tenant:confirm # Delete all non-admin users
```

## Project Structure

```
M365-Agent-Provisioning/
├── src/
│   ├── provision.ts              # Option A - User provisioning
│   ├── enrich-profiles.ts        # Option A - Profile API enrichment (languages, interests)
│   ├── enrich-connector.ts       # Option B - Graph Connector pipeline (skills, aboutMe, etc.)
│   ├── profile-writer.ts         # Profile API writer class
│   ├── graph-client.ts           # Microsoft Graph client
│   ├── oid-cache.ts              # OID cache (UPN → object ID mapping)
│   ├── auth/
│   │   ├── browser-auth-server.ts
│   │   └── token-cache.ts
│   ├── people-connector/
│   │   ├── connection-manager.ts
│   │   ├── schema-builder.ts
│   │   └── item-ingester.ts
│   └── schema/
│       └── user-property-schema.ts
├── config/
│   └── textcraft-europe.csv      # Sample data
├── tools/
│   ├── debug/                    # Debug scripts
│   └── admin/                    # Admin utilities
└── docs/                         # Documentation
```

## Key Files

| File | Purpose |
|------|---------|
| `src/provision.ts` | User provisioning (Option A) |
| `src/enrich-profiles.ts` | Profile API enrichment - languages, interests (Option A) |
| `src/enrich-connector.ts` | Graph Connector pipeline - skills, aboutMe, etc. (Option B) |
| `src/profile-writer.ts` | Profile API writer (languages, interests, etc.) |
| `src/oid-cache.ts` | OID cache (UPN → object ID mapping) |
| `src/people-connector/connection-manager.ts` | Graph Connector connection + profile source registration |
| `src/people-connector/item-ingester.ts` | Graph Connector item ingestion |
| `docs/COPILOT-CONNECTORS-PEOPLE-DATA.md` | **Deep understanding of people data connectors** |

## Common Workflows

### First-Time Setup

```bash
# 1. Configure environment
cp .env.example .env
# Edit .env with credentials

# 2. Create users (Option A)
npm run provision -- --csv config/users.csv

# 3. Enrich via Profile API (Option A - languages, interests)
npm run option-a:enrich -- --csv config/users.csv

# 4. Setup Graph Connector + ingest (Option B - skills, aboutMe, etc.)
npm run option-b:setup -- --csv config/users.csv --connection-id m365people3
```

### Regular Enrichment

```bash
# Update Profile API data (languages, interests)
npm run option-a:enrich -- --csv config/users.csv

# Re-ingest connector data (skills, aboutMe, etc.)
npm run option-b:ingest -- --csv config/users.csv --connection-id m365people3
```

## Known Issues

### Profile API Skills - Proficiency Bug
Microsoft Graph API has a bug where certain `proficiency` values fail:
- **Working**: `elementary`, `limitedWorking`, `expert`
- **Failing**: `generalProfessional`, `advancedProfessional`

**Workaround**: The ProfileWriter omits proficiency or uses working values only.

### Profile API Notes - Schema Format
The `detail` property must be an `itemBody` object:
```json
{
  "detail": {
    "contentType": "text",
    "content": "About me text..."
  }
}
```

### Irish Language
The Profile API doesn't accept "Irish" as a language name. Needs valid BCP-47 tag.

## Security

- Never commit `.env` files
- Tokens cached in `~/.m365-provision/` with restricted permissions
- Protected accounts (admin@*, etc.) are never deleted during reset

---

**Last Updated**: 2026-02-27
**Version**: 5.0.0 (All 13 people data labels enabled, validated with m365people23)
