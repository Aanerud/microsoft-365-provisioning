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
    └─► Option B: Graph Connector Pipeline
        - skills (personSkills), aboutMe (personNote)
        - certifications, awards, projects, etc.
        - Auth: Application (client secret)
        - File: src/enrich-connector.ts
        - API: PUT /external/connections/{id}/items/{id}
```

### Data Routing

| Data Type | Script | Why |
|-----------|--------|-----|
| **Standard User Fields** | Option A (`provision.ts`) | Native identity management |
| **Profile Fields** (languages, interests) | Option A (`enrich-profiles.ts`) | No people data labels exist for Copilot |
| **Profile Fields** (skills, aboutMe) | Option B (`enrich-connector.ts`) | Have people data labels (personSkills, personNote) for Copilot search |
| **Custom Org Fields** (VTeam, CostCenter) | Option B (`enrich-connector.ts`) | Only way to add custom searchable properties |

### Graph Connectors for Copilot Searchability

Graph Connectors with **people data labels** enable Copilot to search for people based on their skills, notes, and custom properties. Key requirements:
- Connection must have `contentCategory: 'people'`
- Schema must use `personAccount` label (required for user mapping)
- Schema can use `personSkills`, `personNote`, `personCertifications`, etc.
- Items must have ACL set to `everyone`/`grant`
- Wait 6+ hours after ingestion for Microsoft 365 to index

See `docs/COPILOT-CONNECTORS-PEOPLE-DATA.md` for complete details.

#### Critical Requirements

1. **Profile Source Registration - `kind` Property**

   When registering a connector as a profile source, the `kind: 'Connector'` property is **REQUIRED**:
   ```typescript
   // CORRECT - kind is required
   const profileSourcePayload = {
     sourceId: connectionId,
     displayName: 'Display Name',
     kind: 'Connector',  // <-- REQUIRED per Microsoft docs
     webUrl: 'https://...',
   };
   ```
   Without this property, profile data may not appear in profile cards or Copilot.

2. **Beta API for Connection Creation**

   The `contentCategory: 'people'` parameter is only supported by the beta API. Using v1.0 will result in `contentCategory: 'uncategorized'`.

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

**Last Updated**: 2026-02-06
**Version**: 4.0.0 (Option A/B split: standalone Profile API enrichment + Graph Connector pipeline)
