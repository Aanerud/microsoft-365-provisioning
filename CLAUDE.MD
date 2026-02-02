# M365 Agent Provisioning - Claude Code Documentation

## Project Overview

**M365-Agent-Provisioning** is a TypeScript CLI tool for bulk provisioning Microsoft 365 users with automatic license assignment and profile enrichment.

### Key Capabilities

- **Option A - User Provisioning**: Create M365 users with standard properties
- **Profile API Enrichment**: Write skills, languages, notes, interests directly to user profiles
- **Graph Connectors**: Add custom organizational properties (VTeam, CostCenter, etc.)
- **License Management**: Automatically assign Microsoft 365 licenses
- **State Management**: CREATE/UPDATE/NOOP detection to avoid duplicate operations

## Architecture

### Hybrid Enrichment Approach

The tool uses a **hybrid approach** for profile enrichment, using the optimal method for each data type:

```
CSV Data
    │
    ├─► Option A: User Provisioning
    │   - Creates Entra ID accounts
    │   - Sets standard properties (name, email, jobTitle, etc.)
    │   - Assigns M365 licenses
    │   - Auth: Delegated (browser login)
    │   - API: POST /users, PATCH /users/{id}
    │
    ├─► Profile API (Direct POST)
    │   - skills → POST /users/{id}/profile/skills
    │   - languages → POST /users/{id}/profile/languages
    │   - aboutMe → POST /users/{id}/profile/notes
    │   - interests → POST /users/{id}/profile/interests
    │   - certifications, awards, projects, etc.
    │   - Auth: Delegated (browser login)
    │   - API: POST /users/{id}/profile/{collection}
    │
    └─► Graph Connectors (Custom Properties)
        - VTeam, CostCenter, BenefitPlan, BuildingAccess, etc.
        - Any custom organizational property
        - Auth: Application (client secret)
        - API: PUT /external/connections/{id}/items/{id}
```

### Why Hybrid?

| Data Type | Method | Why |
|-----------|--------|-----|
| **Standard User Fields** | Option A (Entra ID) | Native identity management |
| **Profile Fields** (languages, interests) | Profile API only | No people data labels exist for Copilot |
| **Profile Fields** (skills, aboutMe) | Profile API + Connectors | Redundant for both profile cards AND Copilot search |
| **Custom Org Fields** (VTeam, CostCenter) | Graph Connectors | Only way to add custom searchable properties |

### Graph Connectors for Copilot Searchability

Graph Connectors with **people data labels** enable Copilot to search for people based on their skills, notes, and custom properties. Key requirements:
- Connection must have `contentCategory: 'people'`
- Schema must use `personAccount` label (required for user mapping)
- Schema can use `personSkills`, `personNote`, `personCertifications`, etc.
- Items must have ACL set to `everyone`/`grant`
- Wait 6+ hours after ingestion for Microsoft 365 to index

See `docs/COPILOT-CONNECTORS-PEOPLE-DATA.md` for complete details.

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

### Hybrid Enrichment (Recommended)
```bash
npm run enrich              # Run both Profile API and Graph Connectors
npm run enrich:setup        # Setup Graph Connector + run enrichment
npm run enrich:dry-run      # Preview without making changes
npm run enrich:profile-api-only  # Only Profile API (skills, languages, etc.)
npm run enrich:connector-only    # Only Graph Connectors (custom properties)
```

### User Provisioning (Option A)
```bash
npm run provision           # Create users from CSV
npm run provision:beta      # Use beta Graph API
npm run list-users          # List provisioned users
npm run update-licenses     # Add missing licenses
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
│   ├── enrich-profiles-hybrid.ts # Hybrid enrichment (Profile API + Connectors)
│   ├── enrich-profiles-direct.ts # Profile API only
│   ├── enrich-profiles.ts        # Graph Connectors only
│   ├── profile-writer.ts         # Profile API writer class
│   ├── graph-client.ts           # Microsoft Graph client
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
| `src/enrich-profiles-hybrid.ts` | Main enrichment - Profile API + Connectors |
| `src/profile-writer.ts` | Profile API writer (skills, languages, notes, etc.) |
| `src/provision.ts` | User provisioning (Option A) |
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

# 3. Enrich profiles with setup (creates Graph Connector)
npm run enrich:setup -- --csv config/users.csv
```

### Regular Enrichment

```bash
# Update profiles from CSV
npm run enrich -- --csv config/users.csv
```

### Profile API Only (No Custom Properties)

```bash
# If you only need skills, languages, notes, interests
npm run enrich:profile-api-only -- --csv config/users.csv
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

**Last Updated**: 2026-01-27
**Version**: 3.2.0 (Fixed beta API requirement for people data connectors)
