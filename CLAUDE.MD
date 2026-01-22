# M365 Agent Provisioning - Claude Code Documentation

## Project Overview

**M365-Agent-Provisioning** is a TypeScript CLI tool for bulk provisioning Microsoft 365 users with automatic license assignment and profile enrichment via Microsoft Graph Connectors.

### Key Capabilities

- **Option A - User Provisioning**: Create M365 users with standard properties
- **Option B - Profile Enrichment**: Add extended "People Data" via Graph Connectors
- **License Management**: Automatically assign Microsoft 365 licenses
- **State Management**: CREATE/UPDATE/NOOP detection to avoid duplicate operations
- **Configuration Export**: Output user details to JSON for downstream consumption

## Architecture

### Two Complementary Options

```
┌─────────────────────────────────────────────────────────────────┐
│  OPTION A: User Provisioning                                    │
│  - Creates Entra ID user accounts                               │
│  - Sets standard properties (name, email, jobTitle, etc.)       │
│  - Assigns M365 licenses                                        │
│  - Auth: OAuth 2.0 Authorization Code Flow (browser-based)      │
└────────────────┬────────────────────────────────────────────────┘
                 │
                 │ Delegated Permissions (User.ReadWrite.All)
                 ▼
┌─────────────────────────────────────────────────────────────────┐
│  Microsoft Graph API v1.0 + Beta                                │
│  POST /users, PATCH /users/{id}, POST /users/{id}/assignLicense │
└─────────────────────────────────────────────────────────────────┘

┌─────────────────────────────────────────────────────────────────┐
│  OPTION B: Profile Enrichment                                   │
│  - Creates Graph Connector connection                           │
│  - Ingests "People Data" (skills, certifications, etc.)         │
│  - Data searchable in M365 Copilot                              │
│  - Auth: OAuth 2.0 Client Credentials Flow (app-only)           │
└────────────────┬────────────────────────────────────────────────┘
                 │
                 │ Application Permissions (ExternalItem.ReadWrite.OwnedBy)
                 ▼
┌─────────────────────────────────────────────────────────────────┐
│  Microsoft Graph Connectors API (Beta)                          │
│  POST /external/connections, PUT /external/connections/{id}/items/{id} │
└─────────────────────────────────────────────────────────────────┘
```

### Why Two Options?

| Aspect | Option A | Option B |
|--------|----------|----------|
| **Purpose** | Identity management | Data enrichment |
| **API** | Entra ID (User management) | Graph Connectors (External data) |
| **Auth** | Delegated (user signs in) | Application (client secret) |
| **Data Stored** | Entra ID directory | Graph Connector index |
| **Searchable In** | GAL, Teams, etc. | M365 Copilot, Search |

## Authentication Architecture

### Option A: Authorization Code Flow (Browser-Based)

```
1. User runs: npm run provision
2. CLI starts local HTTP server on port 5544
3. Browser opens Azure AD login page
4. User authenticates with admin account
5. Azure AD redirects to localhost with auth code
6. CLI exchanges code for tokens
7. Tokens cached in ~/.m365-provision/token-cache.json
```

**Key Files:**
- `src/auth/browser-auth-server.ts` - OAuth callback server
- `src/auth/token-cache.ts` - Token persistence

### Option B: Client Credentials Flow (App-Only)

```
1. CLI reads AZURE_CLIENT_SECRET from .env
2. Uses @azure/identity ClientSecretCredential
3. Requests token with .default scope
4. Token used directly (no caching needed for app-only)
```

**Key Files:**
- `src/enrich-profiles.ts` - Main entry point
- `src/people-connector/item-ingester.ts` - Token usage

## Environment Variables

```bash
# Azure AD Configuration
AZURE_TENANT_ID=your-tenant-id
AZURE_CLIENT_ID=your-client-id
AZURE_CLIENT_SECRET=your-client-secret  # Required for Option B

# Graph API Configuration
GRAPH_API_ENDPOINT=https://graph.microsoft.com
USE_BETA_ENDPOINTS=true

# Provisioning Configuration
# Multiple licenses supported (comma-separated)
LICENSE_SKU_IDS=sku-id-1,sku-id-2
USER_DOMAIN=yourdomain.onmicrosoft.com

# Auth Server (Option A)
AUTH_SERVER_PORT=5544
```

## Azure AD Permissions

### Option A (Delegated)
- `User.ReadWrite.All` - Create and manage users
- `Directory.ReadWrite.All` - Full directory access
- `Organization.Read.All` - Read organization details
- `offline_access` - Refresh tokens

### Option B (Application)
- `ExternalConnection.ReadWrite.OwnedBy` - Manage owned connections
- `ExternalItem.ReadWrite.OwnedBy` - Manage items in owned connections

## Project Structure

```
M365-Agent-Provisioning/
├── src/                          # TypeScript source code
│   ├── provision.ts              # Main CLI - Option A user provisioning
│   ├── enrich-profiles.ts        # Main CLI - Option B profile enrichment
│   ├── graph-client.ts           # Microsoft Graph API client
│   ├── graph-beta-client.ts      # Graph Beta API client
│   ├── state-manager.ts          # State tracking (CREATE/UPDATE/NOOP)
│   ├── export.ts                 # JSON configuration export
│   ├── wait-for-schema.ts        # Schema readiness checker
│   ├── auth/                     # Authentication modules
│   │   ├── browser-auth-server.ts  # OAuth callback server
│   │   └── token-cache.ts          # Token persistence
│   ├── people-connector/         # Graph Connector modules (Option B)
│   │   ├── connection-manager.ts   # Connection lifecycle
│   │   ├── schema-builder.ts       # Schema definition
│   │   └── item-ingester.ts        # External item creation
│   ├── schema/                   # Property schema definitions
│   │   └── user-property-schema.ts
│   ├── safety/                   # Security modules
│   │   └── account-protection.ts   # Real account protection
│   ├── extensions/               # Graph extensions (experimental)
│   │   └── open-extension-manager.ts
│   └── utils/                    # Utilities
│       └── logger.ts
│
├── config/                       # Configuration files
│   ├── agents-template.csv       # Main agent definitions
│   └── agents-test-*.csv         # Test configurations
│
├── output/                       # Generated output (gitignored)
│   ├── agents-config.json        # Provisioned user details
│   └── provisioning-report.md    # Provisioning summary
│
├── state/                        # State tracking (gitignored)
│   └── external-items-state.json # Graph Connector item state
│
├── logs/                         # Log files (gitignored)
│
├── tools/                        # Utility scripts
│   ├── debug/                    # Debug and testing scripts
│   │   ├── check-app-permissions.mjs
│   │   ├── check-connection.mjs
│   │   ├── test-*.mjs
│   │   └── verify-*.mjs
│   └── admin/                    # Administrative scripts
│       ├── cleanup-orphaned-items.mjs
│       └── delete-connection.mjs
│
├── public/                       # Static files
│   └── auth.html                 # OAuth callback page
│
├── docs/                         # Documentation
│   ├── README.md                 # Documentation index
│   ├── SETUP.md                  # Setup guide
│   ├── USAGE.md                  # Usage guide
│   └── *.md                      # Additional guides
│
├── dist/                         # Compiled JavaScript (generated)
├── .env                          # Environment variables (gitignored)
├── .env.example                  # Environment template
└── package.json                  # Dependencies and scripts
```

## Critical Files

### Core Modules

| File | Purpose |
|------|---------|
| `src/provision.ts` | Main CLI - Option A user provisioning |
| `src/enrich-profiles.ts` | Main CLI - Option B profile enrichment |
| `src/reset-tenant.ts` | Tenant reset with account protection |
| `src/graph-client.ts` | Microsoft Graph API client (supports multiple licenses) |
| `src/graph-beta-client.ts` | Graph Beta API client |
| `src/state-manager.ts` | State tracking (CREATE/UPDATE/NOOP) |
| `src/export.ts` | JSON configuration export |

### Authentication

| File | Purpose |
|------|---------|
| `src/auth/browser-auth-server.ts` | OAuth callback server for Option A |
| `src/auth/token-cache.ts` | Token persistence |

### Graph Connector (Option B)

| File | Purpose |
|------|---------|
| `src/people-connector/connection-manager.ts` | Connection lifecycle |
| `src/people-connector/schema-builder.ts` | Schema definition |
| `src/people-connector/item-ingester.ts` | External item creation |

### Supporting Modules

| File | Purpose |
|------|---------|
| `src/schema/user-property-schema.ts` | Property definitions |
| `src/safety/account-protection.ts` | Real account protection |
| `src/utils/logger.ts` | Logging utility |
| `src/wait-for-schema.ts` | Schema readiness checker |

## npm Scripts

```bash
# Option A - User Provisioning
npm run provision              # Standard provisioning
npm run provision:beta         # With beta Graph API
npm run list-users             # List provisioned users
npm run list-licenses          # List available licenses (uses delegated auth)
npm run cleanup                # Delete provisioned users

# Tenant Management
npm run reset-tenant           # Preview users to delete (dry run)
npm run reset-tenant:confirm   # Delete all non-admin users

# Option B - Profile Enrichment
npm run enrich-profiles:setup  # Setup Graph Connector
npm run enrich-profiles        # Enrich profiles from CSV
npm run enrich-profiles:dry-run # Preview without changes
npm run enrich-profiles:wait   # Wait for schema ready

# Authentication
npm run logout                 # Clear cached tokens
npm run test-connection        # Test Graph API (uses delegated auth)
npm run test-beta              # Test Beta API

# Development
npm run build                  # Compile TypeScript
npm run dev                    # Watch mode
npm run clean                  # Remove dist/
```

## State Management

The tool uses state files to track what has been provisioned:

- **Option A State**: Tracked via Graph API (check if user exists)
- **Option B State**: `state/external-items-state.json`

### State Actions

| Action | When |
|--------|------|
| CREATE | Item doesn't exist |
| UPDATE | Item exists but data changed |
| NOOP | Item exists with same data |

## Output Files

| File | Purpose |
|------|---------|
| `output/agents-config.json` | Provisioned user details |
| `output/provisioning-report.md` | Human-readable summary |
| `output/passwords.txt` | Generated passwords (secure!) |

## Tools Directory

### Debug Tools (`tools/debug/`)

Ad-hoc scripts for debugging and testing. These often have hardcoded values and are not part of the main workflow.

| Script | Purpose |
|--------|---------|
| `check-app-permissions.mjs` | Verify Azure AD app permissions |
| `check-connection.mjs` | Check Graph Connector status |
| `check-token-scopes.mjs` | Inspect token scopes |
| `test-*.mjs` | Various test scripts |
| `verify-*.mjs` | Verification scripts |

### Admin Tools (`tools/admin/`)

Administrative utilities for managing the Graph Connector.

| Script | Purpose |
|--------|---------|
| `cleanup-orphaned-items.mjs` | Delete orphaned external items |
| `delete-connection.mjs` | Delete Graph Connector connection |

## Documentation

All user documentation is in `docs/`:

| Document | Purpose |
|----------|---------|
| `docs/README.md` | Documentation index |
| `docs/SETUP.md` | Azure AD setup guide |
| `docs/USAGE.md` | Usage instructions |
| `docs/ARCHITECTURE-OPTION-A-B.md` | Architecture details |
| `docs/PEOPLE-DATA-AUTH-SOLUTION.md` | Auth solution for Option B |
| `docs/OPTION-B-IMPLEMENTATION-GUIDE.md` | Option B implementation |
| `docs/STATE-MANAGEMENT.md` | State tracking details |

## Common Workflows

### First-Time Setup

```bash
# 1. Configure environment
cp .env.example .env
# Edit .env with credentials

# 2. Setup Graph Connector (Option B)
npm run enrich-profiles:setup

# 3. Wait for schema to be ready
npm run enrich-profiles:wait

# 4. Create users (Option A)
npm run provision

# 5. Enrich profiles (Option B)
npm run enrich-profiles
```

### Update Existing Users

```bash
# Update CSV file, then:
npm run provision          # Updates standard properties
npm run enrich-profiles    # Updates enrichment data
```

### Tenant Reset (Start Fresh)

```bash
# 1. Preview what will be deleted (safe - no changes)
npm run reset-tenant

# 2. Actually delete all non-admin users
npm run reset-tenant:confirm

# 3. Provision new users
npm run provision -- --csv config/new-users.csv
```

**Protected accounts are NEVER deleted:**
- Email patterns: `admin@*`, `administrator@*`, `root@*`, `systemadmin@*`
- Azure AD roles: Global Administrator, Security Administrator, etc.
- Custom exclusions from `.env` (PROTECTED_EMAILS, PROTECTED_EMAIL_PATTERNS)

## Troubleshooting

### Option A Issues

**"Invalid client"**: App not configured correctly in Azure AD
**"Insufficient privileges"**: Missing delegated permissions or admin consent
**"License assignment failed"**: No available licenses

### Option B Issues

**"401 Unauthorized"**: Missing client secret or application permissions
**"Schema stuck in draft"**: Wait 10+ minutes, provisioning is slow
**"Items not searchable"**: Wait 1-2 hours for indexing

### Common Fixes

```bash
# Clear cached tokens and re-authenticate
npm run logout

# Rebuild TypeScript
npm run build

# Check Graph Connector status
node tools/debug/check-connection.mjs
```

## Security Considerations

### Token Storage
- Option A tokens: `~/.m365-provision/token-cache.json` (chmod 600)
- Option B: Uses client secret from .env (never commit!)

### Best Practices
- Never commit `.env` files
- Keep `output/` directory gitignored
- Rotate client secrets periodically
- Use separate Azure AD apps for dev/prod

## Key Learnings

### Graph Connectors Require Application Permissions

Graph Connectors operate as background services and require OAuth 2.0 Client Credentials Flow with Application Permissions. Delegated permissions (user sign-in) will result in 401 errors.

### Schema Provisioning Takes Time

After creating a Graph Connector connection, the schema must transition from "draft" to "ready" state. This can take 10+ minutes. Use `npm run enrich-profiles:wait` to poll for readiness.

### Item Indexing Has Latency

After ingesting external items, they may not be immediately searchable. Allow 1-2 hours for full indexing in M365 Search/Copilot.

---

**Last Updated**: 2026-01-22
**Version**: 2.0.0 (OAuth Authorization Code + Client Credentials Flows)
