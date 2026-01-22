# M365 Agent Provisioning

Bulk provisioning tool for creating Microsoft 365 user accounts with profile enrichment via Microsoft Graph Connectors.

## Overview

This tool provides two complementary capabilities:

### Option A: User Provisioning
- Create Microsoft 365 user accounts via Entra ID
- Assign Microsoft 365 licenses (E3/E5)
- Set standard user properties (name, email, jobTitle, department, etc.)
- Uses **OAuth 2.0 Authorization Code Flow** (browser-based, delegated permissions)

### Option B: Profile Enrichment
- Enrich user profiles with extended "People Data" attributes
- Skills, certifications, interests, awards, projects, schools
- Uses **Microsoft Graph Connectors** to make data searchable in M365 Copilot
- Uses **OAuth 2.0 Client Credentials Flow** (application permissions)

## Quick Start

```bash
# Install dependencies
npm install

# Copy and configure environment variables
cp .env.example .env
# Edit .env with your Azure AD credentials

# Option A: Create users
npm run provision -- --csv config/agents-template.csv

# Option B: Enrich profiles (first time - sets up Graph Connector)
npm run enrich-profiles:setup

# Option B: Enrich profiles (subsequent runs)
npm run enrich-profiles -- --csv config/agents-template.csv
```

## Prerequisites

- Node.js 18+
- Azure AD administrator access
- Microsoft 365 E3/E5 licenses available
- Azure AD app registration with appropriate permissions

## Documentation

All documentation is in the `docs/` folder:

- **[docs/SETUP.md](./docs/SETUP.md)**: Azure AD app registration and environment setup
- **[docs/USAGE.md](./docs/USAGE.md)**: Detailed usage instructions and workflow
- **[docs/README.md](./docs/README.md)**: Documentation index with all guides
- **[docs/ARCHITECTURE-OPTION-A-B.md](./docs/ARCHITECTURE-OPTION-A-B.md)**: System architecture

## Project Structure

```
M365-Agent-Provisioning/
├── src/                          # TypeScript source code
│   ├── provision.ts              # Main CLI - user provisioning (Option A)
│   ├── enrich-profiles.ts        # Profile enrichment entry point (Option B)
│   ├── graph-client.ts           # Microsoft Graph API client
│   ├── graph-beta-client.ts      # Graph Beta API client
│   ├── state-manager.ts          # State tracking (CREATE/UPDATE/NOOP)
│   ├── export.ts                 # Configuration export
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
├── CLAUDE.MD                     # Claude Code instructions
├── package.json                  # Dependencies and scripts
└── tsconfig.json                 # TypeScript configuration
```

## npm Scripts

### Provisioning (Option A)
```bash
npm run provision              # Create users from CSV
npm run provision:beta         # Create users with beta Graph API
npm run list-users             # List provisioned users
npm run list-licenses          # List available licenses
npm run cleanup                # Delete provisioned users
```

### Profile Enrichment (Option B)
```bash
npm run enrich-profiles:setup  # Setup Graph Connector (first time)
npm run enrich-profiles        # Enrich profiles from CSV
npm run enrich-profiles:dry-run # Preview without changes
npm run enrich-profiles:wait   # Wait for schema to be ready
```

### Authentication
```bash
npm run logout                 # Clear cached tokens
npm run test-connection        # Test Graph API connection
npm run test-beta              # Test Beta API availability
```

### Development
```bash
npm run build                  # Compile TypeScript
npm run dev                    # Watch mode compilation
npm run clean                  # Remove dist folder
```

## Authentication Architecture

| Feature | Option A (Users) | Option B (Enrichment) |
|---------|------------------|----------------------|
| OAuth Flow | Authorization Code | Client Credentials |
| Permission Type | Delegated | Application |
| User Context | Admin signs in via browser | App-only (no user) |
| Main Permissions | User.ReadWrite.All, Directory.ReadWrite.All | ExternalConnection.ReadWrite.OwnedBy, ExternalItem.ReadWrite.OwnedBy |

## License

MIT
