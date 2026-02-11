# M365 Agent Provisioning - Documentation

**Last Updated**: 2026-01-22

## Quick Start

- **[OPTION-A-QUICK-START.md](./OPTION-A-QUICK-START.md)**: Get started with Option A (user provisioning)
- **[OPTION-B-IMPLEMENTATION-GUIDE.md](./OPTION-B-IMPLEMENTATION-GUIDE.md)**: Complete guide for Option B (profile enrichment)

## Architecture

- **[ARCHITECTURE-OPTION-A-B.md](./ARCHITECTURE-OPTION-A-B.md)**: System architecture and design decisions

## Authentication & Security

- **[PEOPLE-DATA-AUTH-SOLUTION.md](./PEOPLE-DATA-AUTH-SOLUTION.md)**: Authentication solution and learnings for Graph Connectors
- **[DEVICE-CODE-FLOW.md](./DEVICE-CODE-FLOW.md)**: OAuth 2.0 device code flow implementation (Option A)
- **[ACCOUNT-PROTECTION.md](./ACCOUNT-PROTECTION.md)**: Security best practices

## Features & Implementation

- **[STATE-MANAGEMENT.md](./STATE-MANAGEMENT.md)**: State-based provisioning (CREATE/UPDATE/NOOP)
- **[MANAGER-AND-LOGGING.md](./MANAGER-AND-LOGGING.md)**: Manager relationships and logging
- **[BETA-API-GUIDE.md](./BETA-API-GUIDE.md)**: Using Microsoft Graph beta endpoints

## Key Learnings

### Option B: Graph Connectors Require Application Permissions

**Problem**: Initially tried delegated authentication (user sign-in) but consistently got 401 errors when creating external items, even with correct scopes.

**Solution**: Graph Connectors operate as background services and require OAuth 2.0 Client Credentials Flow with Application Permissions:
- `ExternalConnection.ReadWrite.OwnedBy` (Application)
- `ExternalItem.ReadWrite.OwnedBy` (Application)

**Key Insight**: Both approaches are OAuth 2.0, just different flows:
- Authorization Code Flow = Delegated (user sign-in)
- Client Credentials Flow = Application (client secret)

See **[PEOPLE-DATA-AUTH-SOLUTION.md](./PEOPLE-DATA-AUTH-SOLUTION.md)** for complete details.

### Why Separate Option A and Option B?

1. **Different APIs**: Entra ID user objects vs. Graph Connector external items
2. **Different Authentication**: Delegated vs. Application
3. **Different Purposes**: Identity vs. Enrichment

### Data Classification

| Category | Handled By | Storage | Authentication |
|----------|------------|---------|----------------|
| Standard Properties (name, email, jobTitle) | Option A | Entra ID | Delegated (browser) |
| People Data (skills, certifications, awards) | Option B | Graph Connector | Application (client secret) |
| Custom Properties (VTeam, BenefitPlan) | Option B | Graph Connector | Application (client secret) |

## Testing & Migration

- **[OPTION-A-TEST-PLAN.md](./OPTION-A-TEST-PLAN.md)**: Testing strategy for Option A
- **[MSAL-MIGRATION.md](./MSAL-MIGRATION.md)**: Migration from client secret to MSAL (Option A)

## Additional Resources

- **[COPILOT-CONNECTORS-PEOPLE-DATA.md](./COPILOT-CONNECTORS-PEOPLE-DATA.md)**: People data connector deep dive

## Documentation Structure

```
docs/
├── README.md (this file)
│
├── Quick Start
│   ├── OPTION-A-QUICK-START.md
│   └── OPTION-B-IMPLEMENTATION-GUIDE.md
│
├── Core Concepts
│   ├── ARCHITECTURE-OPTION-A-B.md
│   ├── STATE-MANAGEMENT.md
│   └── MANAGER-AND-LOGGING.md
│
├── Authentication
│   ├── PEOPLE-DATA-AUTH-SOLUTION.md (Option B learnings)
│   ├── DEVICE-CODE-FLOW.md (Option A)
│   └── ACCOUNT-PROTECTION.md
│
├── API Reference
│   ├── BETA-API-GUIDE.md
│   └── GRAPH-CONNECTOR-PERMISSIONS.md
│
└── Historical
    ├── MSAL-MIGRATION.md
    └── OPTION-A-TEST-PLAN.md
```

## Common Workflows

### Initial Setup

```bash
# 1. Setup Graph Connector (Option B, once only)
npm run enrich-profiles:setup

# 2. Create users (Option A)
npm run provision -- --csv config/agents-template.csv

# 3. Enrich profiles (Option B)
npm run enrich-profiles -- --csv config/agents-template.csv
```

### Updating Data

```bash
# Update CSV file, then:

# Update standard properties
npm run provision

# Update enrichment data
npm run enrich-profiles
```

## Required Configuration

### .env File

```bash
# Azure AD
AZURE_TENANT_ID=your-tenant-id
AZURE_CLIENT_ID=your-client-id
AZURE_CLIENT_SECRET=your-client-secret  # For Option B only

# M365
LICENSE_SKU_ID=your-license-sku
USER_DOMAIN=yourdomain.onmicrosoft.com

# Required
USE_BETA_ENDPOINTS=true  # Beta endpoints are enforced
AUTH_SERVER_PORT=5544
```

### Azure AD Permissions

**Option A** (Delegated):
- User.ReadWrite.All
- Directory.ReadWrite.All
- Organization.Read.All

**Option B** (Application):
- ExternalConnection.ReadWrite.OwnedBy
- ExternalItem.ReadWrite.OwnedBy

## Getting Help

### Common Issues

1. **401 Errors in Option B**: Verify client secret and Application permissions
2. **Schema Stuck in Draft**: Wait 10 minutes, provisioning is slow
3. **Items Not Searchable**: Wait 1-2 hours for indexing
4. **License Assignment Fails**: Check available licenses in tenant

### Troubleshooting Guides

- Option A issues: See [OPTION-A-QUICK-START.md](./OPTION-A-QUICK-START.md)
- Option B issues: See [OPTION-B-IMPLEMENTATION-GUIDE.md](./OPTION-B-IMPLEMENTATION-GUIDE.md)
- Authentication: See [PEOPLE-DATA-AUTH-SOLUTION.md](./PEOPLE-DATA-AUTH-SOLUTION.md)
- Debug checklist: See [DEBUG-PLAYBOOK.md](./DEBUG-PLAYBOOK.md) and [debug/README.md](../debug/README.md)

## What's Working

✅ **Option A**: User provisioning with standard properties
✅ **Option B**: Profile enrichment via Graph Connectors
✅ **State Management**: CREATE/UPDATE/NOOP detection
✅ **Manager Relationships**: Automatic assignment
✅ **License Assignment**: Automatic M365 license assignment
✅ **Batch Operations**: 20 users per batch (Option A)
✅ **People Data Labels**: Official Microsoft labels for Copilot
✅ **Custom Properties**: Dynamic schema from CSV columns
✅ **Logging**: Comprehensive JSON logs

## Status

- **Option A**: ✅ Production Ready
- **Option B**: ✅ Production Ready
- **Last Tested**: 2026-01-22
- **Test Results**: 3/3 users successfully provisioned and enriched

---

For detailed implementation guides, see the individual documentation files listed above.
