# Architecture: Option A vs Option B

**Last Updated**: 2026-01-22

## Overview

This project separates user provisioning into two distinct operations:

1. **Option A (Core Provisioning)**: Create Entra ID users with standard properties
2. **Option B (Profile Enrichment)**: Enrich profiles via Microsoft Graph Connectors

## Architecture Diagram

```
┌────────────────────────────────────────────────────────────┐
│ CSV Input File                                              │
│ - Standard properties (name, email, jobTitle, etc.)        │
│ - Enrichment properties (skills, interests, aboutMe, etc.) │
│ - Custom properties (VTeam, BenefitPlan, etc.)             │
└─────────────────┬──────────────────────────────────────────┘
                  │
        ┌─────────┴─────────┐
        │                   │
        ▼                   ▼
┌─────────────────┐   ┌─────────────────┐
│   Option A      │   │   Option B      │
│  provision.ts   │   │enrich-profiles  │
└────────┬────────┘   └────────┬────────┘
         │                     │
         ▼                     ▼
┌─────────────────┐   ┌─────────────────┐
│  Entra ID       │   │ Graph Connector │
│  User Objects   │   │ External Items  │
│                 │   │                 │
│ Standard props  │   │ Enrichment data │
└─────────────────┘   └─────────────────┘
         │                     │
         └──────────┬──────────┘
                    ▼
         ┌─────────────────────┐
         │  Microsoft 365      │
         │  - Profile Cards    │
         │  - Search           │
         │  - Copilot          │
         └─────────────────────┘
```

## Option A: Core User Provisioning

### Purpose
Create Entra ID user accounts with standard Microsoft 365 properties.

### File
`src/provision.ts`

### Properties Handled
- **Identity**: givenName, surname, displayName, userPrincipalName
- **Job**: jobTitle, department, employeeType, companyName
- **Contact**: mail, mobilePhone, businessPhones
- **Location**: officeLocation, city, state, country, usageLocation
- **Employment**: employeeId, employeeHireDate
- **Manager**: manager relationships
- **Licenses**: M365 license assignment (requires usageLocation)

### Authentication
- OAuth 2.0 Authorization Code Flow (browser-based)
- Delegated Permissions:
  - User.ReadWrite.All
  - Directory.ReadWrite.All
  - Organization.Read.All

### Operations
- Uses Microsoft Graph v1.0 and beta endpoints
- Batch operations (20 users per batch)
- State-based diff detection (CREATE/UPDATE/NOOP)
- Manager relationship assignment

### Usage
```bash
npm run provision -- --csv config/agents-template.csv
```

### Output
- Real Entra ID user accounts
- Assigned licenses
- Manager relationships
- State file for tracking changes

## Option B: Profile Enrichment via Graph Connectors

### Purpose
Enrich user profiles with additional data that surfaces in Microsoft 365 Copilot, Search, and Profile Cards.

### File
`src/enrich-profiles.ts`

### Properties Handled

**Official People Data** (with Microsoft labels):
- skills → personSkills
- pastProjects → personProjects
- certifications → personCertifications
- awards → personAwards
- aboutMe → personNote
- mySite → personWebSite
- birthday → personAnniversaries

**Custom People Data** (searchable):
- interests
- responsibilities
- schools

**Custom Organization Properties**:
- VTeam, BenefitPlan, CostCenter, BuildingAccess, ProjectCode
- Any additional CSV columns not in schema

### Architecture

```
Graph Connector (m365provisionpeople)
├── Connection (one-time setup)
├── Schema (16 properties, registered once)
└── External Items (one per person)
    ├── id: person-email-domain-com
    ├── accountInformation: links to Entra ID user
    ├── properties: enrichment data
    └── acl: everyone in organization
```

### Authentication
- OAuth 2.0 Client Credentials Flow (app-only)
- Application Permissions:
  - ExternalConnection.ReadWrite.OwnedBy
  - ExternalItem.ReadWrite.OwnedBy
- Requires client secret in `.env`

### Operations
- Uses Microsoft Graph beta endpoint (required for People Data)
- Creates Graph Connector connection (once)
- Registers schema with People Data labels (once)
- Ingests external items linked to Entra ID users
- Individual PUT requests per item (100ms delay for rate limiting)
- Automatic deletion of orphaned items (state-based tracking)

### Usage
```bash
# First time: Setup connection and schema
npm run enrich-profiles:setup

# Ingest data
npm run enrich-profiles -- --csv config/agents-template.csv
```

### Output
- External items in Graph Connector
- State file tracking created items (`state/external-items-state.json`)
- Data surfaces in Microsoft Search
- Data appears in M365 profile cards
- Data available to Copilot
- Automatic cleanup of orphaned items

## Key Architectural Decisions

### 1. Why Separate Option A and Option B?

**Different APIs**:
- Option A: Direct Entra ID user object properties
- Option B: Graph Connector external items

**Different Authentication**:
- Option A: Delegated (user sign-in)
- Option B: Application (client secret)

**Different Purposes**:
- Option A: Core identity and organizational structure
- Option B: Enrichment and discovery

### 2. Why Graph Connectors for Option B?

**Tested Alternatives**:

❌ **Direct User Properties** (skills, interests, aboutMe):
- Cannot be set via batch operations
- Require individual PATCH requests per property
- Not efficient for bulk provisioning
- Limited to predefined schema

❌ **Open Extensions** (custom properties):
- Tested and working for custom properties
- But doesn't support People Data labels
- Data doesn't surface in Copilot
- Decided to move ALL enrichment to Graph Connectors

✅ **Graph Connectors** (final solution):
- Supports official People Data labels
- Data surfaces in Copilot responses
- Appears in M365 profile cards
- Flexible schema (add properties without code changes)
- Single enrichment system
- Searchable in Microsoft Search

### 3. Why Client Secret for Option B?

**Discovery Process**:

1. Initially tried delegated authentication (browser sign-in)
   - Connection management worked ✅
   - Item ingestion failed with 401 ❌

2. Verified token had correct scopes
   - ExternalConnection.ReadWrite.All ✅
   - ExternalItem.ReadWrite.All ✅
   - Still got 401 errors ❌

3. Discovered requirement
   - Graph Connectors operate as background services
   - Require Application permissions (not Delegated)
   - Need OAuth 2.0 Client Credentials Flow

4. Implemented client secret authentication
   - Item ingestion succeeded ✅
   - All 3 test items verified ✅

See [PEOPLE-DATA-AUTH-SOLUTION.md](./PEOPLE-DATA-AUTH-SOLUTION.md) for details.

## Data Flow

### Same CSV File for Both Options

```csv
name,email,jobTitle,department,skills,interests,aboutMe,VTeam
Sarah Chen,sarah@domain.com,CEO,Executive,"['Leadership']","['Innovation']",Experienced...,Platform
```

### Option A Processing

1. Parse CSV → Extract standard properties
2. Batch create/update users in Entra ID
3. Assign licenses
4. Set manager relationships
5. Log deferred properties (Option B properties)

### Option B Processing

1. Parse CSV → Extract enrichment properties
2. Create external items with People Data labels
3. Link to Entra ID users via userPrincipalName
4. Ingest to Graph Connector
5. Data becomes searchable in M365

## Property Classification

| Category | Properties | Handled By | Storage | Notes |
|----------|-----------|------------|---------|-------|
| **Standard Identity** | givenName, surname, displayName, userPrincipalName, mail | Option A | Entra ID user object | Core identity fields |
| **Job Information** | jobTitle, department, employeeType, companyName, officeLocation | Option A | Entra ID user object | Organizational structure |
| **Location & Regional** | usageLocation, preferredLanguage, city, state, country | Option A | Entra ID user object | Required for licensing |
| **Contact** | mobilePhone, businessPhones | Option A | Entra ID user object | Communication |
| **Official People Data** | skills, pastProjects, certifications, awards, aboutMe, mySite, birthday | Option B | Graph Connector | Has People Data labels |
| **Custom People Data** | interests, responsibilities, schools | Option B | Graph Connector | No official labels, searchable |
| **Custom Organization** | VTeam, BenefitPlan, CostCenter, etc. | Option B | Graph Connector | Organization-specific fields |

## Performance Characteristics

### Option A
- **Speed**: Fast (batch operations)
- **100 users**: ~10 seconds
- **Rate limiting**: Rarely an issue (20 users/batch)
- **Bottleneck**: License assignment (sequential)

### Option B
- **Speed**: Moderate (individual requests)
- **100 users**: ~2 minutes
- **Rate limiting**: 100ms delay between items
- **Bottleneck**: Item ingestion rate

### Combined Flow
```bash
# Total time for 100 users:
npm run provision          # ~10 seconds (Option A)
npm run enrich-profiles    # ~2 minutes (Option B)
# Total: ~2 minutes 10 seconds
```

## Workflow Recommendations

### Initial Provisioning (New Environment)

```bash
# Step 1: Setup Graph Connector (once)
npm run enrich-profiles:setup

# Step 2: Create users
npm run provision -- --csv config/agents-template.csv

# Step 3: Enrich profiles
npm run enrich-profiles -- --csv config/agents-template.csv
```

### Updating Existing Users

```bash
# Update standard properties (job titles, departments, etc.)
npm run provision -- --csv config/agents-template.csv

# Update enrichment data (skills, interests, etc.)
npm run enrich-profiles -- --csv config/agents-template.csv
```

### Adding New Users

```bash
# Add new rows to CSV, then:
npm run provision          # Creates new users in Entra ID
npm run enrich-profiles    # Creates enrichment data
```

## Benefits of This Architecture

### Separation of Concerns
✅ Clear responsibility: identity vs. enrichment
✅ Independent operations: can run separately
✅ Different authentication: appropriate for each use case

### Flexibility
✅ Add enrichment properties without changing Option A
✅ Dynamic schema: CSV columns auto-detected
✅ Same CSV file: single source of truth

### Copilot Integration
✅ Official People Data labels recognized by Copilot
✅ Data surfaces in AI responses
✅ Searchable in Microsoft Search
✅ Appears in M365 profile cards

### Maintainability
✅ Clear code organization
✅ Reusable modules (connection-manager, schema-builder, item-ingester)
✅ Comprehensive logging
✅ Error handling with detailed messages

## Security Considerations

### Option A (Delegated)
- User signs in via browser
- Delegated permissions (acts as user)
- Token cached locally (~/.m365-provision/)
- MFA supported
- Audit trail: operations tied to signed-in admin

### Option B (Application)
- Client secret stored in `.env`
- Application permissions (acts as app)
- No user context needed
- Requires secret management
- Audit trail: operations tied to app registration

### Best Practices
1. Never commit `.env` to git
2. Rotate client secrets regularly (set expiration)
3. Use least privilege (`.OwnedBy` not `.All`)
4. Consider certificate-based auth for production
5. Monitor Azure AD sign-in logs

## Troubleshooting

### Option A Issues
- **401 errors**: Token expired, run `npm run logout` and re-authenticate
- **License assignment fails**: Check available licenses
- **Manager not found**: Ensure manager exists before assigning

### Option B Issues
- **401 Unauthenticated**: Verify client secret and Application permissions
- **Schema stuck in draft**: Wait 10 minutes, schema provisioning is slow
- **Items not searchable**: Wait 1-2 hours for indexing
- **Invalid client secret**: Create new secret, copy VALUE (not ID)
- **Orphaned items not deleted**: Verify state file exists at `state/external-items-state.json`

## Documentation

- **[OPTION-B-IMPLEMENTATION-GUIDE.md](./OPTION-B-IMPLEMENTATION-GUIDE.md)**: Complete Option B guide
- **[PEOPLE-DATA-AUTH-SOLUTION.md](./PEOPLE-DATA-AUTH-SOLUTION.md)**: Authentication learnings
- **[OPTION-A-QUICK-START.md](./OPTION-A-QUICK-START.md)**: Option A quick start
- **[STATE-MANAGEMENT.md](./STATE-MANAGEMENT.md)**: State-based provisioning

---

**Status**: ✅ Both options implemented and working
**Last Tested**: 2026-01-22
**Test Results**: 3/3 users successfully provisioned and enriched
