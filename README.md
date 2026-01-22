# How a CSV File Can Be the Master of Your M365 Tenant User Provisioning

> **A Learning Project**: This tool was built to explore Microsoft Graph APIs, OAuth 2.0 flows, and Graph Connectors. It works, but more importantly, it teaches. Check the [docs/](./docs/) folder for our learnings, mistakes, and solutions.

**TL;DR**: Define your users in a CSV → Run one command → Users created in Microsoft 365 with licenses assigned and profiles enriched.

```
┌─────────────────┐      ┌─────────────────┐      ┌─────────────────┐
│   agents.csv    │  →   │  npm run        │  →   │  Microsoft 365  │
│   (your data)   │      │  provision      │      │  (users ready)  │
└─────────────────┘      └─────────────────┘      └─────────────────┘
```

---

## Quick Demo: From CSV to M365 Users

### The Smallest CSV (4 columns)

This is all you need to create users:

```csv
name,email,role,department
Sarah Chen,sarah.chen@yourdomain.onmicrosoft.com,CEO,Executive
Michael Rodriguez,michael.rodriguez@yourdomain.onmicrosoft.com,CTO,Engineering
Emma Wilson,emma.wilson@yourdomain.onmicrosoft.com,Developer,Engineering
```

Run it:
```bash
npm run provision -- --csv config/my-team.csv
```

Result: 3 users created with M365 licenses, ready to use Outlook, Teams, and Office apps.

---

### The Full CSV (30+ columns with custom properties)

When you need the complete employee profile with enrichment data:

```csv
name,email,role,department,givenName,surname,jobTitle,employeeType,companyName,officeLocation,streetAddress,city,state,country,postalCode,usageLocation,preferredLanguage,mobilePhone,businessPhones,employeeId,employeeHireDate,ManagerEmail,VTeam,BenefitPlan,CostCenter,BuildingAccess,ProjectCode
Ingrid Johansen,ingrid.johansen@domain.onmicrosoft.com,CEO,Executive,Ingrid,Johansen,Chief Executive Officer,Employee,Nordic Solutions AS,Oslo HQ,Drammensveien 134,Oslo,Oslo,Norway,0277,NO,nb-NO,+47 915 12 345,"['+47 22 12 34 56']",EMP001,2015-03-15,,Executive Leadership,Executive Plus,CEO-OFFICE,Level-5,EXEC-001
Lars Hansen,lars.hansen@domain.onmicrosoft.com,CTO,Engineering,Lars,Hansen,Chief Technology Officer,Employee,Nordic Solutions AS,Oslo HQ,Drammensveien 134,Oslo,Oslo,Norway,0277,NO,nb-NO,+47 918 45 678,"['+47 22 12 34 57']",EMP002,2016-06-01,ingrid.johansen@domain.onmicrosoft.com,Executive Leadership,Executive Plus,ENG-DEPT,Level-5,ENG-000
```

**Custom columns** (`VTeam`, `BenefitPlan`, `CostCenter`, `BuildingAccess`, `ProjectCode`) become searchable in M365 Copilot via Graph Connectors.

Run both provisioning + enrichment:
```bash
npm run provision -- --csv config/full-team.csv
npm run enrich-profiles -- --csv config/full-team.csv
```

---

## What This Tool Does

| Feature | What Happens |
|---------|--------------|
| **User Creation** | Creates Entra ID accounts from CSV |
| **License Assignment** | Automatically assigns M365 E3/E5 licenses |
| **Manager Hierarchy** | Sets up reporting structure via `ManagerEmail` column |
| **Profile Enrichment** | Adds skills, interests, certifications to profiles |
| **Custom Properties** | Any extra CSV column becomes searchable in Copilot |
| **State Management** | Detects CREATE/UPDATE/NOOP - won't duplicate users |

---

## Setup (One-Time)

### 1. Azure AD App Registration

Create an app in [Azure Portal](https://portal.azure.com) → Azure Active Directory → App registrations.

**Required API Permissions:**

![API Permissions](./docs/images/api-permissions.png)

| Permission | Type | Purpose |
|------------|------|---------|
| `User.ReadWrite.All` | Delegated | Create and update users |
| `Directory.ReadWrite.All` | Delegated | Manage directory objects |
| `Organization.Read.All` | Delegated | Read tenant info |
| `offline_access` | Delegated | Keep tokens refreshed |
| `openid` | Delegated | Sign users in |
| `profile` | Delegated | Read user profiles |
| `ExternalConnection.ReadWrite.OwnedBy` | **Application** | Graph Connector (enrichment) |
| `ExternalItem.ReadWrite.OwnedBy` | **Application** | Graph Connector items |

> **Important**: Click "Grant admin consent" after adding permissions!

### 2. Configure Environment

```bash
cp .env.example .env
```

Edit `.env`:
```bash
AZURE_TENANT_ID=your-tenant-id
AZURE_CLIENT_ID=your-client-id
AZURE_CLIENT_SECRET=your-client-secret  # For profile enrichment
LICENSE_SKU_ID=05e9a617-0261-4cee-bb44-138d3ef5d965  # M365 E3
USER_DOMAIN=yourdomain.onmicrosoft.com
```

### 3. Install & Build

```bash
npm install
npm run build
```

---

## Usage

### Basic: Create Users

```bash
# Preview what will be created (dry run)
npm run provision -- --dry-run --csv config/agents-template.csv

# Actually create users
npm run provision -- --csv config/agents-template.csv
```

### Advanced: Enrich Profiles with People Data

First-time setup (creates Graph Connector):
```bash
npm run enrich-profiles:setup
npm run enrich-profiles:wait  # Wait for schema to be ready (~10 min)
```

Then enrich:
```bash
npm run enrich-profiles -- --csv config/agents-template.csv
```

---

## CSV Column Reference

### Standard User Properties (Option A)

| Column | Required | Example | Maps To |
|--------|----------|---------|---------|
| `name` | ✅ | Sarah Chen | displayName |
| `email` | ✅ | sarah@domain.com | userPrincipalName |
| `role` | ✅ | CEO | (internal use) |
| `department` | ✅ | Executive | department |
| `givenName` | | Sarah | givenName |
| `surname` | | Chen | surname |
| `jobTitle` | | Chief Executive Officer | jobTitle |
| `employeeType` | | Employee | employeeType |
| `companyName` | | Contoso Ltd | companyName |
| `officeLocation` | | Building 1 | officeLocation |
| `city` | | Oslo | city |
| `country` | | Norway | country |
| `mobilePhone` | | +47 900 12 345 | mobilePhone |
| `ManagerEmail` | | boss@domain.com | manager (relationship) |

### Enrichment Properties (Option B - Graph Connector)

These become searchable in M365 Copilot:

| Column | Example | Searchable In |
|--------|---------|---------------|
| `skills` | "['TypeScript','Azure']" | Copilot: "Who knows Azure?" |
| `interests` | "['AI/ML','Hiking']" | Copilot: "Who's interested in AI?" |
| `aboutMe` | "10 years experience..." | Copilot: "Tell me about Sarah" |
| `schools` | "['MIT','Stanford']" | Copilot: "Who went to MIT?" |

### Custom Properties (Your Own Columns)

Any column not in the standard list becomes a custom property:

```csv
name,email,role,department,VTeam,CostCenter,SecurityClearance
Sarah Chen,sarah@domain.com,CEO,Executive,Leadership,C-SUITE,Level-5
```

`VTeam`, `CostCenter`, `SecurityClearance` → All searchable in Copilot!

---

## Output

After provisioning, check `output/`:

```
output/
├── agents-config.json    # User IDs, passwords, all details
├── provisioning-report.md # Human-readable summary
└── passwords.txt          # Generated passwords (KEEP SECURE!)
```

---

## Project Structure

```
├── src/                    # TypeScript source
│   ├── provision.ts        # User provisioning (Option A)
│   ├── enrich-profiles.ts  # Profile enrichment (Option B)
│   └── people-connector/   # Graph Connector modules
├── config/                 # Your CSV files go here
├── docs/                   # Deep-dive documentation
├── tools/                  # Debug & admin utilities
└── output/                 # Generated files (gitignored)
```

---

## Learnings & Documentation

This is a **learning project**. We documented everything we discovered:

| Document | What You'll Learn |
|----------|-------------------|
| [docs/ARCHITECTURE-OPTION-A-B.md](./docs/ARCHITECTURE-OPTION-A-B.md) | Why we split into two options |
| [docs/PEOPLE-DATA-AUTH-SOLUTION.md](./docs/PEOPLE-DATA-AUTH-SOLUTION.md) | Why Graph Connectors need Application permissions (we learned this the hard way) |
| [docs/STATE-MANAGEMENT.md](./docs/STATE-MANAGEMENT.md) | How we detect CREATE vs UPDATE vs NOOP |
| [docs/SETUP.md](./docs/SETUP.md) | Detailed Azure AD setup guide |
| [docs/USAGE.md](./docs/USAGE.md) | Complete usage reference |

**Key Learnings:**

1. **Delegated vs Application permissions matter** - Graph Connectors only work with Client Credentials Flow
2. **Schema provisioning is slow** - Wait 10+ minutes after creating a connection
3. **Items aren't instantly searchable** - Allow 1-2 hours for Copilot indexing
4. **State management prevents duplicates** - CSV is idempotent; run it as many times as you want

---

## Commands Reference

```bash
# Provisioning
npm run provision                      # Create users
npm run provision -- --dry-run         # Preview only
npm run list-users                     # List provisioned users
npm run list-licenses                  # Show available licenses

# Enrichment
npm run enrich-profiles:setup          # Setup Graph Connector
npm run enrich-profiles                # Enrich user profiles
npm run enrich-profiles:dry-run        # Preview enrichment

# Utilities
npm run logout                         # Clear auth tokens
npm run test-connection                # Verify Graph API access
```

---

## License

MIT

---

## Contributing

This is a learning project! Found something interesting? Learned something new? Open a PR to share your findings.
