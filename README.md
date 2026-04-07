# Make Copilot Know Your People

> Ask Copilot *"Who on my team speaks French and has an MBA?"* — and get silence. Copilot cannot answer questions about people it knows nothing about. This tool fixes that.

**The problem is simple**: Microsoft 365 profiles are empty by default. No skills. No education. No languages. No certifications. Copilot sees org chart boxes, not people. Every "find me someone who..." query fails — not because Copilot can't reason, but because there's nothing to reason over.

**This tool bridges the gap.** Define your people in a CSV or JSON file. Run two commands. Copilot can now search across skills, education, languages, patents, publications, and any custom property your organization cares about.

```
Your data (CSV or JSON)
        |
        v
┌──────────────────────────────────────────────────────────┐
│  Option A: Provision          Option B: Enrich           │
│  Entra ID accounts    →      Graph Connector with        │
│  + licenses                  18 people data labels       │
│  + manager hierarchy         + custom org properties     │
└──────────────────────────────────────────────────────────┘
        |
        v
Copilot can now answer:
  "Who has Python skills and speaks German?"
  "Find someone with a patent in distributed systems"
  "Who on the Stockholm team has an MBA?"
```

> **A Learning Project**: Built to explore Microsoft Graph APIs, OAuth 2.0 flows, and Graph Connectors. Everything we learned — including the mistakes — is in [docs/](./docs/).

---

## Two Minutes to Working Profiles

### 1. The simplest case: a CSV

Four columns. That's the minimum:

```csv
name,email,role,department
Sarah Chen,sarah@yourdomain.onmicrosoft.com,CEO,Executive
Michael Rodriguez,michael@yourdomain.onmicrosoft.com,CTO,Engineering
```

```bash
npm run provision -- --csv config/my-team.csv
```

Users created. Licenses assigned. Done.

### 2. Add enrichment data

Add columns. The tool routes each one automatically:

```csv
name,email,role,department,skills,languages,aboutMe,certifications,VTeam
Sarah Chen,sarah@domain.com,CEO,Executive,"['Leadership','Strategy']","['English (Native)']",20 years in tech,['PMP'],Executive
```

```bash
npm run provision -- --csv config/team.csv
npm run option-b:setup -- --csv config/team.csv --connection-id m365people25
```

Now Copilot knows Sarah has leadership skills and a PMP certification. Ask it.

### 3. Rich profiles with JSON

CSV works for simple data. JSON unlocks the full depth of Microsoft's profile schema — education with institutions and programs, languages with proficiency levels, patents with filing details, and per-user license assignment.

The JSON format uses PascalCase keys (matching the Microsoft Graph Profile API schema). The tool auto-detects PascalCase and normalizes it internally:

```json
[
  {
    "MailNickName": "sarah.chen",
    "DisplayName": "Sarah Chen",
    "FirstName": "Sarah",
    "LastName": "Chen",
    "JobTitle": "CEO",
    "Department": "Executive",
    "CompanyName": "Contoso",
    "Manager": null,
    "UsageLocation": "NO",
    "Address": {
      "City": "Oslo",
      "Country": "Norway",
      "CountryOrRegion": "Norway"
    },
    "Licenses": [
      "Office 365 E5 (no Teams)",
      "Microsoft Teams Enterprise",
      "Microsoft 365 Copilot"
    ],
    "Skills": [
      {
        "DisplayName": "Strategic Planning",
        "Categories": ["Business & Strategy"],
        "CollaborationTags": ["askMeAbout", "ableToMentor"],
        "Proficiency": "fullProfessional"
      }
    ],
    "EducationalActivities": [
      {
        "Institution": {
          "DisplayName": "Stanford University",
          "Location": { "City": "Stanford", "CountryOrRegion": "US" }
        },
        "Program": {
          "DisplayName": "MBA",
          "Abbreviation": "MBA",
          "FieldsOfStudy": "Strategy, Finance"
        },
        "StartMonthYear": "2010-09-01",
        "EndMonthYear": "2012-06-30"
      }
    ],
    "Languages": [
      { "DisplayName": "English", "Tag": "en-US", "Spoken": "nativeOrBilingual", "Written": "nativeOrBilingual", "Reading": "nativeOrBilingual" }
    ],
    "TerritoryTier1": "Microsoft",
    "TerritoryTier2": "EMEA"
  }
]
```

```bash
npm run provision -- --json config/team.json
npm run option-b:setup -- --json config/team.json --connection-id m365people01
```

**Key features of JSON input:**
- `MailNickName` + `USER_DOMAIN` from `.env` constructs the full email (no hardcoded domains)
- `Licenses` array assigns per-user licenses by display name (resolved to SKU IDs automatically)
- `Manager` references other users by `MailNickName`
- Rich profile entities match the Microsoft Graph Profile API schema
- Custom org fields (`TerritoryTier1`, `Products`, etc.) become searchable connector properties

**CSV and JSON are not exclusive.** Use both — JSON overrides CSV per-property when you provide both:

```bash
npm run option-b:setup -- --csv config/team.csv --json config/rich-profiles.json --connection-id m365people01
```

No property is forced to be rich — a plain string like `"Leadership"` works identically to a full object with `CollaborationTags` and `Categories`.

---

## How Data Flows

Every field is automatically routed to where it belongs:

```
Your Data                       Destination                  Copilot Searchable?
────────────────────────────────────────────────────────────────────────────────
Entra ID fields (name, email,   Option A → Entra ID          Via composite labels
 jobTitle, city, phone, etc.)

Profile fields (skills, certs,  Option B → Graph Connector   Yes (people data labels)
 education, languages, patents,   (18 official labels)
 interests, aboutMe, awards...)

Custom org fields (VTeam,       Option B → Graph Connector   Yes (searchable custom
 CostCenter, any unknown col)     (auto-detected)              properties)
```

### The 18 People Data Labels

Graph Connectors with people data labels are what make Copilot searchable. Without labels, Copilot cannot find your people by their attributes. This tool enables all 18:

| Group | Labels | What They Carry |
|-------|--------|-----------------|
| **Core profile** | personSkills, personNote, personCertifications, personProjects, personAwards, personAnniversaries, personWebSite | Skills, about me, certifications, projects, awards, birthday, personal site |
| **Identity** | personName, personCurrentPosition, personAddresses, personEmails, personPhones, personWebAccounts | Composite fields from Entra ID data |
| **Rich entities** | personInterests, personEducationalActivities, personLanguages, personPublications, personPatents | Full Microsoft Graph entity schemas with nested fields |

**Rich entities** accept both simple and complex input:
- CSV: `"MIT"` becomes `{"institution":{"displayName":"MIT"}}`
- JSON: Full `educationalActivity` with institution location, program details, and dates

### Custom Properties

Any column not in the [standard schema](./src/schema/user-property-schema.ts) becomes a searchable custom property automatically. If your CSV has `VTeam`, `CostCenter`, `BuildingAccess` — those become part of the connector schema at setup time.

**Schema limitation**: Connector schemas cannot be updated after registration. To add new columns, delete the old connector and create a new one with a new ID.

---

## Setup

### 1. Azure AD App Registration

Create an app in [Azure Portal](https://portal.azure.com) > Entra ID > App registrations.

**Platform configuration:**
- Add a **Single-page application (SPA)** redirect URI: `http://localhost:5544`
- This is the local auth server that handles the browser-based login flow for Option A

**Required permissions:**

| Permission | Type | Used by | Purpose |
|------------|------|---------|---------|
| `User.ReadWrite.All` | Delegated | Option A | Create and update users |
| `Directory.ReadWrite.All` | Delegated | Option A | Manage directory objects |
| `People.Read.All` | Delegated | Option A | Read profile data for enrichment |
| `Organization.Read.All` | Delegated | Option A | Read tenant info |
| `offline_access` | Delegated | Option A | Keep tokens refreshed |
| `ExternalConnection.ReadWrite.OwnedBy` | Application | Option B | Create/manage Graph Connector connections |
| `ExternalItem.ReadWrite.OwnedBy` | Application | Option B | Ingest items into connections |
| `PeopleSettings.ReadWrite.All` | Application | Option B | Register profile source + priority settings |

Grant admin consent after adding all permissions.

**Two auth flows, one app registration:**
- **Option A** (provisioning) uses delegated auth — opens a browser to `http://localhost:5544` for interactive login
- **Option B** (connector) uses application auth — reads `AZURE_CLIENT_SECRET` from `.env`, no browser needed

### 2. Configure Environment

```bash
cp .env.example .env
```

```bash
AZURE_TENANT_ID=your-tenant-id
AZURE_CLIENT_ID=your-client-id
AZURE_CLIENT_SECRET=your-client-secret
USER_DOMAIN=yourdomain.onmicrosoft.com

# Licenses: JSON input uses per-user Licenses array (display names, auto-resolved).
# LICENSE_SKU_IDS is the fallback for CSV input or JSON without Licenses field.
LICENSE_SKU_IDS=sku-id-1,sku-id-2
```

### 3. Install & Build

```bash
npm install && npm run build
```

---

## Commands

### Provisioning

```bash
npm run provision -- --csv config/team.csv           # Create users from CSV
npm run provision -- --json config/team.json         # Create users from JSON
npm run provision -- --dry-run --csv config/team.csv  # Preview changes
```

### Enrichment (Graph Connector)

```bash
# First time: create connection + schema + ingest
npm run option-b:setup -- --csv config/team.csv --connection-id m365people25

# With JSON (rich entity data)
npm run option-b:setup -- --json config/team.json --connection-id m365people25

# Merge mode: CSV base + JSON enrichment
npm run option-b:setup -- --csv config/team.csv --json config/rich.json --connection-id m365people25

# Re-ingest (connection exists)
npm run option-b:ingest -- --csv config/team.csv --connection-id m365people25

# Preview without changes
npm run option-b:dry-run -- --json config/team.json --connection-id m365people25
```

### Tenant Management

```bash
npm run list-users                      # List all users
npm run list-licenses                   # Show available licenses
npm run update-licenses                 # Add missing licenses
npm run reset-tenant                    # Preview cleanup (dry run)
npm run reset-tenant:confirm            # Delete all non-admin users
```

### Verification

```bash
node tools/debug/verify-ingestion-progress.mjs \
  --search-auth delegated \
  --connection-id m365people25 \
  --csv config/team.csv \
  --query "*"
```

Indexing takes 6+ hours. Profile data propagation takes 1-24 hours.

---

## Project Structure

```
src/
  provision.ts                 # Option A — create Entra ID users
  enrich-connector.ts          # Option B — Graph Connector pipeline
  json-loader.ts               # Shared JSON loader + PascalCase normalizer
  license-resolver.ts          # License display name → SKU ID resolver
  people-connector/
    connection-manager.ts      # Connection + profile source registration
    schema-builder.ts          # Schema with 18 labels + custom properties
    item-ingester.ts           # Entity serialization + ingestion
  schema/
    user-property-schema.ts    # Property routing (Option A vs B)
config/
  textcraft-europe.csv         # 95-person demo (CSV — TextCraft Europe)
  textcraft-europe.json        # 95-person demo (JSON, camelCase — TextCraft Europe)
  sample-rich.json             # 2-person sample (PascalCase, full entity depth)
docs/                          # Architecture, auth, state management, lessons
tools/                         # Debug and admin utilities
```

---

## What We Learned

This is a learning project. We documented everything:

| Document | What You'll Learn |
|----------|-------------------|
| [ARCHITECTURE-OPTION-A-B.md](./docs/ARCHITECTURE-OPTION-A-B.md) | Why Entra ID and profile enrichment are separate pipelines |
| [COPILOT-CONNECTORS-PEOPLE-DATA.md](./docs/COPILOT-CONNECTORS-PEOPLE-DATA.md) | End-to-end flow for making people data Copilot-searchable |
| [graph-connector-lessons.md](./docs/graph-connector-lessons.md) | 10 iterations of connector failures and what each taught us |
| [PEOPLE-DATA-AUTH-SOLUTION.md](./docs/PEOPLE-DATA-AUTH-SOLUTION.md) | Why Graph Connectors require Application permissions |
| [STATE-MANAGEMENT.md](./docs/STATE-MANAGEMENT.md) | Idempotent provisioning: CREATE, UPDATE, NOOP detection |

**Hard-won lessons:**

1. **Graph Connectors need Application auth** — Delegated tokens don't work for external items
2. **Profile source registration must happen before ingestion** — Otherwise data never appears on profile cards
3. **Schema is immutable** — Once registered, delete and recreate to change it
4. **Path A vs Path B deserialization** — Labeled properties accept any JSON type; unlabeled properties accept strings only
5. **PCP expects arrays where Graph docs say String** — `fieldsOfStudy`, `activities`, `awards` must be sent as arrays for downstream propagation
6. **Indexing is not instant** — Allow 6-24 hours before Copilot reflects new data
7. **Never refactor working connector code** — Even "harmless" changes correlated with ingestion failures

---

## License

MIT

---

## Contributing

Found something interesting? Learned something new about Graph Connectors or Copilot people data? Open a PR.
