# Make Copilot Know Your People

Copilot cannot answer questions about people it knows nothing about. Ask *"Who speaks French and has an MBA?"* — silence. The profiles are empty. No skills, no education, no languages. Copilot sees org chart boxes, not people.

This tool fills the profiles. Two commands turn a CSV or JSON file into rich, searchable people data across Microsoft 365 and Copilot.

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
Copilot answers:
  "Who has Python skills and speaks German?"
  "Find someone with a patent in distributed systems"
  "Who on the Stockholm team has an MBA?"
```

> Built to explore Microsoft Graph APIs, OAuth 2.0 flows, and Graph Connectors. Everything we learned — including the mistakes — is in [docs/](./docs/).

---

## Why This Matters

Microsoft 365 stores people data in 20 profile collections, defined in the [Beta Profile API](https://learn.microsoft.com/en-us/graph/api/resources/profile?view=graph-rest-beta). Each collection describes one facet of a person — their skills, languages, education, patents, and so on.

**Empty profiles cripple Copilot.** When these collections hold no data, Copilot cannot match people to questions. It cannot find the French speaker, the patent holder, the certified architect. The intelligence is there; the data is not.

This tool writes to **19 of the 20 collections**. One person at a time, one collection at a time, until Copilot has what it needs to give real answers about real people.

### The 20 Profile Collections

Every collection maps to a [Microsoft Graph Beta Profile](https://learn.microsoft.com/en-us/graph/api/resources/profile?view=graph-rest-beta) relationship. This tool supports 19:

| # | Collection | Graph Type | Description | Supported |
|---|------------|------------|-------------|-----------|
| 1 | [`accounts`](https://learn.microsoft.com/en-us/graph/api/resources/useraccountinformation?view=graph-rest-beta) | userAccountInformation | User account identity (UPN, directory object ID) | Yes |
| 2 | [`addresses`](https://learn.microsoft.com/en-us/graph/api/resources/itemaddress?view=graph-rest-beta) | itemAddress | Physical addresses (street, city, country, postal code) | Yes |
| 3 | [`anniversaries`](https://learn.microsoft.com/en-us/graph/api/resources/personanniversary?view=graph-rest-beta) | personAnniversary | Birthday, work anniversary, and other dates | Yes |
| 4 | [`awards`](https://learn.microsoft.com/en-us/graph/api/resources/personaward?view=graph-rest-beta) | personAward | Honors and awards received | Yes |
| 5 | [`certifications`](https://learn.microsoft.com/en-us/graph/api/resources/personcertification?view=graph-rest-beta) | personCertification | Professional certifications (PMP, AWS, etc.) | Yes |
| 6 | [`educationalActivities`](https://learn.microsoft.com/en-us/graph/api/resources/educationalactivity?view=graph-rest-beta) | educationalActivity | Degrees, programs, institutions, fields of study | Yes |
| 7 | [`emails`](https://learn.microsoft.com/en-us/graph/api/resources/itememail?view=graph-rest-beta) | itemEmail | Email addresses with type labels | Yes |
| 8 | [`interests`](https://learn.microsoft.com/en-us/graph/api/resources/personinterest?view=graph-rest-beta) | personInterest | Personal and professional interests | Yes |
| 9 | [`languages`](https://learn.microsoft.com/en-us/graph/api/resources/languageproficiency?view=graph-rest-beta) | languageProficiency | Languages with reading, spoken, and written proficiency | Yes |
| 10 | [`names`](https://learn.microsoft.com/en-us/graph/api/resources/personname?view=graph-rest-beta) | personName | Display name, given name, surname | Yes |
| 11 | [`notes`](https://learn.microsoft.com/en-us/graph/api/resources/personannotation?view=graph-rest-beta) | personAnnotation | About Me / freeform notes | Yes |
| 12 | [`patents`](https://learn.microsoft.com/en-us/graph/api/resources/itempatent?view=graph-rest-beta) | itemPatent | Patents with filing details and status | Yes |
| 13 | [`phones`](https://learn.microsoft.com/en-us/graph/api/resources/itemphone?view=graph-rest-beta) | itemPhone | Phone numbers with type (mobile, business) | Yes |
| 14 | [`positions`](https://learn.microsoft.com/en-us/graph/api/resources/workposition?view=graph-rest-beta) | workPosition | Current and past job positions | Yes |
| 15 | [`projects`](https://learn.microsoft.com/en-us/graph/api/resources/projectparticipation?view=graph-rest-beta) | projectParticipation | Project history and contributions | Yes |
| 16 | [`publications`](https://learn.microsoft.com/en-us/graph/api/resources/itempublication?view=graph-rest-beta) | itemPublication | Books, articles, and published works | Yes |
| 17 | [`skills`](https://learn.microsoft.com/en-us/graph/api/resources/skillproficiency?view=graph-rest-beta) | skillProficiency | Skills with proficiency levels and categories | Yes |
| 18 | [`webAccounts`](https://learn.microsoft.com/en-us/graph/api/resources/webaccount?view=graph-rest-beta) | webAccount | LinkedIn, GitHub, and other web accounts | Yes |
| 19 | [`websites`](https://learn.microsoft.com/en-us/graph/api/resources/personwebsite?view=graph-rest-beta) | personWebsite | Personal and professional websites | Yes |
| 20 | [`responsibilities`](https://learn.microsoft.com/en-us/graph/api/resources/personresponsibility?view=graph-rest-beta) | personResponsibility | Job responsibilities and duties | Not yet |

> `responsibilities` is not yet supported by the Microsoft 365 people data ingestion pipeline. When the `personResponsibilities` semantic label becomes available, this tool will pick it up.

### How Collections Become Copilot-Searchable

Collections alone do not make data searchable. Microsoft 365 Copilot searches people through **Graph Connectors with people data labels** — a mapping layer that tells Copilot which connector properties correspond to which profile facets.

This tool registers **18 people data labels** across three groups:

| Group | Labels | What Copilot Can Search |
|-------|--------|-------------------------|
| **Core** | personSkills, personNote, personCertifications, personProjects, personAwards, personAnniversaries, personWebSite | Skills, about me, certifications, projects, awards, birthday, personal site |
| **Composite** | personName, personCurrentPosition, personAddresses, personEmails, personPhones, personWebAccounts | Identity and contact data composed from Entra ID fields |
| **Rich entities** | personInterests, personEducationalActivities, personLanguages, personPublications, personPatents | Full entity schemas with nested fields (institutions, proficiency levels, filing details) |

**Rich entities** accept both simple and complex input:
- CSV: `"MIT"` becomes `{"institution":{"displayName":"MIT"}}`
- JSON: Full `educationalActivity` with institution, program, fields of study, and dates

### Custom Properties

Any column not in the [standard schema](./src/schema/user-property-schema.ts) becomes a searchable custom property. If your CSV has `VTeam`, `CostCenter`, `BuildingAccess` — those join the connector schema at setup time.

**Schema limitation**: Connector schemas cannot change after registration. To add columns, delete the old connector and create a new one.

---

## Two Minutes to Working Profiles

### 1. A CSV — the simplest case

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

### 2. Add enrichment

More columns, more data. The tool routes each one:

```csv
name,email,role,department,skills,languages,aboutMe,certifications,VTeam
Sarah Chen,sarah@domain.com,CEO,Executive,"['Leadership','Strategy']","['English (Native)']",20 years in tech,['PMP'],Executive
```

```bash
npm run provision -- --csv config/team.csv
npm run option-b:setup -- --csv config/team.csv --connection-id m365people25
```

Copilot now knows Sarah has leadership skills and a PMP certification.

### 3. Rich profiles with JSON

CSV handles flat data. JSON unlocks the full depth of the [Profile API schema](https://learn.microsoft.com/en-us/graph/api/resources/profile?view=graph-rest-beta) — education with institutions and programs, languages with proficiency levels, patents with filing details, per-user license assignment:

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

**JSON input features:**
- `MailNickName` + `USER_DOMAIN` from `.env` builds the full email
- `Licenses` assigns per-user licenses by display name (resolved to SKU IDs)
- `Manager` references other users by `MailNickName`
- Rich entities match the [Microsoft Graph Profile API](https://learn.microsoft.com/en-us/graph/api/resources/profile?view=graph-rest-beta) schema
- Custom org fields (`TerritoryTier1`, `Products`) become searchable connector properties

**CSV and JSON work together.** JSON overrides CSV per-property when both are provided:

```bash
npm run option-b:setup -- --csv config/team.csv --json config/rich-profiles.json --connection-id m365people01
```

A plain string like `"Leadership"` works the same as a full object with `CollaborationTags` and `Categories`. No property is forced to be rich.

---

## How Data Flows

Every field routes to where it belongs:

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

Edit `.env` with your tenant credentials:

```bash
AZURE_TENANT_ID=your-tenant-id          # Azure Portal > Entra ID > Overview
AZURE_CLIENT_ID=your-client-id          # App Registrations > Your App
AZURE_CLIENT_SECRET=your-client-secret  # Required for Option B (connector)
USER_DOMAIN=yourdomain.onmicrosoft.com  # Your tenant domain

# Licenses: JSON input uses per-user Licenses array (display names, auto-resolved).
# LICENSE_SKU_IDS is the fallback for CSV input or JSON without Licenses field.
# Run 'npm run list-licenses' to see available SKUs.
LICENSE_SKU_IDS=
```

**What each option needs from `.env`:**

| Variable | Option A | Option B | Notes |
|----------|----------|----------|-------|
| `AZURE_TENANT_ID` | Required | Required | |
| `AZURE_CLIENT_ID` | Required | Required | |
| `AZURE_CLIENT_SECRET` | Not needed | **Required** | Application auth for connector |
| `USER_DOMAIN` | Required | Required | Constructs UPN from MailNickName |
| `LICENSE_SKU_IDS` | Fallback | Not used | JSON `Licenses` array takes priority |

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

Option B maps each connector item to an Entra ID user via their Object ID. This mapping is stored in an OID cache file. If you ran Option A first, the cache already exists. If you're running Option B standalone, build it first:

```bash
# Build OID cache (required before Option B if you haven't run Option A)
npm run build-oid-cache -- --csv config/team.json
```

Then run the connector pipeline:

```bash
# First time: create connection + schema + ingest
npm run option-b:setup -- --json config/team.json --connection-id m365people01

# Re-ingest (connection exists)
npm run option-b:ingest -- --json config/team.json --connection-id m365people01

# Preview without changes
npm run option-b:dry-run -- --json config/team.json --connection-id m365people01

# Merge mode: CSV base + JSON enrichment
npm run option-b:setup -- --csv config/team.csv --json config/rich.json --connection-id m365people01
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

After ingestion, data does NOT appear on profiles immediately. Microsoft's internal pipeline (CAPIv2) processes items in batches over multiple export cycles. Expect:

- **6-24 hours**: First items start appearing on profiles
- **24-48 hours**: Most items propagated
- **`profileSyncEnabled=False`**: Normal on initial cycles, flips to True over time

Track propagation with the profile matrix tool:

```bash
# Show which users have connector data on their profiles
node tools/debug/profile-matrix.mjs --json config/team.json --connection-id m365people01
```

Output:
```
Name                          SKL INT CRT AWD PRJ EDU LNG PUB PAT NAM POS ADR EML PHN NTE  Source
─────────────────────────────────────────────────────────────────────────────────────────────────
Nora Dahl                      10  ·   1   ·   ·   1   2   ·   ·   1   1   ·   1   2   ·   CONN
Sofia Johansson                 ·  ·   ·   ·   ·   ·   ·   ·   ·   .   .   ·   .   .   ·   aad
─────────────────────────────────────────────────────────────────────────────────────────────────
2/94 users have connector data (2%)
Legend: N = items from connector, . = data from other source, · = empty
```

Run this daily after ingestion. The numbers fill in gradually as Microsoft processes each batch. If a user shows `CONN`, their skills, education, languages, etc. are now Copilot-searchable.

You can also verify individual items and the schema:

```bash
# Check a specific user's profile
node tools/debug/fetch-user-profile.mjs nora.d@yourdomain.onmicrosoft.com

# Verify ingestion via Microsoft Search
node tools/debug/verify-ingestion-progress.mjs \
  --search-auth delegated \
  --connection-id m365people01 \
  --csv config/team.json \
  --query "*"
```

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
  textcraft-europe.json        # 95-person demo (JSON, PascalCase — TextCraft Europe)
  sample-rich.json             # 2-person sample (PascalCase, full entity depth)
tools/debug/
  profile-matrix.mjs           # Track connector data propagation across all users
  fetch-user-profile.mjs       # Full profile dump for a single user
  verify-ingestion-progress.mjs # Compare ingested items vs search index
  check-people-connector-status.mjs # Connector health: schema, source, priority
  check-app-permissions.mjs    # Verify app registration permissions
  run-checklist.mjs            # Run all debug checks in sequence
docs/                          # Architecture, auth, state management, lessons
```

### Debug Tools

Six tools for diagnosing connector propagation issues. All require a valid token (`npm run test-connection` to refresh).

```bash
# Track propagation across all users — run daily after ingestion
node tools/debug/profile-matrix.mjs --json config/team.json --connection-id m365people01

# Full profile dump for one user (all collections, sources, dates)
node tools/debug/fetch-user-profile.mjs nora.d@yourdomain.onmicrosoft.com

# Connector health check (schema, profile source, prioritization)
node tools/debug/check-people-connector-status.mjs --connection-id m365people01

# Verify app registration has required permissions
node tools/debug/check-app-permissions.mjs --connection-id m365people01

# Compare ingested items vs Microsoft Search index
node tools/debug/verify-ingestion-progress.mjs --connection-id m365people01 --csv config/team.json

# Run all checks at once
node tools/debug/run-checklist.mjs --connection-id m365people01
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
5. **The ingestion pipeline expects arrays where Graph docs say String** — `fieldsOfStudy`, `activities`, `awards` must be sent as arrays for downstream propagation
6. **Indexing is not instant** — Allow 6-24 hours before Copilot reflects new data
7. **Never refactor working connector code** — Even "harmless" changes correlated with ingestion failures

---

## License

MIT

---

## Contributing

Found something interesting? Learned something new about Graph Connectors or Copilot people data? Open a PR.
