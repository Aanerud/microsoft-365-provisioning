# Config Contract — users.connector.config.json

Contract between the data-science team (producer) and the ingest pipeline
(consumer). When this file drifts from what the ingest expects, users lose
data silently. Keep it current.

## The four config files

The project previously had one config file per user. It is now split by
concern — each file has a dedicated consumer and its own shape.

| File | Purpose | Primary consumer | Status |
|---|---|---|---|
| `config/users.config.json` | **Entra / Option A** — identity, licenses, job title, manager, office location, phone. Everything needed to create and license the user in Azure AD. | `src/provision.ts` (via `npm run provision`) | active |
| `config/users.connector.config.json` | **Graph Connector / Option B** — skills, interests, languages, certifications, projects, positions, products, territory tiers. Everything indexed by Copilot for people search. | `src/enrich-connector.ts` (via `npm run option-b:ingest`) | active |
| `config/users.connector.properties.config.json` | **Property definitions** — type (`string` or `collection`) + description for each custom property. Read at setup time to build the connector schema with descriptions and render collections as YAML. | `src/enrich-connector.ts` (via `loadPropertyDefinitions`) | active |
| `config/groups.config.json` | **Group memberships** — which Entra groups each user belongs to, plus group definitions (display name, mail, owners). | `src/update-groups.ts` (via `npm run update-groups`) | active |
| `config/contacts.config.json` | **External per-user contacts** — Outlook personal contacts with name, email, company, job title. | ⚠ **not yet wired** (no code consumer) | **incoming** |

### `users.config.json` vs `users.connector.config.json`

The split reflects Option A vs Option B in the architecture:

- **Option A (`users.config.json`)** feeds Entra ID directly. Fields here become
  identity properties — what you see in the Microsoft 365 admin center, what
  Outlook uses, what drives license billing. Changes here take effect on the
  next provisioning run.
- **Option B (`users.connector.config.json`)** feeds the Graph Connector, which
  sends labeled people data to PAPI (Profile API) where Microsoft 365 Copilot
  can find it. Fields here are about discoverability — "who knows React?" —
  not identity.

The same `MailNickName` is the join key between them. A user can exist in
Option A without Option B (no Copilot enrichment), but not the other way
around — Option B ingestion requires the user account to already exist.

### `groups.config.json`

Stand-alone file listing each group and its members. Runs after provisioning.
Not covered in this doc — see the file itself and `src/update-groups.ts`.

### `contacts.config.json` (incoming)

Contains each user's external Outlook contacts (customers, partners, vendors)
with full contact records (`displayName`, `givenName`, `surname`,
`companyName`, `jobTitle`, `emailAddresses`). Currently has no code consumer —
file sits in the config directory but nothing reads it.

**Planned:** a new `src/enrich-contacts.ts` pipeline that writes these as
Outlook personal contacts per user via the Graph `/users/{id}/contacts`
endpoint. Design and implementation deferred to a future PR.

---

The rest of this doc covers **`users.connector.config.json` only**.

## Record shape

Every record is a JSON object keyed by `MailNickName`. Top-level keys can be
PascalCase (the data-science team's native format) or camelCase — the ingest
auto-detects via `detectPascalCaseFormat()` and normalizes PascalCase to
camelCase before downstream code runs.

### Minimum required

```json
{
  "MailNickName": "nora.d"
}
```

### Recommended baseline

```json
{
  "MailNickName": "nora.d",
  "PhoneNumber": "+47-12345678",
  "Skills": [ { "DisplayName": "Leadership", "Proficiency": "expert" } ],
  "Languages": [ { "DisplayName": "Norwegian" } ],
  "Interests": [ { "DisplayName": "Strategy" } ],
  "Certifications": [ { "DisplayName": "ITIL v4" } ],
  "Projects": [ { "DisplayName": "Aurora Launch" } ],
  "Positions": [
    {
      "detail": {
        "jobTitle": "CEO",
        "company": { "displayName": "Contoso" }
      },
      "isCurrent": true
    }
  ]
}
```

### Full column list (currently consumed)

| Column | Type | Consumed by | Notes |
|---|---|---|---|
| `MailNickName` | string | all | Required. Used to build UPN + lookup OID. |
| `PhoneNumber` | string | ingester | |
| `Anniversaries` | array | ingester | `{ type, date }` |
| `Awards` | array | ingester | `{ DisplayName }` |
| `Certifications` | array | ingester | `{ DisplayName }` |
| `EducationalActivities` | array | ingester | rich PCP entity |
| `Emails` | array | ingester | |
| `Interests` | array | ingester | `{ DisplayName }` |
| `Skills` | array | ingester | `{ DisplayName, Proficiency }` |
| `Languages` | array | ingester | |
| `WebAccounts` | array | ingester | |
| `Notes` | array | ingester | |
| `Patents` | array | ingester | |
| `Phones` | array | ingester | |
| `Positions` | array | ingester | composite — renders `personCurrentPosition` |
| `Projects` | array | ingester | |
| `Publications` | array | ingester | |
| `Responsibilities` | array | ingester | |
| `Websites` | array | ingester | |
| `Products` | array of `{Name, Model, GTIN}` | renderer → custom prop | see **Products rendering** below |
| `TerritoryTier1/2/3` | string | custom prop | |

Any additional top-level string column is auto-detected via
`getCustomProperties()` and registered as a connector custom property
(`string` type, Path B). See **Dynamic custom properties** below.

## Position-scoped fields (WorkModality, JobFamilyGroupName, JobFamilyName)

These are **per-position** properties. A person can hold multiple roles over
time, and each role has its own modality and job-family classification. They
live inside each position rather than at the top level:

```json
"Positions": [
  {
    "Detail": {
      "JobTitle": "CEO",
      "WorkModality": "Hybrid",
      "JobFamilyGroupName": "Engineering",
      "JobFamilyName": "Software Engineering"
    },
    "isCurrent": true
  }
]
```

Coverage in current input:

| Field | Users | Scope |
|---|---|---|
| `Detail.WorkModality` | all users | all users |
| `Detail.JobFamilyGroupName` | varies | users with an engineering job family |
| `Detail.JobFamilyName` | varies | users with an engineering job family |

These flow through to the connector as part of the `personCurrentPosition`
labeled property (Path A — searchable by Copilot), not as flat top-level
custom properties.

### Why not flat

The old test tenant exposed these as flat top-level custom properties
(`workModality`, `jobFamilyGroupName`, `jobFamilyName`). That pattern was
denormalized — it couldn't represent multiple positions per user without
losing information. Our current pipeline keeps them inside their position,
where they belong structurally.

**If you're reading a wave2-old-era diff report that flags these as "missing"
at the top level, that's not a data loss — it's a shape change. The values
are present, just inside `positions[].detail.*`.**

## Properties config — `users.connector.properties.config.json`

Sidecar file that defines custom property metadata. The pipeline reads it
automatically (same directory, derived filename).

```json
[
  { "name": "Products",       "type": "collection", "description": "List of products associated with the user" },
  { "name": "WorkModality",   "type": "string",     "description": "The work modality of the user" },
  { "name": "TerritoryTier1", "type": "string",     "description": "The territory information for the user, level 1" }
]
```

### `type` field

- **`string`** — value passes through as-is. `"Hybrid"` stays `"Hybrid"`.
- **`collection`** — value is an array of objects in the data file. The tool
  renders it as YAML (recommended by
  [Microsoft docs](https://learn.microsoft.com/en-us/microsoft-365/copilot/extensibility/build-connectors-with-people-data)
  for complex custom properties). All object keys are lowercased.

### `description` field

Included in the connector schema at setup time:
```json
{ "name": "products", "type": "string", "description": "List of products associated with the user" }
```
Descriptions on labeled properties (person* labels) are ignored by the API.
Only custom properties benefit.

If descriptions are missing, the tool warns before schema creation and asks
the operator to confirm.

### Collection rendering example (Products)

Input (from data file):
```json
"Products": [
  { "Name": "Echo Vault",   "Model": "Model A", "GTIN": "96128715782278" },
  { "Name": "Falcon Ridge", "Model": "Model Y", "GTIN": "95441323966177" }
]
```

Output (YAML string sent to connector):
```yaml
- name: Echo Vault
  model: Model A
  gtin: 96128715782278
- name: Falcon Ridge
  model: Model Y
  gtin: 95441323966177
```

The schema `description` tells Copilot what the field means. The value is
pure data. No headers, no embedded metadata.

### Size cap

Output per collection field is capped at 8KB. Truncated with `[TRUNCATED]`.

### Adding a new collection property

1. Data-science team adds the array-of-objects field to user records.
2. Add an entry with `"type": "collection"` to the properties config.
3. The tool automatically renders it as YAML. No code change needed.

## Dynamic custom properties

Columns that aren't in the hardcoded schema are auto-detected by
`getCustomProperties()` in `src/schema/user-property-schema.ts`. When a new
column appears in the config:

1. It becomes a connector schema property (type `string`) on the NEXT new
   connection.
2. Existing connections do NOT pick it up — schema is immutable (see below).
3. Any typo in a column name also becomes a real property. **No whitelist or
   warning today** — this is a known risk, tracked as a follow-up.

## Schema immutability

Connector schemas cannot be updated in place. Adding a column to the config
and re-running `option-b:ingest` will either silently ignore the new column
or return HTTP 400 from the schema PATCH.

**To deploy a column change:**

1. Add the column to `users.connector.config.json`.
2. Delete the current connection (`npm run enrich:delete-connector`).
3. Create a new connection with a new ID (`m365people25`, next after `m365people24`).
4. The new schema will include the new column.

This is a memory-encoded constraint — see
`docs/graph-connector-lessons.md` for history and validated incidents
(m365people14–23).

## Validation

Run the content-level parity check against the live tenant:

```bash
npm run verify:papi-parity                             # all users
node tools/debug/verify-papi-parity.mjs --user nora.d  # single user
```

Outputs per-field categories:

- `PASS` — input value present in output
- `OUTPUT_MISSING` — input had it, output doesn't (real regression, exit 1)
- `SHAPE_COLLAPSE` — same key, degraded shape (e.g. Products losing Model/GTIN)
- `INPUT_MISSING` — column not in input at all (documented upstream gap)
- `PROFILE_MISSING` — no PAPI profile for this user

Exit code 0 iff zero `OUTPUT_MISSING` and zero `SHAPE_COLLAPSE`.

## Contract change process

When the data-science team wants to add a new column:

1. Open a PR adding the column to a test config.
2. Run `npm run option-b:dry-run --json <test-config>` — confirm the column
   appears in the emitted schema.
3. Update this doc's **Full column list** table.
4. Coordinate a connection rotation (new connection ID) with the ingest owner.
5. After ingest, run `verify-papi-parity` and confirm the new column is `PASS`
   for the users that have it set.
