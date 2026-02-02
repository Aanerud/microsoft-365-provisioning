# State Management System

## Overview

The M365 Agent Provisioning tool now operates as a **declarative state management system** where the CSV file is the **source of truth**. Instead of only creating users, the tool now automatically syncs Azure AD to match the CSV exactly.

## Core Concepts

### CSV as Source of Truth

The CSV file defines the desired state of your Azure AD users. When you run the provisioning tool:

1. **CREATE**: Users that exist in CSV but not in Azure AD are created
2. **UPDATE**: Users that exist in both but have different attributes are updated
3. **DELETE**: Users that exist in Azure AD but not in CSV are deleted (with confirmation)

### State Operations

```
┌─────────────────────────────────────────────────────┐
│                  CSV File (Desired State)           │
│  - 12 users defined with all properties             │
└──────────────────┬──────────────────────────────────┘
                   │
                   ▼
┌─────────────────────────────────────────────────────┐
│             State Manager (Delta Calculation)       │
│  - Fetches current Azure AD state                   │
│  - Compares CSV vs Azure AD                         │
│  - Detects changes in 50+ standard properties       │
│  - Detects changes in custom properties             │
│  - Generates action plan (CREATE/UPDATE/DELETE)     │
└──────────────────┬──────────────────────────────────┘
                   │
                   ▼
┌─────────────────────────────────────────────────────┐
│              Azure AD (Current State)               │
│  - Users created, updated, or deleted               │
│  - State now matches CSV exactly                    │
└─────────────────────────────────────────────────────┘
```

## Comprehensive Property Support

### Standard Properties (Option A only)

Option A only writes properties marked `handledBy=optionA` in the schema. Any Option B properties
(skills, aboutMe, languages, interests, etc.) are ignored here and handled by the Option B pipeline.

**Basic Info (4 properties)**
- displayName, givenName, surname, accountEnabled

**Contact (6 properties)**
- mail, mailNickname, mobilePhone, businessPhones, otherMails, faxNumber

**Address (6 properties)**
- city, state, country, postalCode, streetAddress, officeLocation

**Job Info (10 properties)**
- jobTitle, department, companyName, employeeId, employeeType, employeeHireDate, employeeLeaveDateTime, hireDate, employeeOrgData

**Identity (3 properties)**
- userPrincipalName, userType, onPremisesImmutableId

**Preferences (4 properties)**
- usageLocation, preferredLanguage, preferredDataLocation, mailboxSettings

**Personal (Option B / future)**
- birthday, interests, skills, schools, pastProjects, responsibilities, mySite
- languages (future connector support)

**Security (2 properties)**
- passwordPolicies, passwordProfile

**On-Premises (2 properties)**
- onPremisesExtensionAttributes, onPremisesImmutableId

**Legal/Compliance (2 properties)**
- ageGroup, consentProvidedForMinor

### Custom Properties (Deferred to Option B)

Any CSV column **not in the standard schema** is treated as Option B enrichment and is **not written** by Option A.
Option B can ingest labeled fields as its schema expands (e.g., languages when labels become available).

**Example CSV with Custom Properties:**

```csv
name,email,jobTitle,DeploymentManager,ProjectCode,FavoriteColor
John Doe,john@domain.com,Engineer,true,ENG-001,Blue
Jane Smith,jane@domain.com,Manager,false,MGR-002,Green
```

**Standard properties**: name, email, jobTitle
**Custom properties**: DeploymentManager, ProjectCode, FavoriteColor (Option B)

Custom properties are ingested via the Option B connector instead of being written to the Entra ID user object.

## Usage Examples

### Basic Sync (Full State Management)

```bash
# Preview what would change (dry-run)
npm run provision -- --dry-run

# Apply changes with delete confirmation
npm run provision

# Apply changes including deletions (skip confirmation)
npm run provision -- --force
```

### Selective Sync

```bash
# Only CREATE and UPDATE (don't delete)
npm run provision -- --skip-delete

# Only CREATE new users (skip updates and deletes)
npm run provision -- --skip-update --skip-delete

# Only UPDATE existing users (no create or delete)
npm run provision -- --skip-create --skip-delete
```

### Verbose Output

```bash
# Show detailed diff report
npm run provision -- --dry-run

# Show diff even in live mode
npm run provision -- --show-diff
```

## Diff Report

When running with `--dry-run` or `--show-diff`, the tool generates a detailed report:

```
═══════════════════════════════════════════════════════════════
Provisioning State Changes
═══════════════════════════════════════════════════════════════

Generated: 2026-01-22T15:30:00Z

Summary
───────────────────────────────────────────────────────────────
  Total in CSV:            12 users
  Total in Azure AD:       10 users
  To CREATE:               3 users
  To UPDATE:               2 users
  To DELETE:               1 user
  Unchanged:               6 users
  Custom properties:       2 (DeploymentManager, ProjectCode)

Users to CREATE
───────────────────────────────────────────────────────────────
1. Sarah Chen (sarah.chen@domain.com)
     jobTitle: Chief Executive Officer
     department: Executive
     city: Seattle
     Custom Properties:
       DeploymentManager: true (custom)
       ProjectCode: EXEC-001 (custom)

Users to UPDATE
───────────────────────────────────────────────────────────────
1. Michael Rodriguez (michael.rodriguez@domain.com)
     jobTitle: "CTO" → "Chief Technology Officer"
     city: null → "Seattle"
     DeploymentManager: false → true (custom)

Users to DELETE
───────────────────────────────────────────────────────────────
⚠️  WARNING: These users exist in Azure AD but not in CSV

1. Old User (old.user@domain.com)
     Had custom properties:
       - DeploymentManager
       - ProjectCode

═══════════════════════════════════════════════════════════════
```

## Safety Features

### Delete Confirmation

By default, DELETE operations require explicit confirmation:

```bash
$ npm run provision

⚠️  WARNING: 5 users will be DELETED!
Users to delete:
  - Old User (old.user@domain.com)
  - Test User (test@domain.com)
  ...

Run with --force flag to proceed with deletion.
Run with --skip-delete to only CREATE and UPDATE users.
```

### Skip Delete Flag

Prevent accidental deletions by using `--skip-delete`:

```bash
npm run provision -- --skip-delete
```

This ensures only CREATE and UPDATE operations are performed.

### Dry-Run Mode

Always preview changes before applying:

```bash
npm run provision -- --dry-run
```

This shows exactly what would change without modifying Azure AD.

## Workflow Examples

### Initial Setup (Empty Azure AD)

```bash
# Start: 0 users in Azure AD
# CSV: 12 users

$ npm run provision -- --dry-run
# Shows: CREATE 12 users

$ npm run provision
# Creates all 12 users
```

### Update Scenario

```bash
# Start: 12 users exist
# CSV: Changed jobTitle for 2 users

$ npm run provision -- --dry-run
# Shows: UPDATE 2 users (jobTitle changes)

$ npm run provision
# Updates 2 users in Azure AD
```

### Add and Remove Users

```bash
# Start: 12 users exist
# CSV: Add 2 new users, remove 1 old user

$ npm run provision -- --dry-run
# Shows: CREATE 2, DELETE 1, UNCHANGED 11

$ npm run provision
# Prompts for delete confirmation

$ npm run provision -- --force
# Creates 2, deletes 1, total now 13 users
```

### Custom Property Updates

```bash
# Start: Users have DeploymentManager: false
# CSV: Change DeploymentManager to true for 5 users

$ npm run provision -- --dry-run
# Shows: UPDATE 5 users
#   DeploymentManager: false → true (custom)

$ npm run provision
# Custom properties are deferred to Option B (Graph Connector)
```

## Technical Details

### Change Detection

The state manager performs intelligent comparison:

**Type-Aware Comparison:**
- Strings: case-sensitive exact match
- Numbers: numeric equality
- Booleans: boolean equality
- Arrays: sorted comparison (order-independent)
- Dates: timestamp equality
- Objects: deep JSON comparison

**Null Handling:**
- Empty strings in CSV are treated as "not set"
- null vs undefined are considered equivalent
- Removing a CSV column removes the property

### Batch Operations

All operations use Microsoft Graph batch API:

- **Batch size**: 20 requests per batch
- **Rate limiting**: 500ms delay between batches
- **Parallel processing**: Multiple batches for large datasets
- **Error handling**: Individual failures don't block the batch

### Custom Property Storage

Custom properties are handled by Option B via the Graph Connector schema, not written to the Entra ID user object.

## CLI Reference

### Core Commands

```bash
# Full sync with all operations
npm run provision

# Dry-run (preview only)
npm run provision -- --dry-run

# Force delete without confirmation
npm run provision -- --force
```

### Selective Sync Flags

```bash
# Skip delete operations
npm run provision -- --skip-delete

# Skip update operations
npm run provision -- --skip-update

# Skip create operations
npm run provision -- --skip-create
```

### Other Flags

```bash
# Show detailed diff
npm run provision -- --show-diff

# Use beta API endpoints
npm run provision -- --use-beta

# Skip license assignment
npm run provision -- --skip-licenses

# Custom CSV file
npm run provision -- --csv path/to/file.csv

# Force re-authentication
npm run provision -- --auth
```

## CSV Format

### Required Columns

- `name` - Display name (used as displayName)
- `email` - Email address (used as userPrincipalName)
- `role` - Local role field (not synced to Azure AD)
- `department` - Department (synced to Azure AD)

### Standard Property Columns

Any column name matching a standard Microsoft Graph property:

- `jobTitle` - Job title
- `city` - City
- `state` - State/province
- `country` - Country
- `usageLocation` - Two-letter country code (required for licenses)
- `employeeType` - Employee, Contractor, etc.
- `companyName` - Company name
- `officeLocation` - Office location
- ... and 40+ more properties

### Custom Property Columns

Any column **not** in the standard schema is treated as Option B enrichment and ignored by Option A:

- `DeploymentManager` - Custom field (Option B)
- `ProjectCode` - Custom field (Option B)
- `CostCenter` - Custom field (Option B)
- `TeamColor` - Custom field (Option B)
- ... unlimited custom fields

### Example CSV

```csv
name,email,jobTitle,department,city,usageLocation,DeploymentManager,ProjectCode
John Doe,john@domain.com,Engineer,Engineering,Seattle,US,true,ENG-001
Jane Smith,jane@domain.com,Manager,Engineering,Portland,US,false,MGR-002
Bob Wilson,bob@domain.com,Designer,Design,Seattle,US,false,DES-003
```

## Output Format

### Enhanced AgentConfig

The exported configuration now includes state tracking:

```json
{
  "agents": [
    {
      "name": "Sarah Chen",
      "email": "sarah.chen@domain.com",
      "role": "CEO",
      "department": "Executive",
      "userId": "azure-ad-user-id",
      "password": "generated-password",
      "jobTitle": "Chief Executive Officer",
      "city": "Seattle",
      "usageLocation": "US",
      "lastAction": "CREATE",
      "lastModified": "2026-01-22T15:30:00Z",
      "changedFields": [],
      "createdAt": "2026-01-22T15:30:00Z"
    }
  ],
  "summary": {
    "totalAgents": 1,
    "successfulProvisions": 1,
    "failedProvisions": 0,
    "generatedAt": "2026-01-22T15:30:00Z"
  }
}
```

## Migration from Old Behavior

### Before (Create-Only)

```bash
# Old behavior: Always created new users
npm run provision

# Would fail if users already existed
# No update capability
# No delete capability
```

### After (State Management)

```bash
# New behavior: Syncs to CSV
npm run provision -- --dry-run  # Preview changes first

# Creates new users
# Updates existing users
# Deletes users not in CSV (with confirmation)
```

### Backward Compatibility

The tool remains backward compatible:

- Existing CSV files work without modification
- Output format enhanced but compatible
- Old CLI flags still work
- No breaking changes to API

## Best Practices

### 1. Always Dry-Run First

```bash
npm run provision -- --dry-run
```

Preview changes before applying to catch mistakes.

### 2. Use Skip-Delete for Safety

```bash
npm run provision -- --skip-delete
```

Prevent accidental deletions while developing CSV.

### 3. Use Version Control for CSV

```bash
git add config/agents-template.csv
git commit -m "Add 3 new engineers"
```

Track changes to your user definitions.

### 4. Test with Small Batches

Start with a few users, verify, then scale up.

### 5. Set Usage Location

Always include `usageLocation` column for license assignment:

```csv
name,email,usageLocation
John Doe,john@domain.com,US
```

### 6. Use Custom Properties Wisely

Document your custom properties (used by Option B connector):

```csv
# Custom properties:
# - DeploymentManager: boolean, whether user is a deployment manager
# - ProjectCode: string, assigned project code
# - CostCenter: string, cost center for billing
name,email,DeploymentManager,ProjectCode,CostCenter
```

## Troubleshooting

### "User will be DELETED" Warning

**Cause**: User exists in Azure AD but not in CSV

**Solutions**:
1. Add user back to CSV if they should exist
2. Use `--skip-delete` to preserve user
3. Use `--force` to proceed with deletion

### No Changes Detected

**Cause**: CSV matches Azure AD exactly

**Solution**: Verify CSV has the changes you expect

### Custom Properties Not Updating

**Cause**: Beta endpoints might be unavailable

**Solution**: Ensure `USE_BETA_ENDPOINTS=true` in `.env`

### License Assignment Fails

**Cause**: `usageLocation` not set

**Solution**: Add `usageLocation` column to CSV with two-letter country code

## Performance

### Batch Operations

- **CREATE**: 20 users per batch
- **UPDATE**: 20 users per batch
- **DELETE**: 20 users per batch
- **Extension operations**: 20 per batch

### Timing Examples

- 12 users (no changes): ~5 seconds
- 12 users (all new): ~10 seconds
- 100 users (all new): ~45 seconds
- 1000 users (all new): ~7 minutes

### Rate Limiting

Microsoft Graph API limits:
- 40 requests per second per application
- Tool adds 500ms delay between batches
- Automatic retry on throttling

## Related Documentation

- [Microsoft Graph User Resource](https://learn.microsoft.com/en-us/graph/api/resources/user?view=graph-rest-beta)
- [Graph Connectors](https://learn.microsoft.com/en-us/graph/connecting-external-content-connectors)
- [JSON Batching](https://learn.microsoft.com/en-us/graph/json-batching)
- [Throttling Limits](https://learn.microsoft.com/en-us/graph/throttling-limits)
