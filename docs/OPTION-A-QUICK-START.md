# Option A Test - Quick Start Guide

## What We're Testing

We're pushing the standard `/users` endpoint to its limits by testing **Option A properties only** (no Option B enrichment fields):

### Properties Being Tested
- **Arrays**: `otherMails`
- **Strings**: `faxNumber`
- **Preferences**: `preferredDataLocation`

> Option B owns personal enrichment fields like `skills`, `aboutMe`, `interests`, and `languages`.

## Test Dataset

**File**: `config/agents-test-maxprops.csv`

**Users**: 3 test users with comprehensive profiles:
- **Test User Alpha** - Engineer with standard properties only
- **Test User Beta** - UX Designer with standard properties only
- **Test User Gamma** - Financial Analyst with standard properties only

## Run the Test

### Step 1: Dry-Run (Preview Changes)
```bash
npm run provision -- --csv config/agents-test-maxprops.csv --dry-run
```

**What to look for**:
- Should show "CREATE: 3 users"
- Check that it detects all the new properties
- No errors should appear

### Step 2: Create the Test Users
```bash
npm run provision -- --csv config/agents-test-maxprops.csv
```

**What should happen**:
```
🔍 Fetching current Azure AD state...
📊 Calculating changes...

Summary:
- Total in CSV: 3 users
- To CREATE: 3 users
- Custom columns detected: 3 (VTeam, BenefitPlan, ... — ignored by Option A)

📦 Applying changes...
  Creating 3 users...
  ✓ Created 3 users
  Assigning 2 manager relationships...
  ✓ Assigned 2 managers

✅ Provisioning complete!
```

### Step 3: Verify in Azure AD Portal

1. Go to [Azure AD Portal](https://portal.azure.com/#view/Microsoft_AAD_UsersAndTenants/UserManagementMenuBlade/~/AllUsers)
2. Search for "Test User Alpha"
3. Check profile:
   - **Other emails** should show the array values
   - **Fax number** should show the string
   - **Manager** should show Ingrid Johansen

### Step 4: Verify via Graph API

Use Graph Explorer or curl:

```bash
# Get Option A properties for Test User Alpha
GET https://graph.microsoft.com/beta/users/test.alpha@yourdomain.onmicrosoft.com
?$select=otherMails,faxNumber,preferredDataLocation
```

**Expected response**:
```json
{
  "otherMails": ["test.alpha.personal@gmail.com"],
  "faxNumber": "+47 22 11 11 12",
  "preferredDataLocation": "EUR"
}
```

### Step 5: Check the Log File

```bash
# Find the latest log file
ls -lt logs/ | head -n 2

# View the log
cat logs/provision-2026-01-22T*.log
```

**Look for**:
- No ERROR entries
- No WARN entries about property failures
- Success messages for all 3 users

## What Should Work

### ✅ Expected to Work
- otherMails array
- faxNumber string
- Manager relationships alongside new properties
- Custom columns are ignored by Option A (Option B can ingest labeled fields)

### ⚠️ May Not Work
- **preferredDataLocation**: Requires multi-geo license (E5 or add-on)
  - If you don't have this, it may silently ignore or return an error
  - Not critical for the test

## Test Results Template

Document your findings:

```markdown
## Test Results - [Your Date]

### Arrays
- otherMails: ✅ Array with X items / ❌ Stored as string / ❌ Failed

### Strings
- faxNumber: ✅ PASS

### Preferences
- preferredDataLocation: ✅ PASS / ⚠️ Not supported / ❌ Error

### Issues Found
- [List any issues]

### Notes
- [Any observations]
```

## Next: Test UPDATE

After verifying CREATE works:

1. **Edit the CSV** - Change some otherMails, update faxNumber
2. **Run provision again**:
   ```bash
   npm run provision -- --csv config/agents-test-maxprops.csv
   ```
3. **Verify updates** - Check that changes are reflected in Azure AD
4. **Check UPDATE behavior** - Are arrays replaced or appended?

## Next: Test EMPTY/REMOVE

Test clearing properties:

1. **Edit CSV** - Remove some array items, clear some properties
2. **Run provision again**
3. **Verify** - Are empty arrays clearing the properties?

## Cleanup (Optional)

If you want to remove test users:

```bash
# Delete test users from CSV or run with different CSV
npm run provision -- --csv config/agents-template.csv --force
```

Or manually delete in Azure AD Portal.

## Array Format Reference

Arrays in CSV must be JSON-encoded:

```csv
✅ Correct: "['test.alpha.personal@gmail.com','test.alpha@outlook.com']"
✅ Also correct: "[""test.alpha.personal@gmail.com"",""test.alpha@outlook.com""]"
❌ Wrong: "test.alpha.personal@gmail.com,test.alpha@outlook.com"
❌ Wrong: test.alpha.personal@gmail.com,test.alpha@outlook.com
```

The system now supports:
1. **JSON arrays** (preferred): `['value1','value2']` or `["value1","value2"]`
2. **Comma-separated** (fallback): `value1,value2,value3`

## Questions to Answer

1. **Array limits**: Do 5+ otherMails work? 10+?
2. **Update behavior**: When updating arrays, are they replaced or appended?
3. **Empty arrays**: Does `[]` clear an existing array property?
4. **Visibility**: Are these properties visible in user profile details?

## Success Criteria

- ✅ All 3 users created without errors
- ✅ Arrays stored as arrays (verified via Graph API)
- ✅ No properties stored as strings that should be arrays
- ✅ Manager relationships work

Once this passes, we have confirmed Option A works for simple properties!

## Then: Plan Option B

After Option A testing is complete, we can design Option B (profile resources) for:
- Skills with proficiency levels
- Certifications with issuers and dates
- Work history (positions)
- Awards with organizations
- Languages with proficiency

These benefit from the richer metadata available in profile resources.
