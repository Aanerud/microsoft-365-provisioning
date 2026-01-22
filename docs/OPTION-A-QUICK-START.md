# Option A Test - Quick Start Guide

## What We're Testing

We're pushing the standard `/users` endpoint to its limits by testing **15 new properties** that haven't been used yet:

### Properties Being Tested
- **Arrays**: `skills`, `interests`, `pastProjects`, `responsibilities`, `schools`, `otherMails`
- **Strings**: `aboutMe`, `mySite`, `faxNumber`
- **Date**: `birthday`
- **Preferences**: `preferredDataLocation`

## Test Dataset

**File**: `config/agents-test-maxprops.csv`

**Users**: 3 test users with comprehensive profiles:
- **Test User Alpha** - Engineer with tech skills, open source interests
- **Test User Beta** - UX Designer with design skills, creative interests
- **Test User Gamma** - Financial Analyst with data skills, economics interests

## Run the Test

### Step 1: Dry-Run (Preview Changes)
```bash
npm run provision -- --use-beta --csv config/agents-test-maxprops.csv --dry-run
```

**What to look for**:
- Should show "CREATE: 3 users"
- Check that it detects all the new properties
- No errors should appear

### Step 2: Create the Test Users
```bash
npm run provision -- --use-beta --csv config/agents-test-maxprops.csv
```

**What should happen**:
```
🔍 Fetching current Azure AD state...
📊 Calculating changes...

Summary:
- Total in CSV: 3 users
- To CREATE: 3 users
- Custom properties detected: 3 (VTeam, BenefitPlan, ...)

📦 Applying changes...
  Creating 3 users...
  ✓ Created 3 users
  Creating 3 open extensions...
  ✓ Created 3 open extensions
  Assigning 2 manager relationships...
  ✓ Assigned 2 managers

✅ Provisioning complete!
```

### Step 3: Verify in Azure AD Portal

1. Go to [Azure AD Portal](https://portal.azure.com/#view/Microsoft_AAD_UsersAndTenants/UserManagementMenuBlade/~/AllUsers)
2. Search for "Test User Alpha"
3. Check profile:
   - **About me** section should show the bio
   - **Skills** section should show array of skills
   - **Interests** should be populated
   - **Birthday** should show the date
   - **Personal website** should show the URL
   - **Manager** should show Ingrid Johansen

### Step 4: Verify via Graph API

Use Graph Explorer or curl:

```bash
# Get all personal properties for Test User Alpha
GET https://graph.microsoft.com/beta/users/test.alpha@a830edad9050849coep9vqp9bog.onmicrosoft.com
?$select=aboutMe,skills,interests,pastProjects,responsibilities,schools,mySite,birthday,otherMails,faxNumber,preferredDataLocation
```

**Expected response**:
```json
{
  "aboutMe": "Experienced engineer specializing in...",
  "skills": [
    "TypeScript",
    "Python",
    "Azure",
    "Kubernetes",
    "DevOps",
    "CI/CD",
    "Docker",
    "Terraform"
  ],
  "interests": [
    "Open Source",
    "Cloud Computing",
    "AI/ML",
    "Mountain Hiking",
    "Photography",
    "Tech Blogging"
  ],
  "pastProjects": [
    "Cloud Migration Initiative",
    "Kubernetes Platform Build",
    "CI/CD Pipeline Automation",
    "Microservices Architecture"
  ],
  "responsibilities": [
    "Lead Platform Team",
    "Mentor Junior Engineers",
    "Maintain CI/CD Infrastructure",
    "Technical Documentation"
  ],
  "schools": [
    "Norwegian University of Science and Technology",
    "MIT OpenCourseWare"
  ],
  "mySite": "https://testalpha.dev",
  "birthday": "1985-06-15T00:00:00Z",
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
- All 8 array properties stored as actual arrays (not strings)
- aboutMe string (500+ characters)
- mySite URL validation
- birthday date parsing
- otherMails array
- faxNumber string
- Manager relationships alongside new properties
- Custom properties still in open extensions

### ⚠️ May Not Work
- **preferredDataLocation**: Requires multi-geo license (E5 or add-on)
  - If you don't have this, it may silently ignore or return an error
  - Not critical for the test

## Test Results Template

Document your findings:

```markdown
## Test Results - [Your Date]

### Arrays
- skills: ✅ Array with X items / ❌ Stored as string / ❌ Failed
- interests:
- pastProjects:
- responsibilities:
- schools:
- otherMails:

### Strings
- aboutMe: ✅ PASS (X characters)
- mySite: ✅ PASS
- faxNumber: ✅ PASS

### Date
- birthday: ✅ PASS (ISO format)

### Preferences
- preferredDataLocation: ✅ PASS / ⚠️ Not supported / ❌ Error

### Issues Found
- [List any issues]

### Notes
- [Any observations]
```

## Next: Test UPDATE

After verifying CREATE works:

1. **Edit the CSV** - Change some skills, update aboutMe, add interests
2. **Run provision again**:
   ```bash
   npm run provision -- --use-beta --csv config/agents-test-maxprops.csv
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
npm run provision -- --use-beta --csv config/agents-template.csv --force
```

Or manually delete in Azure AD Portal.

## Array Format Reference

Arrays in CSV must be JSON-encoded:

```csv
✅ Correct: "['TypeScript','Python','Azure']"
✅ Also correct: "[""TypeScript"",""Python"",""Azure""]"
❌ Wrong: "TypeScript,Python,Azure"
❌ Wrong: TypeScript,Python,Azure
```

The system now supports:
1. **JSON arrays** (preferred): `['value1','value2']` or `["value1","value2"]`
2. **Comma-separated** (fallback): `value1,value2,value3`

## Questions to Answer

1. **Array limits**: Do 8 skills work? 50 skills? 100 skills?
2. **String limits**: Does 500-char aboutMe work? 1000? 2000?
3. **Special characters**: Do emoji work in arrays? Unicode characters?
4. **Update behavior**: When updating arrays, are they replaced or appended?
5. **Empty arrays**: Does `[]` clear an existing array property?
6. **Visibility**: Are these properties searchable in Microsoft Search?
7. **Profile cards**: Do these show up in Microsoft 365 profile cards?

## Success Criteria

- ✅ All 3 users created without errors
- ✅ Arrays stored as arrays (verified via Graph API)
- ✅ No properties stored as strings that should be arrays
- ✅ Date parsed correctly
- ✅ Manager relationships work
- ✅ Custom properties still work

Once this passes, we have confirmed Option A works for simple properties!

## Then: Plan Option B

After Option A testing is complete, we can design Option B (profile resources) for:
- Skills with proficiency levels
- Certifications with issuers and dates
- Work history (positions)
- Awards with organizations
- Languages with proficiency

These benefit from the richer metadata available in profile resources.
