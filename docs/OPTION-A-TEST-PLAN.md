# Option A: Standard User Properties - Comprehensive Test Plan

## Test Objective

Push the `/users` endpoint to its limits by testing **all writable simple properties** available in Microsoft Graph Beta API.

## Test Dataset

**File**: `config/agents-test-maxprops.csv`

**Users**: 3 test users with maximum property coverage

## Properties Being Tested

### ✅ Already Tested (Current 20-user dataset)
- Basic: `displayName`, `givenName`, `surname`
- Contact: `mail`, `mobilePhone`, `businessPhones`
- Address: `city`, `state`, `country`, `postalCode`, `streetAddress`, `officeLocation`
- Job: `jobTitle`, `department`, `companyName`, `employeeId`, `employeeType`, `employeeHireDate`
- Preferences: `usageLocation`, `preferredLanguage`
- Custom: `VTeam`, `BenefitPlan`, `CostCenter`, `BuildingAccess`, `ProjectCode`
- Navigation: `ManagerEmail` (manager relationship)

### 🧪 NEW Properties to Test

#### Personal Arrays
| Property | Type | Example Value | Expected Behavior |
|----------|------|---------------|-------------------|
| `skills` | array | `['TypeScript','Python','Azure']` | Array of strings |
| `interests` | array | `['Open Source','AI/ML','Hiking']` | Array of strings |
| `pastProjects` | array | `['Cloud Migration','Platform Build']` | Array of strings |
| `responsibilities` | array | `['Lead Platform Team','Mentor Engineers']` | Array of strings |
| `schools` | array | `['NTNU','MIT OpenCourseWare']` | Array of strings |

#### Personal Strings
| Property | Type | Example Value | Expected Behavior |
|----------|------|---------------|-------------------|
| `aboutMe` | string | "Experienced engineer specializing in..." | Freeform text bio |
| `mySite` | string | "https://testalpha.dev" | Personal website URL |

#### Personal Date
| Property | Type | Example Value | Expected Behavior |
|----------|------|---------------|-------------------|
| `birthday` | date | "1985-06-15" | ISO date format |

#### Contact Arrays
| Property | Type | Example Value | Expected Behavior |
|----------|------|---------------|-------------------|
| `otherMails` | array | `['test.alpha.personal@gmail.com']` | Additional email addresses |
| `faxNumber` | string | "+47 22 11 11 12" | Fax number (legacy) |

#### Preferences
| Property | Type | Example Value | Expected Behavior |
|----------|------|---------------|-------------------|
| `preferredDataLocation` | string | "EUR" | Multi-geo location (EUR, NAM, APC, etc.) |

## Test Scenarios

### Scenario 1: CREATE with Maximum Properties
**Goal**: Create 3 new users with ALL properties populated

**Steps**:
```bash
# Ensure users don't exist
npm run provision -- --use-beta --csv config/agents-test-maxprops.csv --dry-run

# Create users
npm run provision -- --use-beta --csv config/agents-test-maxprops.csv
```

**Expected Results**:
- ✅ All 3 users created
- ✅ Simple strings populated (aboutMe, mySite, faxNumber)
- ✅ Date fields parsed correctly (birthday)
- ✅ Arrays properly stored (skills, interests, pastProjects, responsibilities, schools, otherMails)
- ✅ preferredDataLocation set correctly
- ✅ Manager relationships established
- ✅ Custom properties in open extensions

**Validation Methods**:
1. **Azure AD Portal**: Check user profiles manually
2. **Graph API Query**:
   ```bash
   GET https://graph.microsoft.com/beta/users/test.alpha@domain.com
   ?$select=aboutMe,skills,interests,pastProjects,responsibilities,schools,mySite,birthday,otherMails,faxNumber,preferredDataLocation
   ```
3. **Check Log File**: `logs/provision-*.log`

### Scenario 2: UPDATE Properties
**Goal**: Modify array and string properties on existing users

**Steps**:
1. Edit CSV - change some skills, add interests, update aboutMe
2. Run provision again:
   ```bash
   npm run provision -- --use-beta --csv config/agents-test-maxprops.csv
   ```

**Expected Results**:
- ✅ Changed properties updated in Azure AD
- ✅ Arrays replaced (not appended)
- ✅ Unchanged properties remain intact

### Scenario 3: Remove/Empty Properties
**Goal**: Test emptying array and string properties

**Steps**:
1. Edit CSV - remove some values (empty arrays, blank strings)
2. Run provision:
   ```bash
   npm run provision -- --use-beta --csv config/agents-test-maxprops.csv
   ```

**Expected Results**:
- ✅ Empty arrays should clear the property
- ✅ Empty strings should clear the property
- ⚠️ Possible caveat: Some properties might not accept null/empty

## Array Encoding in CSV

Arrays must be JSON-encoded strings in CSV:

### ✅ Correct Format
```csv
skills,"['TypeScript','Python','Azure']"
interests,"['Open Source','AI/ML']"
businessPhones,"['+47 22 11 11 11']"
otherMails,"['test@gmail.com','test@outlook.com']"
```

### ❌ Incorrect Formats
```csv
skills,TypeScript,Python,Azure          # Wrong - would be 3 separate columns
skills,"TypeScript,Python,Azure"        # Wrong - comma-separated string
skills,"['TypeScript''Python''Azure']"  # Wrong - missing commas
```

## Known Limitations & Edge Cases

### Beta-Only Properties
Some properties require beta endpoint:
- `employeeHireDate` (beta only)
- `preferredDataLocation` (beta only - requires multi-geo license)

### Date Format
- **Input**: ISO 8601 format: `YYYY-MM-DD` or `YYYY-MM-DDTHH:mm:ssZ`
- **Output**: Azure AD stores as ISO 8601 with timezone

### Array Limits
Unknown if Microsoft Graph enforces limits on:
- Number of array items (e.g., max 50 skills?)
- String length per array item
- Total property size

**To Test**: Add 100+ items to skills array and see what happens

### preferredDataLocation
- **Requires**: Multi-geo capabilities in Microsoft 365 (E5 or add-on)
- **Values**: EUR (Europe), NAM (North America), APC (Asia Pacific), etc.
- **Error if not enabled**: May return 400 error or silently ignore

### Personal Properties Visibility
Properties like `skills`, `interests`, `aboutMe` are typically:
- Visible in Microsoft 365 profile cards
- Searchable in Microsoft Search
- Visible to Copilot for context
- Privacy settings may affect visibility

## Success Criteria

### Must Pass
- ✅ All 3 test users created without errors
- ✅ String properties (aboutMe, mySite) stored correctly
- ✅ Date properties (birthday) parsed and stored
- ✅ Arrays (skills, interests, etc.) stored as arrays (not strings)
- ✅ Manager relationships work alongside new properties
- ✅ Custom properties still work in open extensions

### Should Pass
- ✅ Arrays with 10+ items work
- ✅ aboutMe with 500+ characters works
- ✅ Update operations change only modified properties
- ✅ Empty arrays clear existing values

### May Fail (Document for Option B)
- ⚠️ preferredDataLocation (requires multi-geo license)
- ⚠️ Arrays with 100+ items (may hit limits)
- ⚠️ Special characters in array strings
- ⚠️ Very long strings (1000+ characters)

## Test Execution Checklist

- [ ] Build latest code: `npm run build`
- [ ] Test CSV parsing with dry-run
- [ ] CREATE: Provision 3 test users
- [ ] Verify in Azure AD Portal
- [ ] Verify via Graph API query
- [ ] Check log file for errors/warnings
- [ ] UPDATE: Modify CSV and re-provision
- [ ] Verify updates applied correctly
- [ ] EMPTY: Clear some properties and re-provision
- [ ] Verify properties cleared
- [ ] Document any failures or limitations
- [ ] Clean up test users (if needed)

## Results Documentation Template

```markdown
## Test Results - [Date]

### Test 1: CREATE with Maximum Properties
- Status: ✅ PASS / ❌ FAIL
- Users Created: X/3
- Properties Set: X/15 new properties
- Errors: [list any errors]
- Notes: [observations]

### Test 2: Array Properties
- skills: ✅ Stored as array / ❌ Stored as string / ❌ Failed
- interests: [result]
- pastProjects: [result]
- responsibilities: [result]
- schools: [result]
- otherMails: [result]
- Notes: [array item limits, special characters, etc.]

### Test 3: String Properties
- aboutMe: ✅ PASS (X characters) / ❌ FAIL
- mySite: ✅ PASS / ❌ FAIL
- faxNumber: ✅ PASS / ❌ FAIL
- Notes: [character limits, special characters]

### Test 4: Date Properties
- birthday: ✅ PASS / ❌ FAIL
- Format stored: [ISO 8601 format observed]
- Notes: [timezone handling, etc.]

### Test 5: Multi-Geo
- preferredDataLocation: ✅ PASS / ⚠️ Not supported / ❌ FAIL
- Notes: [license requirements]

### Limitations Discovered
1. [Limitation 1]
2. [Limitation 2]
...

### Recommended for Option B (Profile Resources)
- [Properties that would benefit from profile resources]
- [Reasons why]
```

## Next Steps After Testing

Based on test results:

1. **Document working properties** → Update main CSV template
2. **Document limitations** → Add to USAGE.md
3. **Plan Option B** → Identify properties better suited for profile resources
4. **Update schema** → Mark any properties as beta-only or unsupported

## Option B Candidates

Properties that might work better with profile resources:
- **Skills with proficiency levels** → `profile/skills` (includes proficiency, collaboration tags)
- **Work history** → `profile/positions` (rich work position objects)
- **Certifications** → `profile/certifications` (includes issuer, date, etc.)
- **Awards** → `profile/awards` (includes issuing organization, date)
- **Languages** → `profile/languages` (includes proficiency level)

## References

- [User Resource (Beta)](https://learn.microsoft.com/en-us/graph/api/resources/user?view=graph-rest-beta)
- [Update User (Beta)](https://learn.microsoft.com/en-us/graph/api/user-update?view=graph-rest-beta)
- [User Properties](https://learn.microsoft.com/en-us/graph/api/resources/user?view=graph-rest-beta#properties)
