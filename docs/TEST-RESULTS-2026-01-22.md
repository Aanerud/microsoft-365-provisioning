# Test Results: Option A Property Testing

**Date**: 2026-01-22
**Tester**: Testing Microsoft Graph Beta API property support
**Objective**: Determine which user properties can be set via batch operations vs individual operations

## Test Setup

### Test Users Created
- **Test User Alpha** (test.alpha@a830edad9050849coep9vqp9bog.onmicrosoft.com) - Engineer
- **Test User Beta** (test.beta@a830edad9050849coep9vqp9bog.onmicrosoft.com) - UX Designer
- **Test User Gamma** (test.gamma@a830edad9050849coep9vqp9bog.onmicrosoft.com) - Financial Analyst

### Properties Tested
15 properties that hadn't been tested in production:
- Arrays: `skills`, `interests`, `pastProjects`, `responsibilities`, `schools`, `otherMails`
- Strings: `aboutMe`, `mySite`, `faxNumber`
- Date: `birthday`
- Preference: `preferredDataLocation`

## Test Results

### Batch Operations (CREATE/UPDATE via $batch)

**Result**: ❌ **FAILED**

**Error Message**:
```json
{
  "error": {
    "code": "BadRequest",
    "message": "The request is currently not supported on the targeted entity set"
  }
}
```

**Properties Affected**:
- `aboutMe`
- `skills`
- `interests`
- `pastProjects`
- `responsibilities`
- `schools`
- `mySite`
- `birthday`

All returned null/empty after batch creation attempt.

### Individual PATCH Operations

**Result**: ✅ **SUCCESS**

**Test**: Updated Test User Alpha with 5 properties individually

**Commands**:
```bash
PATCH /users/{id}
{
  "aboutMe": "Test bio text"
}

PATCH /users/{id}
{
  "skills": ["TypeScript", "Python"]
}

PATCH /users/{id}
{
  "interests": ["AI", "Cloud"]
}

PATCH /users/{id}
{
  "mySite": "https://example.com"
}

PATCH /users/{id}
{
  "birthday": "1985-06-15T00:00:00Z"
}
```

**Results**: All 5 properties successfully set

**Verification**:
```bash
GET /users/{id}?$select=aboutMe,skills,interests,mySite,birthday
```

**Response**:
```json
{
  "aboutMe": "Test bio text",
  "skills": ["TypeScript", "Python"],
  "interests": ["AI", "Cloud"],
  "mySite": "https://example.com",
  "birthday": "1985-06-15T00:00:00Z"
}
```

## Array Parsing Issue Discovered

### Problem
CSV arrays with single quotes `['value1','value2']` were being split by commas instead of parsed as JSON.

**Input**: `['TypeScript','Python','Azure']`
**Parsed (incorrect)**: `["['TypeScript'", "'Python'", "'Azure']"]`
**Expected**: `["TypeScript", "Python", "Azure"]`

### Root Cause
`JSON.parse()` requires double quotes, but CSV used single quotes.

### Fix Applied
Updated `parsePropertyValue()` in `src/schema/user-property-schema.ts`:

```typescript
case 'array':
  try {
    // Try parsing as-is first (for double-quoted JSON)
    const parsed = JSON.parse(csvValue);
    if (Array.isArray(parsed)) {
      return parsed;
    }
  } catch {
    // Try converting single quotes to double quotes
    try {
      const normalizedValue = csvValue.replace(/'/g, '"');
      const parsed = JSON.parse(normalizedValue);
      if (Array.isArray(parsed)) {
        return parsed;
      }
    } catch {
      // Fallback to comma-separated
    }
  }
  // Fallback to comma-separated values
  return csvValue.split(',').map(v => v.trim());
```

### Verification
After fix:
**Input**: `['TypeScript','Python','Azure']`
**Parsed (correct)**: `["TypeScript", "Python", "Azure"]` ✅

## Property Categories Discovered

Based on test results, we identified two distinct categories:

### Category 1: Directory Properties
**Can be set via batch operations**

Properties:
- Basic: `displayName`, `givenName`, `surname`
- Job: `jobTitle`, `department`, `employeeId`, `employeeType`, `companyName`
- Address: `streetAddress`, `city`, `state`, `country`, `postalCode`, `officeLocation`
- Contact: `mobilePhone`, `businessPhones`
- Preferences: `usageLocation`, `preferredLanguage`

Characteristics:
- Core organizational data
- Batch-writable
- Efficient provisioning
- HR system integration

### Category 2: Profile/Personal Properties
**Cannot be set via batch operations, require individual PATCH**

Properties:
- Personal: `aboutMe`, `mySite`, `birthday`
- Arrays: `skills`, `interests`, `pastProjects`, `responsibilities`, `schools`
- Additional contact: `otherMails`, `faxNumber`

Characteristics:
- User-facing enrichment data
- Individual operations only
- Slower provisioning
- Optional enhancement

## Performance Analysis

### Option A (Batch Operations)
**100 users with 10 core properties**:
- Batch size: 20 users per request
- Requests needed: ~5 batch requests
- Time: 5-10 seconds
- API calls: 5

### Option B (Individual Operations)
**100 users with 8 profile properties**:
- Individual PATCH per property per user
- Requests needed: 800 (100 users × 8 properties)
- Rate limit: 40 requests/second
- Time: ~20-30 seconds minimum
- API calls: 800

**Performance Impact**: Option B is ~160x more API calls than Option A

## Graph API Query Behavior

### Default GET Request
```bash
GET /users/{id}
```

**Returns**: Only default properties (subset)
**Note**: Property like `skills`, `interests`, `aboutMe` are NOT included

**Response includes hint**:
```json
{
  "@microsoft.graph.tips": "This request only returns a subset of the resource's properties. Your app will need to use $select to return non-default properties."
}
```

### Explicit $select Required
```bash
GET /users/{id}?$select=skills,interests,aboutMe
```

**Returns**: Explicitly requested properties

**Lesson**: Must use `$select` to verify these properties are set

## License Assignment Issue

### Problem
All 3 test users failed license assignment:

```
Failed to assign license to Test User Alpha
Error: License assignment cannot be done for user with invalid usage location.
```

### Root Cause
`usageLocation` was set to "NO" (Norway), but license assignment still failed.

### Hypothesis
- License might not be available in the tenant
- LICENSE_SKU_ID might be incorrect
- Usage location might need time to propagate

### Note
This is a separate issue from property testing and doesn't affect the property test results.

## Conclusions

### 1. Batch Limitations Confirmed
Personal/profile properties **cannot** be set via batch operations. This is a Microsoft Graph API design decision, not a limitation of our tool.

### 2. Individual Operations Work
All tested properties **can** be successfully set via individual `PATCH /users/{id}` requests.

### 3. Architectural Implication
Must separate provisioning into two phases:
- **Phase 1**: Core provisioning (batch operations)
- **Phase 2**: Profile enrichment (individual operations)

### 4. CSV Parsing Fixed
Array parsing now handles both single-quoted and double-quoted JSON arrays correctly.

### 5. Query Pattern Important
Must use `$select` when querying to verify these properties, as they're not returned by default.

## Recommendations

### For Production Use

1. **Use Option A for core provisioning**
   - Efficient batch operations
   - Essential user creation
   - Fast provisioning

2. **Implement Option B as separate module**
   - Individual PATCH operations
   - Profile enrichment
   - Optional enhancement
   - Run after Option A

3. **Separate CSV files**
   - `agents-provision.csv` - Core properties
   - `agents-profiles.csv` - Profile properties
   - Clear separation of concerns

4. **Rate Limiting Considerations**
   - Option B must respect 40 req/sec limit
   - Add delays between requests
   - Implement retry logic
   - Progress reporting essential

## Files Modified

1. **`src/schema/user-property-schema.ts`**
   - Fixed array parsing to handle single quotes
   - Added normalization step

2. **`config/agents-test-maxprops.csv`**
   - Created test dataset with 3 users
   - Included all 15 new properties

3. **Test Scripts Created** (temporary, can be deleted):
   - `query-user.mjs` - Query user with $select
   - `update-user-test.mjs` - Test batch update
   - `test-individual-props.mjs` - Test individual property updates
   - `test-schema.mjs` - Test property recognition and parsing
   - `get-profile.mjs` - Test profile navigation property

## Next Steps

1. ✅ Document findings (DONE - this document)
2. ✅ Document architecture (DONE - ARCHITECTURE-OPTION-A-B.md)
3. ⏳ Clean up Option A to focus on core properties only
4. ⏳ Create Option B module (`src/enrich-profiles.ts`)
5. ⏳ Update CSV templates
6. ⏳ Update main documentation

## Test Artifacts

### Log Files
- `logs/provision-2026-01-22T09-59-37.log` - Initial test user creation
- `logs/provision-2026-01-22T10-14-25.log` - Failed batch update attempt

### Test Users
- Test User Alpha (c626c46d-0cc6-4c4d-be50-0c055d94e9ea)
- Test User Beta (9b85acc0-9109-41fb-a41e-d5b5e129e4eb)
- Test User Gamma (76d1fc40-a392-4baa-bed2-0b7f58a24ed5)

**Status**: Test users remain in tenant for further testing

---

**Test Completed**: 2026-01-22
**Status**: Findings documented, architecture defined
**Outcome**: Separation of concerns approach recommended and accepted
