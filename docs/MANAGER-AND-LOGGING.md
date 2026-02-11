# Manager Relationships & Logging System

## Overview

This document covers two major enhancements to the provisioning tool:

1. **Manager Relationship Management** - Organizational hierarchy
2. **Comprehensive Logging System** - Error tracking and audit trails

## 🏢 Manager Relationships

### What is a Manager Relationship?

Microsoft Graph API allows you to establish organizational hierarchy by assigning managers to users. This creates a reporting structure visible in:
- Azure AD organizational charts
- Microsoft Teams
- Outlook
- Microsoft 365 apps

### How It Works

**API Endpoint**: `PUT /users/{userId}/manager/$ref`

**Request Body**:
```json
{
  "@odata.id": "https://graph.microsoft.com/beta/users/{managerId}"
}
```

### CSV Configuration

Add a `ManagerEmail` column to your CSV:

```csv
name,email,department,ManagerEmail
Ingrid Johansen,ingrid.johansen@domain.com,Executive,
Lars Hansen,lars.hansen@domain.com,Engineering,ingrid.johansen@domain.com
Kari Andersen,kari.andersen@domain.com,Engineering,lars.hansen@domain.com
```

**Rules**:
- Leave `ManagerEmail` empty for top-level executives (no manager)
- Use the manager's email address (must exist in Azure AD or CSV)
- Manager must be created before assignment (tool handles this automatically)

### Current CSV Structure (20 Norwegian Users)

The CSV already includes proper manager relationships:

**Executive Level**:
- **Ingrid Johansen** (CEO) - No manager
- **Lars Hansen** (CTO) - Reports to Ingrid

**Department Managers**:
- **Kari Andersen** (Engineering Manager) - Reports to Lars
- **Kristian Svendsen** (Product Manager) - Reports to Ingrid
- **Anne Bakke** (Marketing Manager) - Reports to Ingrid
- **Silje Haugen** (Sales Director) - Reports to Ingrid
- **Tone Eriksen** (HR Manager) - Reports to Ingrid
- **Bjørn Strand** (Finance Manager) - Reports to Ingrid

**Team Members** - Report to their department managers

### Provisioning Flow

```
1. CREATE all users in Azure AD
   ↓
2. Assign licenses (if configured)
   ↓
3. Create custom property extensions
   ↓
4. Resolve manager relationships:
   - Check if manager exists in current batch
   - Check if manager already exists in Azure AD
   - Build manager assignment list
   ↓
5. Batch assign managers (20 per batch)
   ↓
6. Log successes and failures
```

### Manager Assignment Methods

**Single Assignment**:
```typescript
await graphClient.assignManager(userId, managerId);
```

**Batch Assignment** (Recommended):
```typescript
await graphClient.assignManagersBatch([
  { userId: 'user-id-1', managerId: 'manager-id-1' },
  { userId: 'user-id-2', managerId: 'manager-id-2' },
]);
```

**Get Manager**:
```typescript
const manager = await graphClient.getManager(userId);
```

**Remove Manager**:
```typescript
await graphClient.removeManager(userId);
```

### Querying Managers in Graph API

**Get user with manager**:
```
GET https://graph.microsoft.com/beta/users/{userId}?$expand=manager
```

**Get user's direct reports**:
```
GET https://graph.microsoft.com/beta/users/{userId}/directReports
```

**Get manager's details**:
```
GET https://graph.microsoft.com/beta/users/{userId}/manager
```

### Error Handling

**Common Scenarios**:

1. **Manager not found**:
   ```
   ⚠ Manager lars.hansen@domain.com not found for Kari Andersen
   ```
   - Manager email doesn't exist in CSV or Azure AD
   - Check spelling
   - Ensure manager is created first

2. **Circular reference**:
   - Cannot assign user as their own manager
   - Cannot create circular reporting chains

3. **Permission error**:
   - Requires `User.ReadWrite.All` or `Directory.ReadWrite.All`
   - Check API permissions in Azure AD app

### Organizational Chart Example

After provisioning with manager relationships:

```
Ingrid Johansen (CEO)
├── Lars Hansen (CTO)
│   └── Kari Andersen (Engineering Manager)
│       ├── Ola Nilsen (Senior Developer)
│       ├── Sofie Berg (Frontend Developer)
│       ├── Erik Olsen (Backend Developer)
│       ├── Maria Pettersen (DevOps Engineer)
│       ├── Jonas Kristiansen (Cloud Architect)
│       └── Emma Larsen (QA Engineer)
├── Kristian Svendsen (Product Manager)
│   └── Lise Moen (UX Designer)
│       └── Henrik Dahl (Graphic Designer)
├── Anne Bakke (Marketing Manager)
│   └── Per Solberg (Content Specialist)
├── Silje Haugen (Sales Director)
│   └── Geir Lund (Account Manager)
├── Tone Eriksen (HR Manager)
│   └── Marte Holm (HR Specialist)
└── Bjørn Strand (Finance Manager)
    └── Hilde Vik (Accountant)
```

## 📝 Logging System

### What Gets Logged?

The logging system tracks:

1. **Errors** - Failed operations (license assignment, manager assignment, API errors)
2. **Warnings** - Non-critical issues (missing managers, skipped operations)
3. **Info** - General operational info
4. **Success** - Completed operations
5. **Debug** - Detailed technical info

### Log File Location

```
logs/provision-YYYY-MM-DDTHH-mm-ss.log
```

Example: `logs/provision-2026-01-22T15-30-45.log`

### Log File Format

```
[2026-01-22T15:30:45.123Z] [INFO   ] Logging to: logs/provision-2026-01-22T15-30-45.log
[2026-01-22T15:30:46.234Z] [INFO   ] Loading agents from config/agents-template.csv...
[2026-01-22T15:30:47.345Z] [SUCCESS] Created user: Ingrid Johansen
[2026-01-22T15:30:48.456Z] [WARN   ] Failed to assign license to Lars Hansen
{
  "userId": "user-id-123",
  "error": "Insufficient licenses available"
}
[2026-01-22T15:30:49.567Z] [ERROR  ] Failed to assign manager for Kari Andersen
{
  "userId": "user-id-456",
  "managerId": "manager-id-789",
  "error": "Manager not found"
}

================================================================================
Provisioning Summary
================================================================================
Created:  20
Updated:  0
Deleted:  0
Errors:   2
Warnings: 5
Completed: 2026-01-22T15:31:00.000Z
================================================================================
```

### Logger API

```typescript
// Initialize logger
const logger = await initializeLogger();

// Log messages
logger.info('User created successfully');
logger.warn('License assignment failed', { userId: 'id-123' });
logger.error('API error occurred', { error: 'details' });
logger.success('Operation completed');
logger.debug('Debug information');

// Get log file path
const logPath = logger.getLogFilePath();

// Write summary
await logger.writeSummary({
  created: 20,
  updated: 0,
  deleted: 0,
  errors: 2,
  warnings: 5,
});

// Close logger
await logger.close();
```

### Console + File Output

All log messages go to **both**:
1. **Console** - Real-time feedback with emoji indicators
2. **Log file** - Permanent record with timestamps

**Console Output**:
```
✅ Created user: Ingrid Johansen
⚠️  Failed to assign license to Lars Hansen
❌ Failed to assign manager for Kari Andersen
```

**Log File Output**:
```
[2026-01-22T15:30:47.345Z] [SUCCESS] Created user: Ingrid Johansen
[2026-01-22T15:30:48.456Z] [WARN   ] Failed to assign license to Lars Hansen
[2026-01-22T15:30:49.567Z] [ERROR  ] Failed to assign manager for Kari Andersen
```

### License Assignment Errors

**Common Error Logged**:
```
[WARN   ] Failed to assign license to Lars Hansen
{
  "userId": "abc123-def456",
  "error": "Subscription is not available"
}
```

**Possible Causes**:
1. No licenses available in tenant
2. Incorrect LICENSE_SKU_ID in .env
3. UsageLocation not set
4. License already assigned
5. User not eligible for license

**Solution**:
1. Check LICENSE_SKU_ID in .env matches your tenant
2. Ensure usageLocation is set (NO for Norway)
3. Verify licenses available in Azure AD portal
4. Check log file for specific error details

### Viewing Logs

**After provisioning**:
```bash
# View latest log
ls -lt logs/ | head -n 2

# Read log file
cat logs/provision-2026-01-22T15-30-45.log

# Search for errors
grep ERROR logs/provision-2026-01-22T15-30-45.log

# Search for warnings
grep WARN logs/provision-2026-01-22T15-30-45.log

# View summary
tail -n 20 logs/provision-2026-01-22T15-30-45.log
```

### Log Analysis

**Check for license errors**:
```bash
grep "Failed to assign license" logs/*.log
```

**Check for manager assignment errors**:
```bash
grep "Failed to assign manager" logs/*.log
```

**Count errors**:
```bash
grep -c ERROR logs/provision-2026-01-22T15-30-45.log
```

**Count warnings**:
```bash
grep -c WARN logs/provision-2026-01-22T15-30-45.log
```

## Testing

### Test Manager Relationships

1. **Dry-run first**:
```bash
npm run provision -- --dry-run
```

2. **Provision users**:
```bash
npm run provision
```

3. **Verify in Azure AD**:
   - Go to Azure AD Portal
   - Select a user (e.g., Kari Andersen)
   - Check "Manager" field
   - Should show Lars Hansen

4. **Query via Graph API**:
```bash
# Using Graph Explorer (https://developer.microsoft.com/graph/graph-explorer)
GET https://graph.microsoft.com/beta/users/kari.andersen@domain.com?$expand=manager

# Response includes manager details
{
  "displayName": "Kari Andersen",
  "manager": {
    "displayName": "Lars Hansen",
    "userPrincipalName": "lars.hansen@domain.com"
  }
}
```

5. **Check organizational chart**:
   - Open Microsoft Teams
   - Go to user profile
   - View organizational chart

### Test Logging

1. **Run provision**:
```bash
npm run provision
```

2. **Check console for log file path**:
```
📄 Log file: logs/provision-2026-01-22T15-30-45.log
```

3. **Review log file**:
```bash
cat logs/provision-2026-01-22T15-30-45.log
```

4. **Look for errors/warnings**:
```bash
grep -E "(ERROR|WARN)" logs/provision-2026-01-22T15-30-45.log
```

## Summary

### ✅ What's Now Supported

**Manager Relationships**:
- ✅ Assign managers via CSV `ManagerEmail` column
- ✅ Automatic manager resolution
- ✅ Batch manager assignment (20 per batch)
- ✅ Error handling for missing managers
- ✅ Organizational hierarchy in Azure AD

**Logging System**:
- ✅ Automatic log file creation
- ✅ Timestamps on all entries
- ✅ Error and warning tracking
- ✅ License assignment failure logging
- ✅ Manager assignment failure logging
- ✅ Comprehensive summary report
- ✅ Both console and file output

### 📊 Current CSV Features

Your 20-user Norwegian dataset includes:

**Standard Properties (17)**:
- name, givenName, surname, email
- jobTitle, department, employeeType, companyName
- officeLocation, streetAddress, city, state, country, postalCode
- usageLocation (NO), preferredLanguage (nb-NO)
- mobilePhone, businessPhones, employeeId, employeeHireDate

**Custom Properties (6)**:
- VTeam, BenefitPlan, CostCenter, BuildingAccess, ProjectCode

**NEW: Navigation Property (1)**:
- ManagerEmail (establishes organizational hierarchy)

### 🔍 Investigating Issues

**License Assignment Failures**:
1. Check log file for specific error
2. Verify LICENSE_SKU_ID in .env
3. Check available licenses in Azure AD
4. Ensure usageLocation is set

**Manager Assignment Failures**:
1. Check log file for missing managers
2. Verify manager email exists
3. Ensure manager created before assignment
4. Check API permissions

## References

**Sources**:
- [Assign manager - Microsoft Graph beta](https://learn.microsoft.com/en-us/graph/api/user-post-manager?view=graph-rest-beta)
- [Update user - Microsoft Graph beta](https://learn.microsoft.com/en-us/graph/api/user-update?view=graph-rest-beta)
- [Working with users in Microsoft Graph](https://learn.microsoft.com/en-us/graph/api/resources/users?view=graph-rest-beta)
- [Set-MgUserManagerByRef PowerShell](https://learn.microsoft.com/en-us/powershell/module/microsoft.graph.users/set-mgusermanagerbyref?view=graph-powershell-1.0)

---

**Your provisioning tool now includes**:
- 🛡️ Account Protection (admin accounts safe)
- 🏢 Manager Relationships (organizational hierarchy)
- 📝 Comprehensive Logging (error tracking and audit trails)
- 📊 50+ Standard Properties + Unlimited Custom Properties
- 🔄 Full State Management (CREATE/UPDATE/DELETE)

Ready for production use! 🚀
