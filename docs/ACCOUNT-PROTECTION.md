# Account Protection System

## Overview

The Account Protection System is a **critical safety feature** that prevents the provisioning tool from accidentally modifying or deleting admin accounts and other protected users. This system operates automatically and requires no user intervention.

## 🛡️ Protection Layers

The system uses **three layers of protection**:

### 1. Email Pattern Matching (Fast)

Protected email patterns using wildcards:
- `admin@*` - Any email starting with "admin@"
- `administrator@*` - Any email starting with "administrator@"
- `root@*` - Any email starting with "root@"
- `systemadmin@*` - Any email starting with "systemadmin@"

**Example**: `admin@a830edad9050849coep9vqp9bog.onmicrosoft.com` is automatically protected.

### 2. Explicit Exclusion List

Exact email addresses that should always be protected:
```bash
PROTECTED_EMAILS=ceo@domain.com,finance@domain.com
```

### 3. Azure AD Role Detection (Automatic)

Users with privileged Azure AD roles are automatically protected:
- **Global Administrator**
- **Privileged Role Administrator**
- **Security Administrator**
- **User Administrator**
- **Directory Synchronization Accounts**

The system queries Azure AD during delta calculation to check user roles.

## Configuration

### Environment Variables

Add to your `.env` file:

```bash
# Protected email patterns (wildcards supported)
PROTECTED_EMAIL_PATTERNS=admin@*,administrator@*,root@*,systemadmin@*

# Explicit protected emails (exact matches)
PROTECTED_EMAILS=ceo@domain.com,finance@domain.com

# Check for Azure AD admin roles (recommended)
CHECK_ADMIN_ROLES=true
```

### Default Protection (No Configuration Required)

Even without any configuration, the system protects:
- All `admin@*` accounts
- All `administrator@*` accounts
- All `root@*` accounts
- All users with Global Administrator role

## How It Works

### During Delta Calculation

```
1. Tool calculates what needs to be created/updated/deleted
2. Protection service checks each UPDATE/DELETE action
3. Protected accounts are filtered out
4. Warning message displayed to user
5. Protected accounts moved to "NO_CHANGE" category
```

### Warning Display

When protected accounts are detected:

```
⚠️ ⚠️ ⚠️ ⚠️ ⚠️ ⚠️ ⚠️ ⚠️ ⚠️ ⚠️ ⚠️ ⚠️ ⚠️ ⚠️ ⚠️ ⚠️ ⚠️ ⚠️ ⚠️ ⚠️
🛡️  PROTECTED ACCOUNTS - DELETE BLOCKED
⚠️ ⚠️ ⚠️ ⚠️ ⚠️ ⚠️ ⚠️ ⚠️ ⚠️ ⚠️ ⚠️ ⚠️ ⚠️ ⚠️ ⚠️ ⚠️ ⚠️ ⚠️ ⚠️ ⚠️

The following 1 account(s) are protected and will NOT be DELETEd:

1. admin@a830edad9050849coep9vqp9bog.onmicrosoft.com
   Reason: Email matches protected pattern
   Role: Global Administrator

To modify protection settings, see .env configuration:
  - PROTECTED_EMAIL_PATTERNS
  - PROTECTED_EMAILS
  - CHECK_ADMIN_ROLES

────────────────────────────────────────────────────────────────────────────────
```

## Protected Operations

### DELETE Protection

Protected accounts are **never deleted**, even if:
- They don't exist in the CSV file
- `--force` flag is used
- `--skip-delete` is NOT used

**Result**: Account remains in Azure AD unchanged.

### UPDATE Protection

Protected accounts are **never updated**, even if:
- CSV has different values
- Properties have changed

**Result**: Account remains unchanged with current values.

### CREATE Protection

CREATE operations are **not protected** because:
- New accounts don't exist yet
- No risk of overwriting existing admin accounts

## Use Cases

### Use Case 1: Your Admin Account

**Scenario**: You're signed in as `admin@domain.com` and running the provisioning tool.

**Protection**:
- Pattern matching: ✅ Matches `admin@*`
- Role detection: ✅ Has Global Administrator role

**Result**: Your account is never touched, even if not in CSV.

### Use Case 2: Service Accounts

**Scenario**: You have service accounts for automation:
- `automation@domain.com`
- `sync@domain.com`

**Protection**:
```bash
# Add to .env
PROTECTED_EMAILS=automation@domain.com,sync@domain.com
```

**Result**: These accounts are never modified/deleted.

### Use Case 3: Executive Accounts

**Scenario**: CEO and CFO accounts should never be managed by this tool:
- `ceo@domain.com`
- `cfo@domain.com`

**Protection**:
```bash
# Add to .env
PROTECTED_EMAILS=ceo@domain.com,cfo@domain.com
```

**Result**: Executive accounts remain unchanged.

### Use Case 4: Global Administrators

**Scenario**: Multiple Global Administrators in your tenant.

**Protection**:
- Automatic role detection: ✅ Enabled by default
- All Global Administrators protected automatically

**Result**: No configuration needed, all admins protected.

## Testing Protection

### Dry-Run Test

```bash
# Start with admin account in Azure AD, not in CSV
npm run provision -- --dry-run

# Expected output:
# 🛡️  Applying account protection filters...
# ⚠️  PROTECTED ACCOUNTS - DELETE BLOCKED
# 1. admin@domain.com
#    Reason: Email matches protected pattern
```

### Live Test (Safe)

```bash
# Even with --force, protected accounts are safe
npm run provision -- --force

# Protected accounts will NOT be deleted
```

## Disabling Protection (Not Recommended)

### Disable Email Pattern Protection

```bash
# Set to empty string (not recommended)
PROTECTED_EMAIL_PATTERNS=
```

⚠️ **Warning**: This removes protection from admin@* accounts!

### Disable Role-Based Protection

```bash
# Disable role checking (not recommended)
CHECK_ADMIN_ROLES=false
```

⚠️ **Warning**: Global Administrators will no longer be automatically protected!

## Advanced Configuration

### Custom Protected Patterns

Add organization-specific patterns:

```bash
# Protect all accounts in specific subdomain
PROTECTED_EMAIL_PATTERNS=admin@*,*@security.domain.com,*@compliance.domain.com

# Protect service accounts with naming pattern
PROTECTED_EMAIL_PATTERNS=admin@*,svc-*@*,automation-*@*
```

### Multiple Explicit Exclusions

```bash
# Comma-separated list
PROTECTED_EMAILS=admin@domain.com,ceo@domain.com,cfo@domain.com,backup-admin@domain.com
```

### Pattern Syntax

- `*` - Matches any characters
- `?` - Matches single character
- Case-insensitive matching

**Examples**:
- `admin@*` matches `admin@domain.com`, `Admin@example.com`
- `*admin@*` matches `superadmin@domain.com`, `admin@domain.com`
- `test-??@*` matches `test-01@domain.com`, `test-99@domain.com`

## Troubleshooting

### Account Not Protected

**Problem**: Expected account is being deleted/updated

**Solutions**:
1. Check email pattern: Does it match `PROTECTED_EMAIL_PATTERNS`?
2. Check exclusion list: Is email in `PROTECTED_EMAILS`?
3. Check role detection: Is `CHECK_ADMIN_ROLES=true`?
4. Verify Azure AD role: Does user have protected role?

### Too Many Protected Accounts

**Problem**: Too many accounts are being protected

**Solution**: Be more specific with patterns:
```bash
# Too broad (protects everything)
PROTECTED_EMAIL_PATTERNS=*@domain.com

# Better (specific patterns)
PROTECTED_EMAIL_PATTERNS=admin@*,root@*
```

### Performance Issues

**Problem**: Role detection is slow

**Solution**: Disable role checking if not needed:
```bash
CHECK_ADMIN_ROLES=false
```

Note: Email pattern matching and exclusion list remain fast.

## Security Best Practices

### 1. Always Protect Admin Accounts

```bash
# Minimum recommended protection
PROTECTED_EMAIL_PATTERNS=admin@*,administrator@*,root@*
CHECK_ADMIN_ROLES=true
```

### 2. Protect Service Accounts

```bash
# Add automation/service accounts
PROTECTED_EMAILS=automation@domain.com,sync@domain.com,backup@domain.com
```

### 3. Test in Development First

```bash
# Always test with dry-run first
npm run provision -- --dry-run
```

### 4. Monitor Protection Warnings

- Review protection warnings in output
- Verify expected accounts are protected
- Adjust patterns if needed

### 5. Document Custom Patterns

```bash
# Comment your custom patterns
# Pattern for security team accounts
PROTECTED_EMAIL_PATTERNS=admin@*,*@security.domain.com

# Pattern for service accounts
PROTECTED_EMAIL_PATTERNS=admin@*,svc-*@*,automation-*@*
```

## API Reference

### AccountProtectionService

```typescript
// Create protection service
const protectionService = AccountProtectionService.fromEnvironment(graphClient);

// Check if account is protected
const result = await protectionService.isAccountProtected(
  'admin@domain.com',
  'user-id-123'
);

// Filter protected accounts
const { allowed, protectedAccounts } = await protectionService.filterProtectedAccounts([
  { email: 'admin@domain.com', userId: 'id-1' },
  { email: 'user@domain.com', userId: 'id-2' }
]);
```

### Protection Result

```typescript
interface ProtectedAccount {
  email: string;        // Email address
  reason: string;       // Why it's protected
  role?: string;        // Azure AD role (if applicable)
}
```

## Related Documentation

- [State Management](./STATE-MANAGEMENT.md)
- [Microsoft Graph API - Directory Roles](https://learn.microsoft.com/en-us/graph/api/resources/directoryrole)
- [Azure AD Built-in Roles](https://learn.microsoft.com/en-us/azure/active-directory/roles/permissions-reference)

## Summary

The Account Protection System provides **automatic, multi-layered protection** for critical accounts:

✅ **Email pattern matching** - Fast, no API calls
✅ **Explicit exclusion list** - Organization-specific protection
✅ **Azure AD role detection** - Automatic admin protection
✅ **No configuration required** - Works out of the box
✅ **Customizable** - Add patterns and exclusions as needed

**Default Protection**: All `admin@*` accounts and Global Administrators are protected automatically.

Your admin account is safe! 🛡️
