# Graph Connector Permissions Guide

## Issue Discovered During Testing

When using **delegated authentication** (user signs in via browser), Graph Connector requires different permissions than documented initially.

## Required Permissions for Delegated Authentication

| Permission | Type | Why Needed |
|------------|------|------------|
| `ExternalConnection.ReadWrite.All` | Delegated | Create and manage connections (delegated context) |
| `ExternalItem.ReadWrite.All` | Delegated | Ingest external items (delegated context) |
| `User.Read.All` | Delegated | Read user information for linking |
| `Directory.Read.All` | Delegated | Read directory for user lookups |

## Why Not `.OwnedBy`?

The `.OwnedBy` permissions only work for **application-owned connections** (using client credentials/app-only auth):

- `ExternalConnection.ReadWrite.OwnedBy` ❌ (doesn't work with delegated auth)
- `ExternalItem.ReadWrite.OwnedBy` ❌ (doesn't work with delegated auth)

Since this tool uses **delegated authentication** (user signs in), we need the `.All` scopes instead.

## How to Grant the Correct Permissions

### Azure Portal

1. Go to [Azure Portal](https://portal.azure.com)
2. Navigate to **Azure Active Directory** → **App registrations**
3. Select your application
4. Click **API permissions** in left menu
5. Remove the incorrect permissions if present:
   - `ExternalConnection.ReadWrite.OwnedBy` ❌
   - `ExternalItem.ReadWrite.OwnedBy` ❌
6. Click **+ Add a permission**
7. Select **Microsoft Graph** → **Delegated permissions**
8. Search for and add:
   - `ExternalConnection.ReadWrite.All` ✅
   - `ExternalItem.ReadWrite.All` ✅
9. Click **Grant admin consent for [Your Tenant]**
10. Wait for green checkmarks

### After Granting Permissions

1. **Logout** to clear cached token:
   ```bash
   npm run logout
   ```

2. **Wait for schema** to be ready (if not already):
   ```bash
   npm run enrich-profiles:wait
   ```

3. **Run ingestion**:
   ```bash
   npm run enrich-profiles -- --csv config/agents-test-enrichment.csv
   ```

## Testing Permissions

Use this command to test if permissions are working:

```bash
node test-permissions.mjs
```

Expected output when permissions are correct:
```
✅ Test 1: List all external connections - Success
✅ Test 2: Get specific connection - Success
✅ Test 3: Create a simple test item - Success
```

## Current Status

- ✅ Connection created (`m365provisionpeople`)
- ✅ Schema registered and ready
- ⚠️ Item ingestion failing with 403 (permission issue)
- 🔧 Need to grant `.All` permissions instead of `.OwnedBy`

## Alternative: Use Application-Only Authentication

If you prefer to use `.OwnedBy` permissions, you would need to:

1. Use **client credentials** instead of delegated auth
2. Store `AZURE_CLIENT_SECRET` in `.env`
3. Connections would be owned by the application
4. No user sign-in required

However, this is NOT the current architecture of this tool (which uses delegated/browser auth).
