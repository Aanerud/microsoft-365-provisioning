# MSAL Migration Guide

## Overview

This guide walks through migrating the M365-Agent-Provisioning project from **Client Secret authentication** to **MSAL Device Code Flow** authentication.

## Before vs After

### Before (Client Secret - v1.x)

```typescript
// Client Secret Credential (Application permissions)
const credential = new ClientSecretCredential(
  process.env.AZURE_TENANT_ID,
  process.env.AZURE_CLIENT_ID,
  process.env.AZURE_CLIENT_SECRET  // Secret stored in .env
);

const graphClient = new GraphClient();
await graphClient.createUser(params);
```

**Characteristics:**
- ❌ Client secret stored in `.env` file
- ❌ Application permissions (acts as app, not user)
- ❌ No audit trail of which admin performed actions
- ❌ Secret can be accidentally committed to git
- ✅ No user interaction required
- ✅ Works in automated pipelines

### After (MSAL Device Code Flow - v2.x)

```typescript
// Device Code Flow (Delegated permissions)
const auth = new DeviceCodeAuth({
  tenantId: process.env.AZURE_TENANT_ID,
  clientId: process.env.AZURE_CLIENT_ID,
  // No client secret needed!
});

const authResult = await auth.getAccessToken();

const graphClient = new GraphClient({
  accessToken: authResult.accessToken,
  useBeta: true,
});

await graphClient.createUser(params);
```

**Characteristics:**
- ✅ No client secrets stored anywhere
- ✅ Delegated permissions (acts on behalf of signed-in admin)
- ✅ Full audit trail (logs show which admin performed actions)
- ✅ Impossible to leak secrets
- ✅ Token caching (no repeated authentication)
- ✅ MFA support
- ⚠️ Requires initial browser authentication

## Migration Steps

### Step 1: Update Azure AD App Registration

#### 1.1 Configure as Public Client

1. Go to Azure Portal → Azure AD → App registrations → Your app
2. Navigate to **Authentication** in left menu
3. Scroll to **Advanced settings**
4. Under **Allow public client flows**:
   - Set to **YES**
5. Click **Save**

#### 1.2 Update Permissions

**Remove Application Permissions:**
1. Go to **API Permissions**
2. For each permission, click the three dots → **Remove permission**:
   - ❌ `User.ReadWrite.All` (Application)
   - ❌ `Directory.ReadWrite.All` (Application)
   - ❌ `Organization.Read.All` (Application)

**Add Delegated Permissions:**
1. Click **Add a permission** → **Microsoft Graph** → **Delegated permissions**
2. Add these permissions:
   - ✅ `User.ReadWrite.All` (Delegated)
   - ✅ `Directory.ReadWrite.All` (Delegated)
   - ✅ `Organization.Read.All` (Delegated)
   - ✅ `offline_access` (Delegated)
3. Click **Grant admin consent for [Your Organization]**
4. Verify all show green checkmarks

#### 1.3 Remove Client Secret (Optional)

Since client secrets are no longer needed, you can delete them:

1. Go to **Certificates & secrets**
2. For each client secret, click **Delete**
3. Confirm deletion

**Note**: If you need to maintain backward compatibility temporarily, keep the secret until migration is complete.

### Step 2: Update Environment Variables

#### 2.1 Backup Old .env

```bash
cp .env .env.backup-v1
```

#### 2.2 Update .env File

Remove the client secret line and add new variables:

**Before:**
```env
AZURE_TENANT_ID=your-tenant-id
AZURE_CLIENT_ID=your-client-id
AZURE_CLIENT_SECRET=your-client-secret  # ← Remove this
```

**After:**
```env
AZURE_TENANT_ID=your-tenant-id
AZURE_CLIENT_ID=your-client-id
# AZURE_CLIENT_SECRET removed - using MSAL Device Code Flow

GRAPH_API_ENDPOINT=https://graph.microsoft.com
USE_BETA_ENDPOINTS=true
```

### Step 3: Install New Dependencies

```bash
# Install MSAL Node
npm install @azure/msal-node@^2.13.0

# Verify installation
npm list @azure/msal-node
```

### Step 4: Update Code

The migration has already been implemented in the codebase. Key changes:

#### 4.1 Authentication Module

- New: `src/auth/device-code-auth.ts` - MSAL authentication
- New: `src/auth/token-cache.ts` - Token management

#### 4.2 Graph Client

- Updated: `src/graph-client.ts` - Now accepts `accessToken` parameter
- Backward compatible: Still supports client secret via environment variables

#### 4.3 Provision Script

- Updated: `src/provision.ts` - Integrated device code authentication at startup

### Step 5: Test Authentication Flow

#### 5.1 First Authentication

```bash
npm run provision
```

**Expected output:**
```
🔐 Microsoft 365 Authentication Required

📋 To sign in, use a web browser to open the page:
   https://microsoft.com/devicelogin

📝 And enter the code: A1B2C3D4

⏳ Waiting for authentication...
```

**Steps:**
1. Open browser to https://microsoft.com/devicelogin
2. Enter the displayed code (e.g., A1B2C3D4)
3. Sign in with your Microsoft 365 admin account
4. Grant consent to permissions
5. Return to CLI - it continues automatically

#### 5.2 Verify Token Caching

```bash
# Check token cache
npm run token-info
```

**Expected output:**
```
📦 Token Cache Information

Location: /Users/you/.m365-provision/token-cache.json
Exists: Yes
Accounts: 1
Token Expires: 1/21/2026, 11:30:00 AM

Status: ✅ Token valid (expires in ~1h)
```

#### 5.3 Verify Cached Token Usage

```bash
# Second run should use cached token (no re-authentication)
npm run provision
```

**Expected output:**
```
✅ Using cached authentication token

Loading agents from config/agents-template.csv...
```

### Step 6: Test Beta Endpoints

```bash
# Provision with beta features
npm run provision:beta

# Or with flag
npm run provision -- --use-beta
```

**Expected output:**
```
Configuration:
  Beta Features: ✓ Enabled

✓ Created user [beta]: Sarah Chen (sarah.chen@domain.com)
```

### Step 7: Test Logout

```bash
# Clear cached tokens
npm run logout
```

**Expected output:**
```
✅ Token cache cleared
```

## Backward Compatibility

The code maintains backward compatibility:

```typescript
// GraphClient supports both methods
const graphClient = new GraphClient({
  // Method 1: MSAL (new)
  accessToken: token,
  useBeta: true,
});

// Method 2: Client Secret (legacy - still works)
// Uses AZURE_CLIENT_SECRET from environment
const graphClient = new GraphClient();
```

**Deprecation Timeline:**
- v2.0: Both methods supported
- v3.0 (future): Client secret support may be removed

## Testing Checklist

Use this checklist to verify migration:

- [ ] Azure AD app configured as public client
- [ ] "Allow public client flows" set to YES
- [ ] Delegated permissions added (not application)
- [ ] Admin consent granted
- [ ] Client secret removed from `.env`
- [ ] `@azure/msal-node` installed
- [ ] First authentication prompts for device code
- [ ] Browser authentication completes successfully
- [ ] Token cached in `~/.m365-provision/`
- [ ] Subsequent runs use cached token
- [ ] Beta endpoints work with `--use-beta` flag
- [ ] User provisioning completes successfully
- [ ] License assignment works
- [ ] MCP tokens generated
- [ ] Output file created correctly
- [ ] Logout clears cache
- [ ] Re-authentication works after logout

## Troubleshooting

### Issue: "Public client flows not allowed"

**Cause**: App not configured as public client

**Solution**:
```
Azure AD → App registrations → Your app → Authentication
→ Advanced settings → "Allow public client flows" → YES
```

### Issue: "Insufficient privileges"

**Cause**: Using application permissions instead of delegated

**Solution**:
1. Remove application permissions
2. Add delegated permissions
3. Grant admin consent
4. Verify signed-in user has required Azure AD role

### Issue: "Device code expired"

**Cause**: User took >15 minutes to complete authentication

**Solution**:
- Run provision again to get new code
- Complete authentication within 15 minutes

### Issue: "Token not refreshing"

**Cause**: Cached refresh token expired or revoked

**Solution**:
```bash
npm run logout
npm run provision  # Will re-authenticate
```

### Issue: "Beta endpoints not working"

**Cause**: Beta features unavailable or not enabled

**Solution**:
- Verify `USE_BETA_ENDPOINTS=true` in `.env`
- Tool will automatically fall back to v1.0
- Check Microsoft Graph changelog for beta status

## Rollback Plan

If migration encounters issues, rollback to v1.x:

### Step 1: Restore Environment

```bash
cp .env.backup-v1 .env
```

### Step 2: Restore Azure AD App

1. Azure AD → App registrations → Your app → Authentication
2. Set "Allow public client flows" to **NO**
3. API Permissions → Remove delegated permissions
4. API Permissions → Add application permissions back
5. Grant admin consent

### Step 3: Create New Client Secret

1. Certificates & secrets → New client secret
2. Copy value to `AZURE_CLIENT_SECRET` in `.env`

### Step 4: Use Client Secret Auth

The code automatically falls back to client secret if available:

```typescript
// If AZURE_CLIENT_SECRET exists, uses client secret
const graphClient = new GraphClient();
```

## Benefits of Migration

### Security

| Aspect | Before (Client Secret) | After (MSAL) |
|--------|------------------------|--------------|
| Secret Storage | Environment variable | None |
| Secret Rotation | Manual (90 days) | N/A |
| Leak Risk | High (can be committed) | None |
| MFA Support | No | Yes |
| Conditional Access | No | Yes |

### Auditing

| Aspect | Before (Client Secret) | After (MSAL) |
|--------|------------------------|--------------|
| Audit Logs | "Application: M365-Agent-Provisioning" | "User: admin@domain.com" |
| User Context | No | Yes |
| Action Attribution | Application-level | User-specific |

### Operations

| Aspect | Before (Client Secret) | After (MSAL) |
|--------|------------------------|--------------|
| Setup Complexity | Medium | Medium |
| First Run | Automatic | Requires authentication |
| Subsequent Runs | Automatic | Automatic (cached) |
| Token Expiration | N/A | Auto-refresh (90 days) |
| Revocation | Delete secret | Revoke user session |

## Additional Resources

- [Device Code Flow Documentation](./DEVICE-CODE-FLOW.md)
- [MSAL Node Repository](https://github.com/AzureAD/microsoft-authentication-library-for-js/tree/dev/lib/msal-node)
- [OAuth 2.0 Device Authorization Grant](https://tools.ietf.org/html/rfc8628)
- [Microsoft Identity Platform](https://docs.microsoft.com/en-us/azure/active-directory/develop/)

## Support

If you encounter issues during migration:

1. Check [SETUP.md](../SETUP.md) troubleshooting section
2. Verify Azure AD configuration
3. Check token cache: `npm run token-info`
4. Enable MSAL debug logging in `device-code-auth.ts`
5. Review Azure AD sign-in logs in portal

---

**Migration Version**: 1.x → 2.x
**Last Updated**: 2026-01-21
**MSAL Version**: @azure/msal-node ^2.13.0
