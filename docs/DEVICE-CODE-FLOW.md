# Device Code Flow - Technical Documentation

## Overview

This document provides technical details about the Device Code Flow implementation in the M365-Agent-Provisioning project using MSAL (Microsoft Authentication Library) for Node.js.

## What is Device Code Flow?

**Device Code Flow** (also known as Device Authorization Grant) is an OAuth 2.0 authentication flow designed for devices with limited input capabilities or CLI applications.

### Use Cases

- **CLI Applications**: Command-line tools like Azure CLI (`az login`)
- **Headless Servers**: Servers without browser access
- **IoT Devices**: Smart TVs, printers, or devices without keyboards
- **Remote Authentication**: User can authenticate on a different device

### How It Works

```
┌─────────────┐                  ┌─────────────┐                 ┌─────────────┐
│             │                  │             │                 │             │
│  CLI App    │                  │  Azure AD   │                 │   Browser   │
│             │                  │             │                 │ (User Auth) │
└──────┬──────┘                  └──────┬──────┘                 └──────┬──────┘
       │                                │                                │
       │ 1. Request Device Code         │                                │
       │───────────────────────────────>│                                │
       │                                │                                │
       │ 2. Return Device Code + URL    │                                │
       │<───────────────────────────────│                                │
       │                                │                                │
       │ 3. Display URL & Code to User  │                                │
       │                                │                                │
       │                                │    4. User Opens URL           │
       │                                │<───────────────────────────────│
       │                                │                                │
       │                                │    5. User Enters Code         │
       │                                │<───────────────────────────────│
       │                                │                                │
       │                                │    6. User Signs In            │
       │                                │<───────────────────────────────│
       │                                │                                │
       │                                │    7. User Grants Consent      │
       │                                │<───────────────────────────────│
       │                                │                                │
       │ 8. Poll for Authorization      │                                │
       │───────────────────────────────>│                                │
       │                                │                                │
       │ 9. Return Access Token         │                                │
       │<───────────────────────────────│                                │
       │                                │                                │
       │ 10. Make API Calls with Token  │                                │
       │───────────────────────────────>│                                │
       │                                │                                │
```

## Implementation Details

### 1. MSAL Configuration

```typescript
import * as msal from '@azure/msal-node';

const msalConfig: msal.Configuration = {
  auth: {
    clientId: process.env.AZURE_CLIENT_ID!,
    authority: `https://login.microsoftonline.com/${process.env.AZURE_TENANT_ID}`,
  },
  cache: {
    cachePlugin: createCachePlugin(), // Persistent token storage
  },
  system: {
    loggerOptions: {
      loggerCallback: (level, message, containsPii) => {
        if (!containsPii && (level === msal.LogLevel.Error || level === msal.LogLevel.Warning)) {
          console.error(`[MSAL] ${message}`);
        }
      },
      piiLoggingEnabled: false,
      logLevel: msal.LogLevel.Error,
    },
  },
};

const publicClientApp = new msal.PublicClientApplication(msalConfig);
```

### 2. Device Code Request

```typescript
const deviceCodeRequest: msal.DeviceCodeRequest = {
  deviceCodeCallback: (response) => {
    // Display to user
    console.log(`Visit: ${response.verificationUri}`);
    console.log(`Code: ${response.userCode}`);
  },
  scopes: [
    'User.ReadWrite.All',
    'Directory.ReadWrite.All',
    'Organization.Read.All',
    'offline_access',
  ],
};

const response = await publicClientApp.acquireTokenByDeviceCode(deviceCodeRequest);
```

### 3. Token Response

```typescript
interface AuthenticationResult {
  accessToken: string;        // Bearer token for API calls
  expiresOn: Date;            // Token expiration time
  account: msal.AccountInfo;  // User account information
  scopes: string[];           // Granted scopes
  tokenType: string;          // Always "Bearer"
}
```

### 4. Token Caching

Tokens are cached to avoid repeated authentication:

```typescript
const cachePlugin: msal.ICachePlugin = {
  beforeCacheAccess: async (cacheContext) => {
    if (fs.existsSync(cacheFilePath)) {
      const cacheData = fs.readFileSync(cacheFilePath, 'utf-8');
      cacheContext.tokenCache.deserialize(cacheData);
    }
  },
  afterCacheAccess: async (cacheContext) => {
    if (cacheContext.cacheHasChanged) {
      const cacheData = cacheContext.tokenCache.serialize();
      fs.writeFileSync(cacheFilePath, cacheData, { mode: 0o600 });
    }
  },
};
```

### 5. Silent Token Acquisition

For subsequent requests, try to use cached tokens:

```typescript
const silentRequest: msal.SilentFlowRequest = {
  account: cachedAccount,
  scopes: requestedScopes,
  forceRefresh: false,
};

try {
  const response = await publicClientApp.acquireTokenSilent(silentRequest);
  // Use response.accessToken
} catch (error) {
  // Token expired or invalid - fall back to device code flow
  await authenticateDeviceCode();
}
```

## Security Considerations

### 1. No Client Secrets

**Why it's secure:**
- No secrets stored in environment variables or files
- All authentication happens via secure browser flow
- Secrets can't be accidentally committed to git

### 2. Delegated Permissions

**Benefits:**
- All API operations performed on behalf of signed-in user
- Audit logs show which admin performed actions
- Can be revoked instantly by Azure AD admin
- Respects user's actual permissions (no over-privileged app)

### 3. Token Storage Security

**Protection measures:**
- Cache file permissions: `0o600` (owner read/write only)
- Cache location: `~/.m365-provision/` (user's home directory)
- Refresh tokens encrypted by MSAL library
- Tokens expire automatically (1 hour for access token)

### 4. Token Expiration

| Token Type | Lifetime | Refresh Behavior |
|------------|----------|------------------|
| Access Token | 1 hour | Automatic via refresh token |
| Refresh Token | 90 days | Extended on use (rolling refresh) |
| Device Code | 15 minutes | One-time use only |

### 5. MFA Support

Device Code Flow fully supports:
- Multi-Factor Authentication (MFA)
- Conditional Access Policies
- Security defaults
- Azure AD Identity Protection

## Error Handling

### Common Errors

#### 1. `authorization_pending`

**Cause**: User hasn't completed authentication yet
**Handling**: Continue polling (MSAL handles this automatically)

#### 2. `expired_token`

**Cause**: Device code expired (15-minute timeout)
**Handling**: Request new device code and display to user

```typescript
if (error.errorCode === 'expired_token') {
  console.error('Device code expired. Please try again.');
  return await authenticateDeviceCode(); // Retry
}
```

#### 3. `interaction_required`

**Cause**: Token refresh failed, user must re-authenticate
**Handling**: Fall back to interactive device code flow

```typescript
if (error.errorCode === 'interaction_required') {
  console.log('Re-authentication required');
  return await authenticateDeviceCode();
}
```

#### 4. `consent_required`

**Cause**: User needs to grant additional permissions
**Handling**: Initiate new device code flow (consent screen will appear)

```typescript
if (error.errorCode === 'consent_required') {
  console.log('Additional permissions required');
  return await authenticateDeviceCode();
}
```

## Performance Characteristics

### Latency

| Operation | Typical Duration |
|-----------|------------------|
| Device code request | 200-500ms |
| User authentication (manual) | 30-120 seconds |
| Silent token acquisition | 50-200ms |
| Token refresh | 200-500ms |

### Polling Behavior

MSAL automatically polls Azure AD:
- **Interval**: 5 seconds
- **Timeout**: 15 minutes
- **Retry**: Exponential backoff on errors

## Comparison with Other Flows

| Feature | Device Code Flow | Client Credentials | Authorization Code |
|---------|------------------|--------------------|--------------------|
| Use Case | CLI/Headless | Background services | Web applications |
| User Context | Yes (delegated) | No (application) | Yes (delegated) |
| Browser Required | Yes (once) | No | Yes (always) |
| Client Secret | No | Yes | Yes |
| MFA Support | Yes | N/A | Yes |
| Audit Trail | User-specific | Application-only | User-specific |
| Token Caching | Yes | N/A | Yes |

## Best Practices

### 1. Token Management

```typescript
// ✅ Good: Check cache first
const token = await auth.getAccessToken(false);

// ❌ Bad: Always force new authentication
const token = await auth.getAccessToken(true);
```

### 2. Error Handling

```typescript
// ✅ Good: Handle specific error codes
try {
  await auth.getAccessToken();
} catch (error) {
  if (error.errorCode === 'interaction_required') {
    // Fallback to device code
  } else {
    throw error; // Re-throw unexpected errors
  }
}

// ❌ Bad: Catch all errors silently
try {
  await auth.getAccessToken();
} catch {
  // Silent failure - no way to diagnose issues
}
```

### 3. Logout Implementation

```typescript
// ✅ Good: Clear both cache and in-memory tokens
async clearCache() {
  // Remove cache file
  if (fs.existsSync(cacheFilePath)) {
    fs.unlinkSync(cacheFilePath);
  }

  // Clear in-memory cache
  const accounts = await cache.getAllAccounts();
  for (const account of accounts) {
    await cache.removeAccount(account);
  }
}

// ❌ Bad: Only remove cache file
async clearCache() {
  fs.unlinkSync(cacheFilePath);
  // In-memory tokens still present!
}
```

### 4. Secure Token Storage

```typescript
// ✅ Good: Secure file permissions
fs.writeFileSync(cacheFilePath, cacheData, { mode: 0o600 });

// ❌ Bad: World-readable cache file
fs.writeFileSync(cacheFilePath, cacheData);
```

## Testing Strategies

### 1. Test Authentication Flow

```bash
# Test first-time authentication
npm run provision -- --auth

# Verify output shows device code
# Complete authentication in browser
# Verify token is cached
```

### 2. Test Token Caching

```bash
# First run (authenticates)
npm run provision

# Second run (should use cache)
npm run provision  # Should not prompt for authentication
```

### 3. Test Token Refresh

```bash
# Check current token expiration
node dist/auth/token-cache.js info

# Wait for token to expire (1 hour)
# Run again - should auto-refresh
npm run provision
```

### 4. Test Logout

```bash
# Logout
npm run provision -- --logout

# Verify cache cleared
ls ~/.m365-provision/  # Should be empty or no token-cache.json

# Next run should re-authenticate
npm run provision
```

## Debugging

### Enable MSAL Logging

```typescript
system: {
  loggerOptions: {
    logLevel: msal.LogLevel.Verbose, // Enable debug logging
    piiLoggingEnabled: false,        // Never enable in production!
  },
}
```

### Inspect Token Cache

```bash
# View cache contents (formatted)
node dist/auth/token-cache.js info

# Raw cache contents
cat ~/.m365-provision/token-cache.json | jq .
```

### Network Debugging

```bash
# Use environment variable to log HTTP requests
NODE_DEBUG=http,https npm run provision
```

## Azure AD Configuration

### Required App Settings

1. **Platform**: Public client / Native application
2. **Allow public client flows**: YES (critical!)
3. **Redirect URIs**: Not needed for device code flow
4. **Permissions**: Delegated (NOT application)
5. **Admin Consent**: Required for elevated permissions

### Verification

```bash
# Check app configuration in Azure Portal
# Azure AD → App registrations → Your app → Authentication
# Verify "Allow public client flows" is enabled
```

## References

- [OAuth 2.0 Device Authorization Grant](https://tools.ietf.org/html/rfc8628)
- [MSAL Node Documentation](https://github.com/AzureAD/microsoft-authentication-library-for-js/tree/dev/lib/msal-node)
- [Microsoft Identity Platform](https://docs.microsoft.com/en-us/azure/active-directory/develop/)
- [Device Code Flow Guide](https://docs.microsoft.com/en-us/azure/active-directory/develop/v2-oauth2-device-code)

---

**Last Updated**: 2026-01-21
**MSAL Version**: @azure/msal-node ^2.13.0
