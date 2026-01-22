# Graph Connector People Data - Authentication Solution

**Status**: ✅ **RESOLVED** (2026-01-22)

## Summary

Microsoft Graph Connectors for People Data require **OAuth 2.0 Client Credentials Flow** (application-only authentication with client secret), not the Authorization Code Flow (delegated authentication with user sign-in).

## The Issue

We initially tried using delegated authentication (browser-based sign-in) with these scopes:
- `ExternalConnection.ReadWrite.All` (delegated)
- `ExternalItem.ReadWrite.All` (delegated)

**Results**:
- ✅ Connection management worked (list, get, create connections)
- ❌ Item ingestion failed with 401 Unauthenticated errors

Even though the token contained all required scopes and admin consent was granted, creating external items consistently returned:
```
401 Unauthenticated: The request has not been applied because it lacks valid
authentication credentials for the target resource.
```

## Root Cause

Graph Connectors operate as **background services** and require **Application Permissions**, not Delegated Permissions.

From Microsoft documentation:
> "Microsoft Graph connector agents run as background services and require Microsoft Graph application permissions. Delegated permissions aren't supported for connector agent registration and cause registration failures, even when the permissions appear correctly configured."

## The Solution

### Authentication: OAuth 2.0 Client Credentials Flow

Both approaches are OAuth 2.0, just different flows:

| Flow Type | Authentication Method | Use Case |
|-----------|----------------------|----------|
| **Authorization Code Flow** | User sign-in (delegated) | Interactive applications, user context needed |
| **Client Credentials Flow** | Client secret (application) | Background services, no user context |

Graph Connectors require **Client Credentials Flow**.

### Required Configuration

#### 1. Azure AD App Registration

**Permissions** (Application, not Delegated):
- `ExternalConnection.ReadWrite.OwnedBy` (Application)
- `ExternalItem.ReadWrite.OwnedBy` (Application)

**Setup Steps**:
1. Azure Portal > Azure AD > App registrations > Your app
2. API permissions > Add permission > Microsoft Graph
3. Select **Application permissions** (NOT Delegated)
4. Add:
   - `ExternalConnection.ReadWrite.OwnedBy`
   - `ExternalItem.ReadWrite.OwnedBy`
5. Grant admin consent

#### 2. Client Secret

1. Azure Portal > Your app > Certificates & secrets
2. New client secret
3. Copy the **Value** (not the Secret ID)
4. Add to `.env`:
   ```bash
   AZURE_CLIENT_SECRET=your-secret-value-here
   ```

### Implementation

The `enrich-profiles.ts` script automatically detects the client secret and uses the appropriate flow:

```typescript
const clientSecret = process.env.AZURE_CLIENT_SECRET;

if (clientSecret) {
  // Use OAuth 2.0 Client Credentials Flow (app-only)
  this.graphClient = new GraphClient({
    tenantId: this.tenantId,
    clientId: this.clientId,
    clientSecret: clientSecret
  });
} else {
  // Fallback to OAuth 2.0 Authorization Code Flow (delegated)
  // Note: This will fail for item ingestion
  const authServer = new BrowserAuthServer({ ... });
  const authResult = await authServer.authenticate();
  this.graphClient = new GraphClient({ accessToken: authResult.accessToken });
}
```

## Test Results

**Before** (Delegated Auth):
```
❌ Failed: 3/3 items
Error: 401 Unauthenticated
```

**After** (Application Auth with Client Secret):
```
✅ Successful: 3/3 items
✅ All properties ingested correctly
✅ Items verified in Graph Connector
```

### Verified Items

All 3 test users successfully ingested with:
- Email addresses (linked to Entra ID users)
- Skills (array properties with People Data labels)
- Interests (custom array properties)
- About Me descriptions (with People Data labels)
- Certifications, awards, projects (all with appropriate labels)

## Security Considerations

### Client Secret Management

**Storage**:
- Store in `.env` file (never commit to git)
- File permissions: 600 (owner read/write only)
- Consider using Azure Key Vault for production

**Rotation**:
- Set expiration period (6-24 months recommended)
- Create new secret before expiration
- Update `.env` with new value
- Delete old secret after transition

**Alternatives to Client Secrets**:

1. **Certificate-Based Authentication**:
   - More secure than client secrets
   - Upload certificate to Azure AD app
   - Configure app to use certificate instead of secret

2. **Managed Identity** (Azure only):
   - No secrets or certificates needed
   - Only works when running in Azure (App Service, Functions, etc.)
   - Not applicable for local CLI tools

## Key Learnings

1. **OAuth 2.0 has multiple flows**: Both client credentials and authorization code are OAuth 2.0 - Microsoft's statement "OAuth 2.0 required" was correct but ambiguous about which flow.

2. **Permission types matter**: Application permissions and delegated permissions are fundamentally different:
   - Delegated: Acts on behalf of a signed-in user
   - Application: Acts as the application itself (no user context)

3. **Graph Connectors are background services**: They don't operate in a user context, so they require application permissions.

4. **Connection management vs. item ingestion**: Connection operations can work with delegated auth, but item ingestion requires application auth.

5. **Error messages aren't always clear**: 401 Unauthenticated with "valid authentication credentials" missing doesn't indicate it's a permission type mismatch.

## Working Configuration

### .env File
```bash
AZURE_TENANT_ID=your-tenant-id
AZURE_CLIENT_ID=your-client-id
AZURE_CLIENT_SECRET=your-client-secret-value
```

### Azure AD Permissions (Application)
- ExternalConnection.ReadWrite.OwnedBy ✅
- ExternalItem.ReadWrite.OwnedBy ✅

### Usage
```bash
# Setup connection and schema (first time only)
npm run enrich-profiles:setup

# Ingest people data
npm run enrich-profiles -- --csv config/agents-template.csv

# Dry run (see what would be created)
npm run enrich-profiles:dry-run
```

## References

- [Microsoft Graph Permissions Reference](https://learn.microsoft.com/en-us/graph/permissions-reference)
- [OAuth 2.0 Client Credentials Flow](https://learn.microsoft.com/en-us/azure/active-directory/develop/v2-oauth2-client-creds-grant-flow)
- [Microsoft 365 Copilot Connectors for People Data](https://learn.microsoft.com/en-us/graph/peopleconnectors)
- [Resolve Authorization Errors](https://learn.microsoft.com/en-us/graph/resolve-auth-errors)

---

**Resolution Date**: 2026-01-22
**Implementation**: Option B - Unified Enrichment via Graph Connectors
**Status**: ✅ Working with OAuth 2.0 Client Credentials Flow
