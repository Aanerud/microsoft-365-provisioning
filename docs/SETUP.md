# Setup Guide

## Prerequisites

- Node.js 18 or later
- npm or yarn package manager
- Azure AD administrator access
- Microsoft 365 tenant with available licenses

## Azure AD App Registration

### 1. Register Application

1. Navigate to [Azure Portal](https://portal.azure.com)
2. Go to **Azure Active Directory** → **App registrations** → **New registration**
3. Name: `M365-Agent-Provisioning`
4. Supported account types: **Accounts in this organizational directory only**
5. **Redirect URI**: Leave blank (not needed for device code flow)
6. Click **Register**

### 2. Configure as Public Client (IMPORTANT)

**This step is critical for device code flow authentication:**

1. Go to **Authentication** in the left menu
2. Scroll to **Advanced settings** section
3. Under **Allow public client flows**:
   - Set **"Enable the following mobile and desktop flows"** to **YES**
4. Click **Save**

**Why this matters**: Device code flow requires the app to be configured as a public client (native/mobile application). Without this setting, authentication will fail.

### 3. Configure API Permissions

Add the following **Delegated permissions** (NOT application permissions):

#### Microsoft Graph API
- `User.ReadWrite.All` - Create and manage users (requires admin consent)
- `Directory.ReadWrite.All` - Manage directory objects (requires admin consent)
- `Organization.Read.All` - Read organization info (requires admin consent)
- `offline_access` - Maintain access via refresh tokens

#### After adding permissions:
1. Click **Grant admin consent for [Your Organization]**
2. Verify all permissions show green checkmarks with "Granted for [Organization]"

**Key Differences from Application Permissions:**
- **Delegated permissions**: Act on behalf of signed-in user (audit trail shows who performed actions)
- **Application permissions**: Act as the application itself (no user context)
- Device code flow requires delegated permissions

### 4. Note Required Values

Copy these values for your `.env` file:
- **Application (client) ID**: Found on Overview page
- **Directory (tenant) ID**: Found on Overview page

**Note**: You do NOT need a client secret for device code flow! The authentication happens via browser sign-in.

## Environment Configuration

### 1. Copy Environment Template

```bash
cp .env.example .env
```

### 2. Edit .env File

```env
# Azure AD Configuration (Device Code Flow - No Client Secret Needed!)
AZURE_TENANT_ID=your-tenant-id-here
AZURE_CLIENT_ID=your-client-id-here
# NOTE: AZURE_CLIENT_SECRET is NOT needed - authentication via browser

# Microsoft Graph API Configuration
GRAPH_API_ENDPOINT=https://graph.microsoft.com
USE_BETA_ENDPOINTS=true  # Enable beta features for extended user attributes

# Microsoft 365 License SKU ID
# E3: 05e9a617-0261-4cee-bb44-138d3ef5d965
# E5: 06ebc4ee-1bb5-47dd-8120-11324bc54e06
LICENSE_SKU_ID=05e9a617-0261-4cee-bb44-138d3ef5d965

# User Provisioning Settings
USER_PASSWORD_PROFILE=Auto  # Auto-generate secure passwords
USER_DOMAIN=yourdomain.onmicrosoft.com
```

**Important Notes:**
- **No client secret needed**: Authentication uses device code flow (browser-based)
- **Beta endpoints**: Enabled by default for extended attributes (employeeType, companyName, officeLocation)
- **Token caching**: Tokens are cached in `~/.m365-provision/` for convenience

### 3. Update User Domain

Replace `yourdomain.onmicrosoft.com` with your actual Microsoft 365 domain:
- Find your domain: Azure AD → Custom domain names
- Use the `.onmicrosoft.com` domain or your verified custom domain

## License SKU IDs

To find your organization's available license SKUs:

```bash
# After setup, run this helper script:
npm run list-licenses
```

Common SKU IDs:
- **Microsoft 365 E3**: `05e9a617-0261-4cee-bb44-138d3ef5d965`
- **Microsoft 365 E5**: `06ebc4ee-1bb5-47dd-8120-11324bc54e06`
- **Office 365 E3**: `6fd2c87f-b296-42f0-b197-1e91e994b900`

## Installation

```bash
# Install Node.js dependencies
npm install

# Verify installation
npm run test-connection
```

## Verify Setup

### Test Azure AD Connection

```bash
npm run test-connection
```

Expected output:
```
✓ Connected to Azure AD
✓ Verified Graph API permissions
✓ Found X available licenses
```

## Admin Role Requirements

**Important**: The user signing in during device code authentication must have one of these Azure AD roles:

- **Global Administrator** (full admin access)
- **User Administrator** (can create users and assign licenses)
- **Privileged Role Administrator** (elevated permissions)

Without the proper role, provisioning will fail with "Insufficient privileges" error.

## Troubleshooting

### Error: "Invalid client" or "Public client flows not allowed"

**Cause**: App not configured as public client

**Solution**:
1. Go to Azure AD → App registrations → Your app → Authentication
2. Under "Advanced settings" → "Allow public client flows" → Set to **YES**
3. Click **Save**
4. Try authenticating again

### Error: "Insufficient privileges"

**Cause**: Missing API permissions, admin consent not granted, or user lacks required role

**Solution**:
1. Verify all **delegated** permissions are added (not application permissions)
2. Click "Grant admin consent" in Azure Portal
3. Verify signed-in user has Global Administrator or User Administrator role
4. Wait 5-10 minutes for permissions to propagate
5. Run `npm run provision -- --logout` then try again

### Error: "Device code expired"

**Cause**: User took too long to complete browser authentication (15-minute timeout)

**Solution**:
1. Run the provision command again to get a new device code
2. Complete authentication within 15 minutes
3. Use `--auth` flag if you want to force new authentication: `npm run provision -- --auth`

### Error: "Token refresh failed"

**Cause**: Cached refresh token expired (90 days) or was revoked

**Solution**:
1. Clear cached tokens: `npm run provision -- --logout`
2. Re-authenticate: `npm run provision`
3. Complete browser authentication when prompted

### Error: "License not found"

**Cause**: Incorrect LICENSE_SKU_ID or no licenses available

**Solution**:
1. After authenticating, run to see available SKUs (uses the GraphClient CLI)
2. Update `LICENSE_SKU_ID` in `.env`
3. Contact your Microsoft 365 admin to purchase more licenses if needed

### Error: "Domain not verified"

**Cause**: USER_DOMAIN doesn't match a verified domain in your tenant

**Solution**:
1. Use your `.onmicrosoft.com` domain (always verified)
2. Or verify custom domain in Azure AD → Custom domain names

### Error: "Beta endpoint unavailable"

**Cause**: Beta endpoints not available in tenant or temporarily unavailable

**Solution**:
- Beta endpoints are required; there is no v1.0 fallback
- Verify tenant supports beta endpoints and required permissions
- Check [Microsoft Graph changelog](https://developer.microsoft.com/en-us/graph/changelog) for updates

## Next Steps

Once setup is complete:
1. Review [USAGE.md](./USAGE.md) for creating agent definitions
2. Edit `config/agents-template.csv` with your desired agents
3. Run `npm run provision` to create users and generate tokens
