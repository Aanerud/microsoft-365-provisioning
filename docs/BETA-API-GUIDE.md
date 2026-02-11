# Microsoft Graph BETA API Guide

## Overview

This guide covers the usage of Microsoft Graph **beta** endpoints in the M365-Agent-Provisioning project. Beta endpoints provide early access to new features and extended attributes not yet available in the stable v1.0 API.

## What is Microsoft Graph Beta?

Microsoft Graph has two API versions:

- **v1.0**: Stable, production-ready, backward-compatible
- **beta**: Preview features, subject to change, early access

### Key Differences

| Aspect | v1.0 | beta |
|--------|------|------|
| Stability | ✅ Guaranteed | ⚠️ May change |
| Support | ✅ Full support | ⚠️ Preview only |
| Features | Standard | Extended |
| Breaking Changes | Never | Possible |
| Production Use | ✅ Recommended | ⚠️ Use with caution |

## Beta Features Used in This Project

### 1. Extended User Attributes

Beta endpoints allow setting additional user properties during provisioning:

#### Available Attributes

- **`employeeType`**: Classification of employee
  - Examples: `Employee`, `Contractor`, `Intern`, `Consultant`, `Vendor`
  - Type: String
  - Max length: 256 characters

- **`companyName`**: Organization or company name
  - Examples: `Contoso Ltd`, `Fabrikam Inc`, `Adventure Works`
  - Type: String
  - Max length: 256 characters

- **`officeLocation`**: Physical office location
  - Examples: `Building 1`, `Oslo Office`, `Remote`, `New York - 5th Floor`
  - Type: String
  - Max length: 256 characters

#### Usage Example

```typescript
import { GraphClient } from './graph-client.js';

const client = new GraphClient({
  accessToken: 'your-access-token',
});

const user = await client.createUser({
  displayName: 'Sarah Chen',
  email: 'sarah.chen@domain.com',
  password: 'SecurePassword123!',
  employeeType: 'Employee',
  companyName: 'Contoso Ltd',
  officeLocation: 'Building 1',
});
```

### 2. Advanced Search API

The Microsoft Search API (beta) provides unified search across:

- Users (people)
- Emails (messages)
- Calendar events (events)
- Files (driveItems)

**Note**: Currently not fully implemented. Requires additional permission: `SearchConfiguration.Read.All`

### 3. Extended Profile Fields

Beta user endpoint returns additional fields:

```json
{
  "id": "user-id",
  "displayName": "Sarah Chen",
  "userPrincipalName": "sarah.chen@domain.com",
  "employeeType": "Employee",
  "companyName": "Contoso Ltd",
  "officeLocation": "Building 1",
  "department": "Executive",
  "jobTitle": "CEO"
}
```

## Enforcing Beta Endpoints

All Microsoft Graph calls in this project use **beta endpoints only**. Keep
`USE_BETA_ENDPOINTS=true` in `.env` for clarity, but there is no CLI flag or
per-call toggle.

```typescript
const client = new GraphClient({ accessToken: token });
```

## Beta Endpoints Reference

### User Operations

#### Create User with Extended Attributes

**Endpoint**: `POST /beta/users`

**Request Body**:
```json
{
  "displayName": "Sarah Chen",
  "userPrincipalName": "sarah.chen@domain.com",
  "mailNickname": "sarah.chen",
  "accountEnabled": true,
  "passwordProfile": {
    "password": "SecurePassword123!",
    "forceChangePasswordNextSignIn": false
  },
  "employeeType": "Employee",
  "companyName": "Contoso Ltd",
  "officeLocation": "Building 1"
}
```

**Response**: User object with extended attributes

#### Update User Extended Attributes

**Endpoint**: `PATCH /beta/users/{id}`

**Request Body**:
```json
{
  "employeeType": "Contractor",
  "companyName": "Fabrikam Inc",
  "officeLocation": "Remote"
}
```

#### Get User with Extended Attributes

**Endpoint**: `GET /beta/users/{id}`

**Query Parameters**:
```
?$select=id,displayName,userPrincipalName,mail,employeeType,companyName,officeLocation
```

### Search Operations (Preview)

#### Unified Search Query

**Endpoint**: `POST /beta/search/query`

**Request Body**:
```json
{
  "requests": [
    {
      "entityTypes": ["person"],
      "query": {
        "queryString": "Sarah Chen"
      },
      "from": 0,
      "size": 25
    }
  ]
}
```

**Note**: Requires `SearchConfiguration.Read.All` permission (not yet implemented)

## Beta-Only Behavior

This project does **not** fall back to v1.0. If a beta endpoint is unavailable,
the operation fails so you can address tenant configuration or permissions.

## CSV Configuration

### Standard CSV (core attributes)

```csv
name,email,role,department
Sarah Chen,sarah.chen@domain.com,CEO,Executive
John Smith,john.smith@domain.com,CTO,Technology
```

### Extended CSV (beta)

```csv
name,email,role,department,employeeType,companyName,officeLocation
Sarah Chen,sarah.chen@domain.com,CEO,Executive,Employee,Contoso Ltd,Building 1
John Smith,john.smith@domain.com,CTO,Technology,Employee,Contoso Ltd,Building 2
Jane Doe,jane.doe@domain.com,Consultant,Technology,Contractor,Fabrikam Inc,Remote
```

**Usage**:
```bash
npm run provision -- --csv config/agents-extended.csv
```

## Checking Beta Availability

### Programmatic Check

```typescript
import { GraphBetaClient } from './graph-beta-client.js';

const betaClient = new GraphBetaClient({ accessToken: token });
const available = await betaClient.isBetaAvailable();

if (available) {
  console.log('✅ Beta endpoints available');
} else {
  console.log('⚠️ Beta endpoints unavailable');
}
```

### CLI Check

```bash
node dist/graph-beta-client.js check-availability
```

## Error Handling

### Common Beta Errors

#### 1. Beta Not Available (404)

```
Error: Request failed with status code 404
Reason: Beta endpoint not available in tenant
Solution: Verify tenant supports beta endpoints and required permissions
```

#### 2. Invalid Beta Attribute (400)

```
Error: Invalid property 'employeeType'
Reason: Attribute not supported or misspelled
Solution: Verify attribute name and tenant support
```

#### 3. Permission Denied (403)

```
Error: Insufficient privileges
Reason: Missing required permissions for beta features
Solution: Grant additional permissions in Azure AD
```

### Handling in Code

```typescript
try {
  await betaClient.createUserWithExtendedAttributes(params);
} catch (error) {
  if (error.statusCode === 404) {
    console.error('Beta endpoint unavailable - check tenant support');
  } else if (error.statusCode === 403) {
    // Permission issue
    console.error('Permission denied - check Azure AD permissions');
  } else {
    // Unexpected error
    throw error;
  }
}
```

## Stability and Change Management

### Monitoring Beta Changes

Microsoft publishes changes to the Graph API in the [changelog](https://developer.microsoft.com/en-us/graph/changelog):

- Subscribe to notifications for beta API changes
- Review monthly updates from Microsoft
- Test beta features in non-production environment

### Migration Path

When beta features graduate to v1.0:

1. **Announcement**: Microsoft announces graduation timeline
2. **Testing**: Verify feature works in v1.0 endpoint
3. **Migration**: Update code to use v1.0 (remove beta flag)
4. **Deprecation**: Beta version of feature eventually removed

### Best Practices

1. **Log Beta Usage**: Track when beta endpoints are used
2. **Monitor Changelog**: Stay updated on beta changes
3. **Test Regularly**: Verify beta features still work as expected
4. **Document Dependencies**: Note which features require beta

## Performance Considerations

### Beta vs v1.0 Performance

- **Latency**: Beta endpoints may have slightly higher latency
- **Rate Limits**: Same rate limits as v1.0 (typically)
- **Throttling**: Beta may be throttled more aggressively during high load

### Optimization Tips

1. **Batch Requests**: Use `$batch` endpoint when possible
2. **Selective Queries**: Use `$select` to request only needed fields
3. **Caching**: Cache beta availability check results
4. **Rate Limiting**: Add delays between requests (500ms recommended)

## Code Examples

### Example 1: Provision User with Beta Attributes

```typescript
import { GraphClient } from './graph-client.js';
import { DeviceCodeAuth } from './auth/device-code-auth.js';

async function provisionWithBeta() {
  // Authenticate
  const auth = new DeviceCodeAuth({
    tenantId: process.env.AZURE_TENANT_ID!,
    clientId: process.env.AZURE_CLIENT_ID!,
  });
  const authResult = await auth.getAccessToken();

  // Create client (beta-only enforced)
  const client = new GraphClient({
    accessToken: authResult.accessToken,
  });

  // Provision user
  const user = await client.createUser({
    displayName: 'Sarah Chen',
    email: 'sarah.chen@contoso.com',
    password: 'SecurePass123!',
    employeeType: 'Employee',
    companyName: 'Contoso Ltd',
    officeLocation: 'Building 1',
  });

  console.log('User created:', user.id);
}
```

### Example 2: Bulk Provision with Beta

```typescript
import { GraphBetaClient } from './graph-beta-client.js';

async function bulkProvisionWithBeta() {
  const betaClient = new GraphBetaClient({
    accessToken: token,
  });

  const users = [
    {
      displayName: 'Sarah Chen',
      email: 'sarah.chen@contoso.com',
      attributes: {
        employeeType: 'Employee',
        companyName: 'Contoso Ltd',
        officeLocation: 'Building 1',
      },
    },
    {
      displayName: 'John Smith',
      email: 'john.smith@contoso.com',
      attributes: {
        employeeType: 'Contractor',
        companyName: 'Fabrikam Inc',
        officeLocation: 'Remote',
      },
    },
  ];

  const result = await betaClient.bulkUserProvision(users);

  console.log(`✅ Successful: ${result.successful.length}`);
  console.log(`❌ Failed: ${result.failed.length}`);
}
```

### Example 3: Check and Update Extended Attributes

```typescript
async function updateExtendedAttributes(userId: string) {
  const client = new GraphClient({
    accessToken: token,
  });

  // Check if beta is available
  const betaAvailable = await client.checkBetaAvailability();

  if (!betaAvailable) {
    console.warn('Beta endpoint unavailable - check tenant configuration');
    return;
  }

  await client.updateUserBeta(userId, {
    employeeType: 'Senior Employee',
    officeLocation: 'Building 2',
  });
}
```

## Troubleshooting

### Issue: Beta attributes not appearing in Azure AD

**Cause**: Azure AD portal may not display all beta attributes
**Solution**: Use Microsoft Graph Explorer to verify attributes are set

### Issue: Beta endpoints unavailable

**Cause**: Tenant or cloud does not support beta endpoints
**Solution**: Verify tenant supports beta endpoints; some government clouds may not

### Issue: Permission errors despite admin consent

**Cause**: Delegated permissions require re-consent after changes
**Solution**: Run `npm run provision -- --logout` then re-authenticate

## Additional Resources

- [Microsoft Graph Beta Reference](https://docs.microsoft.com/en-us/graph/api/overview?view=graph-rest-beta)
- [Microsoft Graph Changelog](https://developer.microsoft.com/en-us/graph/changelog)
- [User Resource Type (Beta)](https://docs.microsoft.com/en-us/graph/api/resources/user?view=graph-rest-beta)
- [Best Practices for Graph API](https://docs.microsoft.com/en-us/graph/best-practices-concept)

---

**Last Updated**: 2026-01-21
**API Version**: Microsoft Graph Beta
