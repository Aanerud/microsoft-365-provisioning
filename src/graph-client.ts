import { Client } from '@microsoft/microsoft-graph-client';
import { ClientSecretCredential } from '@azure/identity';
import { TokenCredentialAuthenticationProvider } from '@microsoft/microsoft-graph-client/authProviders/azureTokenCredentials/index.js';
import dotenv from 'dotenv';

dotenv.config();

interface UserCreateParams {
  displayName: string;
  userPrincipalName: string;
  mailNickname: string;
  accountEnabled: boolean;
  passwordProfile: {
    password: string;
    forceChangePasswordNextSignIn: boolean;
  };
  // Standard fields (optional)
  usageLocation?: string;
  // Beta-only fields (optional)
  employeeType?: string;
  companyName?: string;
  officeLocation?: string;
}

interface User {
  id: string;
  displayName: string;
  userPrincipalName: string;
  mail?: string;
  // Beta-only fields (optional)
  employeeType?: string;
  companyName?: string;
  officeLocation?: string;
}

interface License {
  skuId: string;
  skuPartNumber: string;
  consumedUnits: number;
  prepaidUnits: {
    enabled: number;
  };
}

export interface GraphClientConfig {
  // Option 1: Use delegated token (MSAL Device Code Flow)
  accessToken?: string;

  // Option 2: Use client secret (legacy)
  tenantId?: string;
  clientId?: string;
  clientSecret?: string;

  // Optional: Enable beta endpoints
  useBeta?: boolean;
}

export class GraphClient {
  private client!: Client;
  private betaClient!: Client;
  private credential?: ClientSecretCredential;
  private useBeta: boolean;

  constructor(config?: GraphClientConfig) {
    this.useBeta = config?.useBeta || process.env.USE_BETA_ENDPOINTS === 'true' || false;

    // Priority 1: Use provided access token (MSAL)
    if (config?.accessToken) {
      this.initializeWithAccessToken(config.accessToken);
    }
    // Priority 2: Use provided credentials
    else if (config?.tenantId && config?.clientId && config?.clientSecret) {
      this.initializeWithClientSecret(config.tenantId, config.clientId, config.clientSecret);
    }
    // Priority 3: Use environment variables (backward compatibility)
    else {
      const tenantId = process.env.AZURE_TENANT_ID;
      const clientId = process.env.AZURE_CLIENT_ID;
      const clientSecret = process.env.AZURE_CLIENT_SECRET;

      // If client secret exists, use client secret auth (legacy)
      if (tenantId && clientId && clientSecret) {
        this.initializeWithClientSecret(tenantId, clientId, clientSecret);
      } else {
        throw new Error(
          'GraphClient requires either:\n' +
          '  1. accessToken in config (MSAL Device Code Flow), or\n' +
          '  2. AZURE_CLIENT_SECRET in environment (legacy)\n' +
          'For MSAL authentication, pass { accessToken: "..." } to constructor.'
        );
      }
    }
  }

  /**
   * Initialize client with access token (MSAL Device Code Flow)
   */
  private initializeWithAccessToken(accessToken: string): void {
    const authProvider = (done: any) => {
      done(null, accessToken);
    };

    this.client = Client.init({
      authProvider,
      defaultVersion: 'v1.0',
    });

    this.betaClient = Client.init({
      authProvider,
      defaultVersion: 'beta',
    });
  }

  /**
   * Initialize client with client secret (legacy)
   */
  private initializeWithClientSecret(tenantId: string, clientId: string, clientSecret: string): void {
    this.credential = new ClientSecretCredential(tenantId, clientId, clientSecret);

    const authProvider = new TokenCredentialAuthenticationProvider(this.credential, {
      scopes: ['https://graph.microsoft.com/.default'],
    });

    this.client = Client.initWithMiddleware({
      authProvider,
    });

    this.betaClient = Client.initWithMiddleware({
      authProvider,
    });
  }

  /**
   * Get the appropriate client (v1.0 or beta)
   */
  private getClient(forceBeta: boolean = false): Client {
    return (this.useBeta || forceBeta) ? this.betaClient : this.client;
  }

  /**
   * Create a new user in Azure AD
   */
  async createUser(params: {
    displayName: string;
    email: string;
    password?: string;
    // Beta-only fields
    employeeType?: string;
    companyName?: string;
    officeLocation?: string;
  }): Promise<User> {
    const mailNickname = params.email.split('@')[0];
    const password = params.password || this.generatePassword();

    const userParams: UserCreateParams = {
      displayName: params.displayName,
      userPrincipalName: params.email,
      mailNickname,
      accountEnabled: true,
      passwordProfile: {
        password,
        forceChangePasswordNextSignIn: false,
      },
    };

    // Add beta fields if provided and beta is enabled
    const hasBetaFields = params.employeeType || params.companyName || params.officeLocation;
    if (hasBetaFields) {
      if (params.employeeType) userParams.employeeType = params.employeeType;
      if (params.companyName) userParams.companyName = params.companyName;
      if (params.officeLocation) userParams.officeLocation = params.officeLocation;
    }

    try {
      // Use beta client if beta fields are provided
      const client = hasBetaFields ? this.getClient(true) : this.getClient(false);
      const user = await client.api('/users').post(userParams);

      const betaIndicator = hasBetaFields ? ' [beta]' : '';
      console.log(`✓ Created user${betaIndicator}: ${params.displayName} (${params.email})`);
      return user;
    } catch (error: any) {
      if (error.code === 'Request_ResourceAlreadyExists') {
        console.log(`⚠ User already exists: ${params.email}`);
        const existingUser = await this.getUserByEmail(params.email);
        return existingUser;
      }

      // If beta endpoint failed, try falling back to v1.0
      if (hasBetaFields && error.statusCode === 400) {
        console.warn(`⚠ Beta endpoint failed, falling back to v1.0 without extended attributes`);
        return await this.createUser({
          displayName: params.displayName,
          email: params.email,
          password: params.password,
        });
      }

      throw error;
    }
  }

  /**
   * Get user by email
   */
  async getUserByEmail(email: string): Promise<User> {
    const user = await this.client.api(`/users/${email}`).get();
    return user;
  }

  /**
   * Get user details by ID
   */
  async getUserDetails(userId: string): Promise<User> {
    const user = await this.client.api(`/users/${userId}`).get();
    return user;
  }

  /**
   * Assign license to user
   */
  async assignLicense(userId: string, skuId: string): Promise<void> {
    try {
      await this.client.api(`/users/${userId}/assignLicense`).post({
        addLicenses: [
          {
            skuId,
          },
        ],
        removeLicenses: [],
      });
      console.log(`✓ Assigned license to user: ${userId}`);
    } catch (error: any) {
      if (error.code === 'Request_InvalidValue') {
        console.log(`⚠ License already assigned or invalid SKU: ${userId}`);
      } else {
        throw error;
      }
    }
  }

  /**
   * List all available licenses in the tenant
   */
  async listLicenses(): Promise<License[]> {
    const response = await this.client.api('/subscribedSkus').get();
    return response.value;
  }

  /**
   * List all users (filtered by user type if needed)
   */
  async listUsers(filter?: string): Promise<User[]> {
    let request = this.client.api('/users').select('id,displayName,userPrincipalName,mail');

    if (filter) {
      request = request.filter(filter);
    }

    const response = await request.get();
    return response.value;
  }

  /**
   * Check if mailbox is provisioned for user
   */
  async checkMailbox(userId: string): Promise<boolean> {
    try {
      await this.client.api(`/users/${userId}/mailFolders`).get();
      return true;
    } catch (error: any) {
      if (error.code === 'MailboxNotEnabledForRESTAPI') {
        return false;
      }
      throw error;
    }
  }

  /**
   * Delete user
   */
  async deleteUser(userId: string): Promise<void> {
    await this.client.api(`/users/${userId}`).delete();
    console.log(`✓ Deleted user: ${userId}`);
  }

  /**
   * Test connection to Microsoft Graph
   */
  async testConnection(): Promise<void> {
    try {
      const org = await this.client.api('/organization').get();
      console.log(`✓ Connected to Azure AD`);
      console.log(`  Organization: ${org.value[0].displayName}`);

      const licenses = await this.listLicenses();
      console.log(`✓ Found ${licenses.length} license types`);

      for (const license of licenses) {
        const available = license.prepaidUnits.enabled - license.consumedUnits;
        console.log(`  - ${license.skuPartNumber}: ${available} available`);
      }
    } catch (error: any) {
      console.error(`✗ Connection failed: ${error.message}`);
      throw error;
    }
  }

  /**
   * Update user with beta attributes
   *
   * Updates extended user attributes that are only available in beta endpoints.
   * Falls back gracefully if beta is not available.
   */
  async updateUserBeta(userId: string, updates: {
    employeeType?: string;
    companyName?: string;
    officeLocation?: string;
  }): Promise<void> {
    try {
      const client = this.getClient(true); // Force beta
      await client.api(`/users/${userId}`).patch(updates);
      console.log(`✓ Updated user [beta]: ${userId}`);
    } catch (error: any) {
      if (error.statusCode === 404 || error.message?.includes('beta')) {
        console.warn(`⚠ Beta endpoint unavailable, skipping extended attribute updates`);
      } else {
        throw error;
      }
    }
  }

  /**
   * Get user with beta fields
   *
   * Retrieves user with extended attributes from beta endpoint.
   * Falls back to v1.0 if beta is unavailable.
   */
  async getUserBeta(userId: string): Promise<User> {
    try {
      const client = this.getClient(true); // Force beta
      const user = await client
        .api(`/users/${userId}`)
        .select('id,displayName,userPrincipalName,mail,employeeType,companyName,officeLocation')
        .get();
      return user;
    } catch (error: any) {
      if (error.statusCode === 404 || error.message?.includes('beta')) {
        console.warn(`⚠ Beta endpoint unavailable, falling back to v1.0`);
        return await this.getUserDetails(userId);
      }
      throw error;
    }
  }

  /**
   * Check if beta endpoints are available
   *
   * Makes a test call to beta endpoint to verify availability.
   */
  async checkBetaAvailability(): Promise<boolean> {
    try {
      const client = this.getClient(true); // Force beta
      await client.api('/organization').get();
      return true;
    } catch (error: any) {
      return false;
    }
  }

  /**
   * Create multiple users using batch requests (up to 20 per batch)
   * More efficient than creating users one by one
   */
  async createUsersBatch(users: Array<{
    displayName: string;
    email: string;
    password?: string;
    employeeType?: string;
    companyName?: string;
    officeLocation?: string;
    usageLocation?: string;
  }>): Promise<{
    successful: User[];
    failed: Array<{ user: any; error: string }>;
  }> {
    const successful: User[] = [];
    const failed: Array<{ user: any; error: string }> = [];

    // Microsoft Graph batch limit is 20 requests per batch
    const BATCH_SIZE = 20;

    for (let i = 0; i < users.length; i += BATCH_SIZE) {
      const batch = users.slice(i, i + BATCH_SIZE);

      // Build batch request
      const batchRequests = batch.map((user, index) => {
        const mailNickname = user.email.split('@')[0];
        const password = user.password || this.generatePassword();

        const userParams: UserCreateParams = {
          displayName: user.displayName,
          userPrincipalName: user.email,
          mailNickname,
          accountEnabled: true,
          passwordProfile: {
            password,
            forceChangePasswordNextSignIn: false,
          },
        };

        // Add standard fields if provided
        if (user.usageLocation) userParams.usageLocation = user.usageLocation;

        // Add beta fields if provided
        if (user.employeeType) userParams.employeeType = user.employeeType;
        if (user.companyName) userParams.companyName = user.companyName;
        if (user.officeLocation) userParams.officeLocation = user.officeLocation;

        return {
          id: `${i + index}`,
          method: 'POST',
          url: '/users',
          body: userParams,
          headers: {
            'Content-Type': 'application/json',
          },
        };
      });

      try {
        // Use beta client if beta fields are present
        const hasBetaFields = batch.some(u => u.employeeType || u.companyName || u.officeLocation);
        const client = hasBetaFields ? this.getClient(true) : this.getClient(false);

        const batchResponse = await client.api('/$batch').post({
          requests: batchRequests,
        });

        // Process batch response
        for (let j = 0; j < batchResponse.responses.length; j++) {
          const response = batchResponse.responses[j];
          const user = batch[j];

          if (response.status >= 200 && response.status < 300) {
            successful.push(response.body);
            console.log(`✓ Created user: ${user.displayName} (${user.email})`);
          } else {
            const error = response.body?.error?.message || `HTTP ${response.status}`;
            failed.push({ user, error });
            console.error(`✗ Failed to create ${user.displayName}: ${error}`);
          }
        }

        // Small delay between batches to avoid rate limiting (40 requests per second limit)
        if (i + BATCH_SIZE < users.length) {
          await new Promise(resolve => setTimeout(resolve, 500));
        }

      } catch (error: any) {
        // If batch fails entirely, mark all users in batch as failed
        batch.forEach(user => {
          failed.push({
            user,
            error: error.message || String(error),
          });
          console.error(`✗ Batch failed for ${user.displayName}: ${error.message}`);
        });
      }
    }

    return { successful, failed };
  }

  /**
   * Update multiple users using batch requests (up to 20 per batch)
   * More efficient than updating users one by one
   */
  async updateUsersBatch(updates: Array<{
    userId: string;
    updates: Record<string, any>;
  }>): Promise<{
    successful: Array<{ userId: string; user: User }>;
    failed: Array<{ userId: string; error: string }>;
  }> {
    const successful: Array<{ userId: string; user: User }> = [];
    const failed: Array<{ userId: string; error: string }> = [];

    const BATCH_SIZE = 20;

    for (let i = 0; i < updates.length; i += BATCH_SIZE) {
      const batch = updates.slice(i, i + BATCH_SIZE);

      // Build batch request
      const batchRequests = batch.map((update, index) => ({
        id: `${i + index}`,
        method: 'PATCH',
        url: `/users/${update.userId}`,
        body: update.updates,
        headers: {
          'Content-Type': 'application/json',
        },
      }));

      try {
        // Use beta client if beta fields are present
        const hasBetaFields = batch.some(u =>
          Object.keys(u.updates).some(key =>
            ['employeeType', 'companyName', 'officeLocation'].includes(key)
          )
        );
        const client = hasBetaFields ? this.getClient(true) : this.getClient(false);

        const batchResponse = await client.api('/$batch').post({
          requests: batchRequests,
        });

        // Process batch response
        for (let j = 0; j < batchResponse.responses.length; j++) {
          const response = batchResponse.responses[j];
          const { userId } = batch[j];

          if (response.status >= 200 && response.status < 300) {
            successful.push({ userId, user: response.body });
            console.log(`✓ Updated user: ${userId}`);
          } else {
            const error = response.body?.error?.message || `HTTP ${response.status}`;
            failed.push({ userId, error });
            console.error(`✗ Failed to update ${userId}: ${error}`);
          }
        }

        // Small delay between batches to avoid rate limiting
        if (i + BATCH_SIZE < updates.length) {
          await new Promise(resolve => setTimeout(resolve, 500));
        }

      } catch (error: any) {
        // If batch fails entirely, mark all users in batch as failed
        batch.forEach(({ userId }) => {
          failed.push({
            userId,
            error: error.message || String(error),
          });
          console.error(`✗ Batch failed for ${userId}: ${error.message}`);
        });
      }
    }

    return { successful, failed };
  }

  /**
   * Delete multiple users using batch requests (up to 20 per batch)
   * More efficient than deleting users one by one
   */
  async deleteUsersBatch(userIds: string[]): Promise<{
    successful: string[];
    failed: Array<{ userId: string; error: string }>;
  }> {
    const successful: string[] = [];
    const failed: Array<{ userId: string; error: string }> = [];

    const BATCH_SIZE = 20;

    for (let i = 0; i < userIds.length; i += BATCH_SIZE) {
      const batch = userIds.slice(i, i + BATCH_SIZE);

      // Build batch request
      const batchRequests = batch.map((userId, index) => ({
        id: `${i + index}`,
        method: 'DELETE',
        url: `/users/${userId}`,
      }));

      try {
        const batchResponse = await this.client.api('/$batch').post({
          requests: batchRequests,
        });

        // Process batch response
        for (let j = 0; j < batchResponse.responses.length; j++) {
          const response = batchResponse.responses[j];
          const userId = batch[j];

          if (response.status >= 200 && response.status < 300) {
            successful.push(userId);
            console.log(`✓ Deleted user: ${userId}`);
          } else if (response.status === 404) {
            // User already deleted - treat as success
            successful.push(userId);
            console.log(`✓ User already deleted: ${userId}`);
          } else {
            const error = response.body?.error?.message || `HTTP ${response.status}`;
            failed.push({ userId, error });
            console.error(`✗ Failed to delete ${userId}: ${error}`);
          }
        }

        // Small delay between batches to avoid rate limiting
        if (i + BATCH_SIZE < userIds.length) {
          await new Promise(resolve => setTimeout(resolve, 500));
        }

      } catch (error: any) {
        // If batch fails entirely, mark all users in batch as failed
        batch.forEach(userId => {
          failed.push({
            userId,
            error: error.message || String(error),
          });
          console.error(`✗ Batch failed for ${userId}: ${error.message}`);
        });
      }
    }

    return { successful, failed };
  }

  /**
   * Assign manager to a user
   * Sets the organizational hierarchy by establishing manager relationship
   */
  async assignManager(userId: string, managerId: string): Promise<void> {
    try {
      await this.client.api(`/users/${userId}/manager/$ref`).put({
        '@odata.id': `https://graph.microsoft.com/v1.0/users/${managerId}`,
      });
      console.log(`✓ Assigned manager for user: ${userId}`);
    } catch (error: any) {
      if (error.statusCode === 404) {
        console.warn(`⚠ Manager or user not found: ${userId}`);
      } else {
        console.error(`✗ Failed to assign manager for ${userId}: ${error.message}`);
        throw error;
      }
    }
  }

  /**
   * Get user's manager
   */
  async getManager(userId: string): Promise<User | null> {
    try {
      const manager = await this.client.api(`/users/${userId}/manager`).get();
      return manager;
    } catch (error: any) {
      if (error.statusCode === 404) {
        return null; // No manager assigned
      }
      throw error;
    }
  }

  /**
   * Remove manager from user
   */
  async removeManager(userId: string): Promise<void> {
    try {
      await this.client.api(`/users/${userId}/manager/$ref`).delete();
      console.log(`✓ Removed manager from user: ${userId}`);
    } catch (error: any) {
      if (error.statusCode === 404) {
        return; // No manager to remove
      }
      throw error;
    }
  }

  /**
   * Batch assign managers to multiple users
   * More efficient for bulk manager assignments
   */
  async assignManagersBatch(assignments: Array<{
    userId: string;
    managerId: string;
  }>): Promise<{
    successful: string[];
    failed: Array<{ userId: string; error: string }>;
  }> {
    const successful: string[] = [];
    const failed: Array<{ userId: string; error: string }> = [];

    const BATCH_SIZE = 20;

    for (let i = 0; i < assignments.length; i += BATCH_SIZE) {
      const batch = assignments.slice(i, i + BATCH_SIZE);

      // Build batch request
      const batchRequests = batch.map((assignment, index) => ({
        id: `${i + index}`,
        method: 'PUT',
        url: `/users/${assignment.userId}/manager/$ref`,
        body: {
          '@odata.id': `https://graph.microsoft.com/v1.0/users/${assignment.managerId}`,
        },
        headers: {
          'Content-Type': 'application/json',
        },
      }));

      try {
        const batchResponse = await this.client.api('/$batch').post({
          requests: batchRequests,
        });

        // Process batch response
        for (let j = 0; j < batchResponse.responses.length; j++) {
          const response = batchResponse.responses[j];
          const { userId } = batch[j];

          if (response.status >= 200 && response.status < 300) {
            successful.push(userId);
            console.log(`✓ Assigned manager for user: ${userId}`);
          } else {
            const error = response.body?.error?.message || `HTTP ${response.status}`;
            failed.push({ userId, error });
            console.error(`✗ Failed to assign manager for ${userId}: ${error}`);
          }
        }

        // Small delay between batches
        if (i + BATCH_SIZE < assignments.length) {
          await new Promise(resolve => setTimeout(resolve, 500));
        }

      } catch (error: any) {
        // If batch fails entirely, mark all users in batch as failed
        batch.forEach(({ userId }) => {
          failed.push({
            userId,
            error: error.message || String(error),
          });
          console.error(`✗ Batch failed for ${userId}: ${error.message}`);
        });
      }
    }

    return { successful, failed };
  }

  /**
   * Get the Client instances for extension manager
   */
  getClients(): { client: Client; betaClient: Client } {
    return {
      client: this.client,
      betaClient: this.betaClient,
    };
  }

  /**
   * Generate a secure random password
   */
  private generatePassword(): string {
    const length = 16;
    const charset = 'abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789!@#$%^&*';
    let password = '';

    // Ensure password has at least one of each required character type
    password += 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'[Math.floor(Math.random() * 26)];
    password += 'abcdefghijklmnopqrstuvwxyz'[Math.floor(Math.random() * 26)];
    password += '0123456789'[Math.floor(Math.random() * 10)];
    password += '!@#$%^&*'[Math.floor(Math.random() * 8)];

    // Fill the rest randomly
    for (let i = password.length; i < length; i++) {
      password += charset[Math.floor(Math.random() * charset.length)];
    }

    // Shuffle the password
    return password.split('').sort(() => Math.random() - 0.5).join('');
  }
}

// CLI support for direct execution
if (import.meta.url === `file://${process.argv[1]}`) {
  const command = process.argv[2];

  // Check for required environment variables
  const tenantId = process.env.AZURE_TENANT_ID;
  const clientId = process.env.AZURE_CLIENT_ID;
  const clientSecret = process.env.AZURE_CLIENT_SECRET;

  if (!tenantId || !clientId) {
    console.error('❌ Error: AZURE_TENANT_ID and AZURE_CLIENT_ID must be set in .env file');
    console.error('   For MSAL authentication, these are the only required variables.');
    process.exit(1);
  }

  try {
    // Try client secret first (backward compatibility), then fall back to MSAL
    let client: GraphClient;

    if (clientSecret) {
      // Legacy: Use client secret
      console.log('Using client secret authentication (legacy)...\n');
      client = new GraphClient({ tenantId, clientId, clientSecret });
    } else {
      // Modern: Use browser-based MSAL
      const { BrowserAuthServer } = await import('./auth/browser-auth-server.js');
      const authServer = new BrowserAuthServer({ tenantId, clientId });
      console.log('Starting browser authentication...\n');
      const authResult = await authServer.authenticate();
      console.log(`✅ Authenticated as: ${authResult.account.username}\n`);
      client = new GraphClient({ accessToken: authResult.accessToken });
    }

    switch (command) {
      case 'test-connection':
        await client.testConnection();
        break;

      case 'list-users':
        const users = await client.listUsers();
        console.log(`\n✓ Found ${users.length} users:\n`);
        users.forEach(user => {
          console.log(`  ${user.displayName} (${user.userPrincipalName})`);
        });
        break;

      case 'list-licenses':
        const licenses = await client.listLicenses();
        console.log(`\n✓ Available licenses:\n`);
        licenses.forEach(license => {
          const available = license.prepaidUnits.enabled - license.consumedUnits;
          console.log(`  ${license.skuPartNumber}`);
          console.log(`    SKU ID: ${license.skuId}`);
          console.log(`    Available: ${available}/${license.prepaidUnits.enabled}\n`);
        });
        break;

      case 'check-mailboxes':
        const allUsers = await client.listUsers();
        console.log(`\nChecking mailboxes for ${allUsers.length} users...\n`);
        for (const user of allUsers) {
          const hasMailbox = await client.checkMailbox(user.id);
          const status = hasMailbox ? '✓' : '✗';
          console.log(`  ${status} ${user.displayName}`);
        }
        break;

      default:
        console.log('Usage: node dist/graph-client.js [command]');
        console.log('Commands: test-connection, list-users, list-licenses, check-mailboxes');
        console.log('\nNote: Requires AZURE_TENANT_ID and AZURE_CLIENT_ID in .env');
        console.log('      Authentication via MSAL Device Code Flow (no client secret needed)');
    }
  } catch (error: any) {
    console.error(`\n❌ Error: ${error.message}\n`);
    process.exit(1);
  }
}
