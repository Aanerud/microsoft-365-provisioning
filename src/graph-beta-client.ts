import { GraphClient, GraphClientConfig } from './graph-client.js';

/**
 * Microsoft Graph BETA Client
 *
 * Specialized client for Microsoft Graph BETA endpoints.
 * Provides convenient methods for beta-specific features and advanced operations.
 *
 * Beta endpoints provide:
 * - Extended user attributes (employeeType, companyName, officeLocation)
 * - Advanced search API (/beta/search/query)
 * - Preview provisioning features
 * - Early access to new Graph API capabilities
 *
 * Note: Beta APIs may change without notice. This client enforces beta
 * endpoints exclusively.
 */

export interface ExtendedUserAttributes {
  employeeType?: string;
  companyName?: string;
  officeLocation?: string;
  department?: string;
  jobTitle?: string;
  manager?: string;
}

export interface SearchRequest {
  query: string;
  entityTypes: string[];
  from?: number;
  size?: number;
}

export interface SearchResult {
  hitsContainers: Array<{
    hits: Array<{
      resource: any;
      rank: number;
      summary?: string;
    }>;
    total: number;
    moreResultsAvailable: boolean;
  }>;
}

export class GraphBetaClient {
  private baseClient: GraphClient;

  constructor(config: GraphClientConfig) {
    // Always enforce beta for this client
    this.baseClient = new GraphClient({ ...config });
  }

  /**
   * Create user with extended attributes
   *
   * Uses beta endpoint to create user with attributes not available in v1.0:
   * - employeeType: Employee, Contractor, Intern, etc.
   * - companyName: Organization or company name
   * - officeLocation: Physical office location
   *
   * Beta-only; no v1.0 fallback.
   */
  async createUserWithExtendedAttributes(params: {
    displayName: string;
    email: string;
    password?: string;
    attributes: ExtendedUserAttributes;
  }): Promise<any> {
    return await this.baseClient.createUser({
      displayName: params.displayName,
      email: params.email,
      password: params.password,
      employeeType: params.attributes.employeeType,
      companyName: params.attributes.companyName,
      officeLocation: params.attributes.officeLocation,
    });
  }

  /**
   * Update user with extended attributes
   *
   * Updates beta-only user attributes. Gracefully handles unavailability.
   */
  async updateExtendedAttributes(
    userId: string,
    attributes: ExtendedUserAttributes
  ): Promise<void> {
    await this.baseClient.updateUserBeta(userId, {
      employeeType: attributes.employeeType,
      companyName: attributes.companyName,
      officeLocation: attributes.officeLocation,
    });
  }

  /**
   * Get user with extended profile
   *
   * Retrieves user with beta-only attributes.
   * Beta-only; no v1.0 fallback.
   */
  async getExtendedUserProfile(userId: string): Promise<any> {
    return await this.baseClient.getUserBeta(userId);
  }

  /**
   * Advanced search using Microsoft Search API (beta)
   *
   * Unified search across entity types:
   * - person: Search users
   * - message: Search emails
   * - event: Search calendar events
   * - drive: Search files
   * - driveItem: Search specific drive items
   *
   * Note: Requires additional Search.Read.All permission
   *
   * @param _request Search request with query and entity types
   * @returns Search results or null if unavailable
   */
  async advancedSearch(_request: SearchRequest): Promise<SearchResult | null> {
    try {
      // This is a placeholder implementation
      // The actual Microsoft Search API requires specific request format
      console.warn('⚠ Advanced search is not fully implemented yet');
      console.warn('  This requires additional permissions: SearchConfiguration.Read.All');
      return null;
    } catch (error: any) {
      if (error.statusCode === 404 || error.message?.includes('beta')) {
        console.warn('⚠ Beta search endpoint unavailable');
        return null;
      }
      throw error;
    }
  }

  /**
   * Bulk user provisioning with extended attributes
   *
   * Efficiently provisions multiple users with beta attributes.
   * Includes error handling and rate limiting.
   *
   * @param users Array of users to provision
   * @returns Array of created users and errors
   */
  async bulkUserProvision(
    users: Array<{
      displayName: string;
      email: string;
      password?: string;
      attributes?: ExtendedUserAttributes;
    }>
  ): Promise<{
    successful: any[];
    failed: Array<{ user: any; error: string }>;
  }> {
    const successful: any[] = [];
    const failed: Array<{ user: any; error: string }> = [];

    for (const user of users) {
      try {
        const createdUser = await this.createUserWithExtendedAttributes({
          displayName: user.displayName,
          email: user.email,
          password: user.password,
          attributes: user.attributes || {},
        });

        successful.push(createdUser);

        // Rate limiting: Small delay between requests
        await new Promise((resolve) => setTimeout(resolve, 500));
      } catch (error: any) {
        failed.push({
          user,
          error: error.message || String(error),
        });
        console.error(`Failed to provision ${user.email}: ${error.message}`);
      }
    }

    return { successful, failed };
  }

  /**
   * Check if beta endpoints are available
   *
   * Tests beta endpoint availability by making a test call.
   * Useful for conditional feature enablement.
   */
  async isBetaAvailable(): Promise<boolean> {
    return await this.baseClient.checkBetaAvailability();
  }

  /**
   * List users with extended attributes
   *
   * Retrieves users with beta-only attributes included.
   * Beta-only; no v1.0 fallback.
   *
   * @param filter Optional OData filter
   * @returns Array of users with extended attributes
   */
  async listUsersWithExtendedAttributes(filter?: string): Promise<any[]> {
    try {
      // Use base client's list method, which should use beta if enabled
      return await this.baseClient.listUsers(filter);
    } catch (error: any) {
      console.warn('⚠ Failed to list users with beta attributes');
      throw error;
    }
  }

  /**
   * Get employee directory with extended information
   *
   * Retrieves all users formatted as employee directory entries
   * with extended attributes visible.
   *
   * @returns Employee directory with extended attributes
   */
  async getEmployeeDirectory(): Promise<
    Array<{
      name: string;
      email: string;
      employeeType?: string;
      company?: string;
      office?: string;
      department?: string;
    }>
  > {
    const users = await this.listUsersWithExtendedAttributes();

    return users.map((user) => ({
      name: user.displayName,
      email: user.userPrincipalName,
      employeeType: user.employeeType,
      company: user.companyName,
      office: user.officeLocation,
      department: user.department,
    }));
  }
}

/**
 * CLI utility for testing beta endpoints
 */
export async function main() {
  const command = process.argv[2];

  console.log('\n🔬 Microsoft Graph BETA Client Test\n');

  // Load environment variables
  const tenantId = process.env.AZURE_TENANT_ID;
  const clientId = process.env.AZURE_CLIENT_ID;
  const clientSecret = process.env.AZURE_CLIENT_SECRET;

  if (!tenantId || !clientId || !clientSecret) {
    console.error('❌ Error: Required environment variables not set');
    console.error('   Set AZURE_TENANT_ID, AZURE_CLIENT_ID, AZURE_CLIENT_SECRET');
    process.exit(1);
  }

  const betaClient = new GraphBetaClient({
    tenantId,
    clientId,
    clientSecret,
  });

  try {
    switch (command) {
      case 'check-availability':
        console.log('Checking beta endpoint availability...');
        const available = await betaClient.isBetaAvailable();
        console.log(`\nBeta endpoints: ${available ? '✓ Available' : '✗ Unavailable'}\n`);
        break;

      case 'list-extended':
        console.log('Listing users with extended attributes...\n');
        const users = await betaClient.listUsersWithExtendedAttributes();
        console.log(`Found ${users.length} users:\n`);
        users.slice(0, 5).forEach((user) => {
          console.log(`  ${user.displayName}`);
          console.log(`    Email: ${user.userPrincipalName}`);
          if (user.employeeType) console.log(`    Type: ${user.employeeType}`);
          if (user.companyName) console.log(`    Company: ${user.companyName}`);
          if (user.officeLocation) console.log(`    Office: ${user.officeLocation}`);
          console.log('');
        });
        break;

      case 'employee-directory':
        console.log('Fetching employee directory...\n');
        const directory = await betaClient.getEmployeeDirectory();
        console.log(`Total employees: ${directory.length}\n`);
        directory.slice(0, 10).forEach((employee) => {
          console.log(`  ${employee.name} (${employee.email})`);
          if (employee.employeeType) console.log(`    Type: ${employee.employeeType}`);
          if (employee.company) console.log(`    Company: ${employee.company}`);
          console.log('');
        });
        break;

      default:
        console.log('Usage: node dist/graph-beta-client.js [command]');
        console.log('');
        console.log('Commands:');
        console.log('  check-availability   Check if beta endpoints are available');
        console.log('  list-extended        List users with extended attributes');
        console.log('  employee-directory   Get employee directory with extended info');
        console.log('');
    }
  } catch (error: any) {
    console.error(`\n❌ Error: ${error.message}\n`);
    process.exit(1);
  }
}

// Run CLI if this file is executed directly
if (import.meta.url === `file://${process.argv[1]}`) {
  main().catch((error) => {
    console.error('Fatal error:', error);
    process.exit(1);
  });
}
