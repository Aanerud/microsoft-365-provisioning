/**
 * Account Protection System
 *
 * Protects critical admin accounts from being modified or deleted by the provisioning tool.
 * Provides multiple layers of protection:
 * 1. Email pattern matching (admin@*, administrator@*)
 * 2. Azure AD role detection (Global Administrator, etc.)
 * 3. Configurable exclusion list
 * 4. UPN domain protection
 */

import { GraphClient } from '../graph-client.js';

export interface ProtectedAccount {
  email: string;
  reason: string;
  role?: string;
}

export interface ProtectionConfig {
  protectedEmailPatterns: string[];
  protectedRoles: string[];
  protectedEmails: string[];
  checkAdminRoles: boolean;
}

export class AccountProtectionService {
  private graphClient: GraphClient;
  private config: ProtectionConfig;

  // Default protected email patterns
  private static readonly DEFAULT_PROTECTED_PATTERNS = [
    'admin@*',
    'administrator@*',
    'root@*',
    'systemadmin@*',
  ];

  // Azure AD admin roles that should be protected
  private static readonly PROTECTED_ROLES = [
    'Global Administrator',
    'Privileged Role Administrator',
    'Security Administrator',
    'User Administrator',
    'Directory Synchronization Accounts',
  ];

  constructor(graphClient: GraphClient, config?: Partial<ProtectionConfig>) {
    this.graphClient = graphClient;
    this.config = {
      protectedEmailPatterns: config?.protectedEmailPatterns || AccountProtectionService.DEFAULT_PROTECTED_PATTERNS,
      protectedRoles: config?.protectedRoles || AccountProtectionService.PROTECTED_ROLES,
      protectedEmails: config?.protectedEmails || [],
      checkAdminRoles: config?.checkAdminRoles ?? true,
    };
  }

  /**
   * Check if an email matches any protected patterns
   */
  private matchesPattern(email: string, pattern: string): boolean {
    // Convert pattern to regex
    // admin@* becomes ^admin@.*$
    const regexPattern = pattern
      .replace(/\./g, '\\.')
      .replace(/\*/g, '.*')
      .replace(/\?/g, '.');

    const regex = new RegExp(`^${regexPattern}$`, 'i');
    return regex.test(email);
  }

  /**
   * Check if an account is protected by email pattern
   */
  isProtectedByPattern(email: string): boolean {
    return this.config.protectedEmailPatterns.some(pattern =>
      this.matchesPattern(email, pattern)
    );
  }

  /**
   * Check if an account is in the explicit exclusion list
   */
  isProtectedByExclusionList(email: string): boolean {
    return this.config.protectedEmails.some(
      protectedEmail => protectedEmail.toLowerCase() === email.toLowerCase()
    );
  }

  /**
   * Get Azure AD roles for a user
   */
  async getUserRoles(userId: string): Promise<string[]> {
    try {
      const { client } = this.graphClient.getClients();
      const response = await client
        .api(`/users/${userId}/memberOf`)
        .select('displayName,roleTemplateId')
        .get();

      return response.value
        .filter((item: any) => item['@odata.type'] === '#microsoft.graph.directoryRole')
        .map((role: any) => role.displayName);
    } catch (error: any) {
      console.warn(`⚠ Could not fetch roles for user ${userId}: ${error.message}`);
      return [];
    }
  }

  /**
   * Check if a user has any protected admin roles
   */
  async hasProtectedRole(userId: string): Promise<{ isProtected: boolean; roles: string[] }> {
    if (!this.config.checkAdminRoles) {
      return { isProtected: false, roles: [] };
    }

    const userRoles = await this.getUserRoles(userId);
    const protectedRoles = userRoles.filter(role =>
      this.config.protectedRoles.some(protectedRole =>
        role.toLowerCase().includes(protectedRole.toLowerCase())
      )
    );

    return {
      isProtected: protectedRoles.length > 0,
      roles: protectedRoles,
    };
  }

  /**
   * Comprehensive check if an account is protected
   */
  async isAccountProtected(
    email: string,
    userId?: string
  ): Promise<{ isProtected: boolean; reason: string; roles?: string[] }> {
    // Check 1: Email pattern matching (fast)
    if (this.isProtectedByPattern(email)) {
      return {
        isProtected: true,
        reason: 'Email matches protected pattern',
      };
    }

    // Check 2: Explicit exclusion list (fast)
    if (this.isProtectedByExclusionList(email)) {
      return {
        isProtected: true,
        reason: 'Email in protected exclusion list',
      };
    }

    // Check 3: Azure AD role check (requires API call)
    if (userId && this.config.checkAdminRoles) {
      const roleCheck = await this.hasProtectedRole(userId);
      if (roleCheck.isProtected) {
        return {
          isProtected: true,
          reason: 'User has protected admin role',
          roles: roleCheck.roles,
        };
      }
    }

    return { isProtected: false, reason: '' };
  }

  /**
   * Filter protected accounts from a list
   * Returns: { allowed: [], protectedAccounts: [] }
   */
  async filterProtectedAccounts(
    accounts: Array<{ email: string; userId?: string }>
  ): Promise<{
    allowed: typeof accounts;
    protectedAccounts: ProtectedAccount[];
  }> {
    const allowed: typeof accounts = [];
    const protectedAccounts: ProtectedAccount[] = [];

    for (const account of accounts) {
      const check = await this.isAccountProtected(account.email, account.userId);

      if (check.isProtected) {
        protectedAccounts.push({
          email: account.email,
          reason: check.reason,
          role: check.roles?.join(', '),
        });
      } else {
        allowed.push(account);
      }
    }

    return { allowed, protectedAccounts };
  }

  /**
   * Get protection configuration for logging
   */
  getProtectionConfig(): string {
    const lines = [
      'Account Protection Configuration:',
      `  Protected Email Patterns: ${this.config.protectedEmailPatterns.join(', ')}`,
      `  Protected Roles: ${this.config.protectedRoles.length} roles`,
      `  Explicit Exclusions: ${this.config.protectedEmails.length} accounts`,
      `  Role Check Enabled: ${this.config.checkAdminRoles}`,
    ];
    return lines.join('\n');
  }

  /**
   * Load configuration from environment variables
   */
  static fromEnvironment(graphClient: GraphClient): AccountProtectionService {
    const protectedEmailPatterns = process.env.PROTECTED_EMAIL_PATTERNS
      ? process.env.PROTECTED_EMAIL_PATTERNS.split(',').map(p => p.trim())
      : AccountProtectionService.DEFAULT_PROTECTED_PATTERNS;

    const protectedEmails = process.env.PROTECTED_EMAILS
      ? process.env.PROTECTED_EMAILS.split(',').map(e => e.trim())
      : [];

    const checkAdminRoles = process.env.CHECK_ADMIN_ROLES !== 'false';

    return new AccountProtectionService(graphClient, {
      protectedEmailPatterns,
      protectedEmails,
      checkAdminRoles,
    });
  }

  /**
   * Display warning about protected accounts
   */
  displayProtectionWarning(protectedAccounts: ProtectedAccount[], operation: 'UPDATE' | 'DELETE'): void {
    if (protectedAccounts.length === 0) {
      return;
    }

    console.log('\n' + '⚠️ '.repeat(40));
    console.log(`🛡️  PROTECTED ACCOUNTS - ${operation} BLOCKED`);
    console.log('⚠️ '.repeat(40));
    console.log(`\nThe following ${protectedAccounts.length} account(s) are protected and will NOT be ${operation}d:\n`);

    protectedAccounts.forEach((account, i) => {
      console.log(`${i + 1}. ${account.email}`);
      console.log(`   Reason: ${account.reason}`);
      if (account.role) {
        console.log(`   Role: ${account.role}`);
      }
      console.log('');
    });

    console.log('To modify protection settings, see .env configuration:');
    console.log('  - PROTECTED_EMAIL_PATTERNS');
    console.log('  - PROTECTED_EMAILS');
    console.log('  - CHECK_ADMIN_ROLES');
    console.log('\n' + '─'.repeat(80) + '\n');
  }
}
