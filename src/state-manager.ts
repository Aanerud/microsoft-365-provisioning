/**
 * State Manager
 *
 * Core state management system that transforms the tool from create-only to
 * a declarative state management system where CSV is the source of truth.
 *
 * Handles:
 * - CREATE: Users in CSV but not in Azure AD
 * - UPDATE: Users in both with differing attributes
 * - DELETE: Users in Azure AD but not in CSV
 * - State comparison and change detection for all 50+ Option A properties + custom columns (Option B)
 */

import { GraphClient } from './graph-client.js';
import { AccountProtectionService } from './safety/account-protection.js';
import {
  isStandardProperty,
  getPropertyMetadata,
  parsePropertyValue,
  getCustomProperties,
} from './schema/user-property-schema.js';

export interface UserState {
  email: string;
  displayName: string;
  [key: string]: any; // Only Option A standard properties
  // Note: Custom properties and Option B enrichment data are NOT included
}

export interface FieldChange {
  field: string;
  oldValue: any;
  newValue: any;
  isCustomProperty?: boolean;
}

export type ActionType = 'CREATE' | 'UPDATE' | 'DELETE' | 'NO_CHANGE';

export interface StateAction {
  action: ActionType;
  user: UserState;
  azureAdUser?: any; // Current Azure AD user (if exists)
  changes?: FieldChange[]; // Fields that changed (for UPDATE)
}

export interface StateDelta {
  create: StateAction[];
  update: StateAction[];
  delete: StateAction[];
  noChange: StateAction[];
  summary: {
    totalInCsv: number;
    totalInAzureAd: number;
    toCreate: number;
    toUpdate: number;
    toDelete: number;
    unchanged: number;
    customPropertiesDetected: string[];
  };
}

export interface StateManagerOptions {
  graphClient: GraphClient;
  protectionService?: AccountProtectionService;
}

export class StateManager {
  private graphClient: GraphClient;
  private protectionService: AccountProtectionService;

  constructor(options: StateManagerOptions) {
    this.graphClient = options.graphClient;
    this.protectionService = options.protectionService || AccountProtectionService.fromEnvironment(options.graphClient);
  }

  /**
   * Fetch current Azure AD state for all users (Option A standard properties only)
   */
  async fetchAzureAdState(): Promise<
    Map<string, { user: any }>
  > {
    console.log('🔍 Fetching current Azure AD state (Option A properties)...');

    // Get all users from Azure AD
    const users = await this.graphClient.listUsers();
    console.log(`  Found ${users.length} users in Azure AD`);

    // Build map keyed by email (userPrincipalName)
    const userMap = new Map<string, { user: any }>();
    for (const user of users) {
      userMap.set(user.userPrincipalName, { user });
    }

    return userMap;
  }

  /**
   * Calculate delta between CSV and Azure AD
   * Determines which users need to be created, updated, or deleted
   * Applies account protection filters to prevent modifying/deleting admin accounts
   */
  async calculateDelta(
    csvUsers: any[],
    csvColumns: string[],
    azureAdUsers: Map<string, { user: any }>
  ): Promise<StateDelta> {
    console.log('📊 Calculating state delta...');

    const create: StateAction[] = [];
    const update: StateAction[] = [];
    const deleteActions: StateAction[] = [];
    const noChange: StateAction[] = [];

    // Detect custom properties (columns not in standard schema, excluding internal columns)
    const customPropertiesDetected = getCustomProperties(csvColumns);
    if (customPropertiesDetected.length > 0) {
      console.log(
        `  Detected ${customPropertiesDetected.length} custom enrichment properties: ${customPropertiesDetected.join(', ')}`
      );
    }

    // Process CSV users
    for (const csvUser of csvUsers) {
      const email = csvUser.email;
      const azureAdState = azureAdUsers.get(email);

      if (!azureAdState) {
        // User in CSV but not in Azure AD -> CREATE
        create.push({
          action: 'CREATE',
          user: this.normalizeUserData(csvUser, csvColumns),
        });
      } else {
        // User exists in both -> check for changes (Option A properties only)
        const changes = this.detectChanges(
          csvUser,
          csvColumns,
          azureAdState.user
        );

        if (changes.length > 0) {
          // User has changes -> UPDATE
          update.push({
            action: 'UPDATE',
            user: this.normalizeUserData(csvUser, csvColumns),
            azureAdUser: azureAdState.user,
            changes,
          });
        } else {
          // User unchanged -> NO_CHANGE
          noChange.push({
            action: 'NO_CHANGE',
            user: this.normalizeUserData(csvUser, csvColumns),
            azureAdUser: azureAdState.user,
          });
        }
      }
    }

    // Find users in Azure AD but not in CSV -> DELETE
    const csvEmails = new Set(csvUsers.map(u => u.email));
    for (const [email, azureState] of azureAdUsers) {
      if (!csvEmails.has(email)) {
        deleteActions.push({
          action: 'DELETE',
          user: {
            email,
            displayName: azureState.user.displayName,
            userId: azureState.user.id,
          },
          azureAdUser: azureState.user,
        });
      }
    }

    // Apply account protection filters
    console.log('🛡️  Applying account protection filters...');

    // Filter UPDATE actions - remove protected accounts
    const updateAccountsToCheck = update.map(action => ({
      email: action.user.email,
      userId: action.azureAdUser?.id,
    }));

    const updateFilterResult = await this.protectionService.filterProtectedAccounts(updateAccountsToCheck);
    const protectedFromUpdate = updateFilterResult.protectedAccounts;

    // Remove protected accounts from update list and move to noChange
    const filteredUpdate = update.filter(action =>
      !protectedFromUpdate.some(p => p.email === action.user.email)
    );

    const movedToNoChange = update.filter(action =>
      protectedFromUpdate.some(p => p.email === action.user.email)
    );

    noChange.push(...movedToNoChange);

    // Filter DELETE actions - remove protected accounts
    const deleteAccountsToCheck = deleteActions.map(action => ({
      email: action.user.email,
      userId: action.azureAdUser?.id,
    }));

    const deleteFilterResult = await this.protectionService.filterProtectedAccounts(deleteAccountsToCheck);
    const protectedFromDelete = deleteFilterResult.protectedAccounts;

    // Remove protected accounts from delete list
    const filteredDelete = deleteActions.filter(action =>
      !protectedFromDelete.some(p => p.email === action.user.email)
    );

    // Display warnings if any accounts were protected
    if (protectedFromUpdate.length > 0) {
      this.protectionService.displayProtectionWarning(protectedFromUpdate, 'UPDATE');
    }

    if (protectedFromDelete.length > 0) {
      this.protectionService.displayProtectionWarning(protectedFromDelete, 'DELETE');
    }

    const delta: StateDelta = {
      create,
      update: filteredUpdate,
      delete: filteredDelete,
      noChange,
      summary: {
        totalInCsv: csvUsers.length,
        totalInAzureAd: azureAdUsers.size,
        toCreate: create.length,
        toUpdate: filteredUpdate.length,
        toDelete: filteredDelete.length,
        unchanged: noChange.length,
        customPropertiesDetected,
      },
    };

    console.log('  Delta calculation complete');
    return delta;
  }

  /**
   * Detect changes between CSV user and Azure AD user
   * Returns array of changed fields (Option A properties only)
   * Option B enrichment data and custom properties are NOT compared
   * (Custom properties flow through Graph Connectors, not Entra ID)
   */
  private detectChanges(
    csvUser: any,
    csvColumns: string[],
    azureUser: any
  ): FieldChange[] {
    const changes: FieldChange[] = [];

    // Only compare Option A (standard Entra ID) properties
    for (const column of csvColumns) {
      // Skip non-standard properties (custom properties flow through Graph Connectors)
      if (!isStandardProperty(column)) {
        continue;
      }

      const metadata = getPropertyMetadata(column);
      if (!metadata) {
        continue;
      }

      // Skip Option B properties (enrichment data handled by Graph Connectors)
      if (metadata.handledBy === 'optionB') {
        continue;
      }

      const csvValue = parsePropertyValue(column, csvUser[column]);
      const azureValue = azureUser[metadata.graphPath];

      // Skip undefined/empty values
      if (csvValue === undefined || csvValue === null || csvValue === '') {
        continue;
      }

      // Type-aware comparison
      if (!this.valuesEqual(csvValue, azureValue, metadata.type)) {
        changes.push({
          field: column,
          oldValue: azureValue,
          newValue: csvValue,
          isCustomProperty: false,
        });
      }
    }

    return changes;
  }

  /**
   * Type-aware value comparison
   */
  private valuesEqual(value1: any, value2: any, type: string): boolean {
    // Handle null/undefined
    if (value1 == null && value2 == null) {
      return true;
    }
    if (value1 == null || value2 == null) {
      return false;
    }

    switch (type) {
      case 'array':
        if (!Array.isArray(value1) || !Array.isArray(value2)) {
          return false;
        }
        if (value1.length !== value2.length) {
          return false;
        }
        const sorted1 = [...value1].sort();
        const sorted2 = [...value2].sort();
        return JSON.stringify(sorted1) === JSON.stringify(sorted2);

      case 'date':
        const date1 = new Date(value1).getTime();
        const date2 = new Date(value2).getTime();
        return date1 === date2;

      case 'object':
        return JSON.stringify(value1) === JSON.stringify(value2);

      case 'boolean':
        return Boolean(value1) === Boolean(value2);

      case 'number':
        return Number(value1) === Number(value2);

      default:
        // String comparison
        return String(value1) === String(value2);
    }
  }

  /**
   * Normalize user data from CSV
   * Parses ONLY Option A (standard Entra ID) properties
   * Option B enrichment data and custom properties flow through Graph Connectors
   */
  private normalizeUserData(csvUser: any, csvColumns: string[]): UserState {
    const userState: UserState = {
      email: csvUser.email,
      displayName: csvUser.name || csvUser.displayName,
    };

    // Extract ONLY Option A standard properties
    for (const column of csvColumns) {
      const metadata = getPropertyMetadata(column);

      // Skip Option B properties (enrichment data - handled by Graph Connectors)
      if (metadata && metadata.handledBy === 'optionB') {
        continue;
      }

      // Skip custom properties (flow through Graph Connectors, not Entra ID)
      if (!metadata) {
        continue;
      }

      // Handle ONLY Option A standard properties
      if (metadata.handledBy === 'optionA') {
        const value = parsePropertyValue(column, csvUser[column]);
        if (value !== undefined && value !== null && value !== '') {
          userState[column] = value;
        }
      }
    }

    return userState;
  }

  /**
   * Generate human-readable diff report
   */
  generateDiffReport(delta: StateDelta): string {
    const lines: string[] = [];

    lines.push('═'.repeat(80));
    lines.push('Provisioning State Changes');
    lines.push('═'.repeat(80));
    lines.push('');
    lines.push(`Generated: ${new Date().toISOString()}`);
    lines.push('');

    // Summary
    lines.push('Summary');
    lines.push('─'.repeat(80));
    lines.push(`  Total in CSV:            ${delta.summary.totalInCsv} users`);
    lines.push(`  Total in Azure AD:       ${delta.summary.totalInAzureAd} users`);
    lines.push(`  To CREATE:               ${delta.summary.toCreate} users`);
    lines.push(`  To UPDATE:               ${delta.summary.toUpdate} users`);
    lines.push(`  To DELETE:               ${delta.summary.toDelete} users`);
    lines.push(`  Unchanged:               ${delta.summary.unchanged} users`);

    if (delta.summary.customPropertiesDetected.length > 0) {
      lines.push(
        `  Custom properties:       ${delta.summary.customPropertiesDetected.length} (${delta.summary.customPropertiesDetected.join(', ')})`
      );
    }

    // CREATE actions
    if (delta.create.length > 0) {
      lines.push('');
      lines.push('Users to CREATE');
      lines.push('─'.repeat(80));
      delta.create.forEach((action, i) => {
        lines.push(`${i + 1}. ${action.user.displayName} (${action.user.email})`);

        // Show standard properties
        const standardProps = Object.entries(action.user).filter(
          ([key]) => key !== 'email' && key !== 'displayName' && key !== 'customProperties'
        );
        if (standardProps.length > 0) {
          standardProps.forEach(([key, value]) => {
            lines.push(`     ${key}: ${value}`);
          });
        }

        // Show custom properties
        if (action.user.customProperties && Object.keys(action.user.customProperties).length > 0) {
          lines.push('     Custom Properties:');
          Object.entries(action.user.customProperties).forEach(([key, value]) => {
            lines.push(`       ${key}: ${value} (custom)`);
          });
        }
        lines.push('');
      });
    }

    // UPDATE actions
    if (delta.update.length > 0) {
      lines.push('');
      lines.push('Users to UPDATE');
      lines.push('─'.repeat(80));
      delta.update.forEach((action, i) => {
        lines.push(`${i + 1}. ${action.user.displayName} (${action.user.email})`);
        if (action.changes) {
          action.changes.forEach(change => {
            const customTag = change.isCustomProperty ? ' (custom)' : '';
            lines.push(
              `     ${change.field}: "${change.oldValue}" → "${change.newValue}"${customTag}`
            );
          });
        }
        lines.push('');
      });
    }

    // DELETE actions
    if (delta.delete.length > 0) {
      lines.push('');
      lines.push('Users to DELETE');
      lines.push('─'.repeat(80));
      lines.push('⚠️  WARNING: These users exist in Azure AD but not in CSV');
      lines.push('');
      delta.delete.forEach((action, i) => {
        lines.push(`${i + 1}. ${action.user.displayName} (${action.user.email})`);
        lines.push('');
      });
    }

    lines.push('═'.repeat(80));

    return lines.join('\n');
  }

  /**
   * Generate compact summary for logging
   */
  generateSummary(delta: StateDelta): string {
    const parts = [];

    if (delta.summary.toCreate > 0) {
      parts.push(`CREATE: ${delta.summary.toCreate}`);
    }
    if (delta.summary.toUpdate > 0) {
      parts.push(`UPDATE: ${delta.summary.toUpdate}`);
    }
    if (delta.summary.toDelete > 0) {
      parts.push(`DELETE: ${delta.summary.toDelete}`);
    }
    if (delta.summary.unchanged > 0) {
      parts.push(`UNCHANGED: ${delta.summary.unchanged}`);
    }

    return parts.join(' | ');
  }
}
