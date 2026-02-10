import { Client } from '@microsoft/microsoft-graph-client';

const CONNECTION_ID_REGEX = /^[A-Za-z0-9]+$/;

function assertValidConnectionId(connectionId: string): void {
  if (!CONNECTION_ID_REGEX.test(connectionId)) {
    throw new Error(
      `Invalid connectionId "${connectionId}". People connector connection IDs must be alphanumeric only.`
    );
  }
}

export interface PeopleConnectionConfig {
  connectionId: string;
  name: string;
  description: string;
}

export class PeopleConnectionManager {
  private betaClient: Client;
  private connectionId: string;

  constructor(betaClient: Client, connectionId: string) {
    assertValidConnectionId(connectionId);
    this.betaClient = betaClient;
    this.connectionId = connectionId;
  }

  /**
   * Create the Graph Connector connection
   */
  async createConnection(name: string, description: string): Promise<void> {
    try {
      await this.getConnection();
      console.log(`✓ Connection already exists: ${this.connectionId}`);
      return;
    } catch (error: any) {
      if (error.statusCode !== 404) throw error;
    }

    const connectionRequest = {
      id: this.connectionId,
      name,
      description,
      // REQUIRED for people data connectors
      contentCategory: 'people',
      activitySettings: {
        urlToItemResolvers: []
      }
    };

    // Use beta API for connection creation - required for contentCategory: 'people'
    await this.betaClient.api('/external/connections').post(connectionRequest);
    console.log(`✓ Created connection: ${this.connectionId} (with contentCategory: people)`);
  }

  /**
   * Get connection status
   */
  async getConnection(): Promise<any> {
    return await this.betaClient.api(`/external/connections/${this.connectionId}`).get();
  }

  /**
   * Get schema status
   */
  async getSchema(): Promise<any> {
    return await this.betaClient
      .api(`/external/connections/${this.connectionId}/schema`)
      .header('Prefer', 'include-unknown-enum-members')
      .get();
  }

  /**
   * Register people data schema
   */
  async registerSchema(properties: any[]): Promise<void> {
    const schema = {
      baseType: "microsoft.graph.externalItem",
      properties
    };

    try {
      // Use beta API for schema registration - required for people data labels like 'personAccount'
      // The v1.0 API may treat these labels as 'unknownFutureValue' and fail user mapping
      await this.betaClient
        .api(`/external/connections/${this.connectionId}/schema`)
        .header('Prefer', 'include-unknown-enum-members')
        .patch(schema);

      console.log('✓ Schema registration initiated (using beta API)');

      // Wait for schema to be ready
      await this.waitForSchemaReady();
    } catch (error: any) {
      // 409 Conflict = schema already registered
      if (error.statusCode === 409) {
        console.log('✓ Schema already registered');
        return;
      }
      // 400 with "UpdateNotAllowed" = schema already exists and can't be updated
      if (error.statusCode === 400 && error.body?.includes('UpdateNotAllowed')) {
        console.log('✓ Schema already registered (cannot be updated)');
        console.log('  Delete and recreate the connector to apply schema changes.');
        return;
      }
      throw error;
    }
  }

  /**
   * Wait for schema to be in 'ready' state
   */
  private async waitForSchemaReady(maxWaitMs: number = 300000): Promise<void> {
    const startTime = Date.now();

    while (Date.now() - startTime < maxWaitMs) {
      const schemaStatus = await this.getSchema();
      let state = schemaStatus?.state ?? schemaStatus?.status;
      let failureReason = schemaStatus?.failureReason;

      if (!state) {
        const connection = await this.getConnection();
        state = connection.state;
        failureReason = connection.failureReason;
      }

      if (state === 'ready') {
        console.log('✓ Schema is ready');
        return;
      } else if (state === 'failed') {
        throw new Error(`Schema registration failed: ${failureReason || 'Unknown reason'}`);
      }

      console.log(`  Schema state: ${state}, waiting...`);
      await new Promise(resolve => setTimeout(resolve, 3000));
    }

    throw new Error('Schema registration timeout');
  }

  /**
   * Verify that the registered schema has the expected people data labels.
   * Throws if labels are missing or returned as 'unknownFutureValue'.
   */
  async verifySchemaLabels(expectedLabels: string[] = ['personAccount', 'personSkills']): Promise<void> {
    console.log('Verifying schema labels...');
    const schema = await this.getSchema();
    const properties: any[] = schema.properties || [];

    const foundLabels = new Set<string>();
    const problems: string[] = [];

    for (const prop of properties) {
      for (const label of prop.labels || []) {
        if (label === 'unknownFutureValue') {
          problems.push(`Property "${prop.name}" has label "unknownFutureValue" (Prefer header may be missing or beta API not used)`);
        } else {
          foundLabels.add(label);
        }
      }
    }

    for (const expected of expectedLabels) {
      if (!foundLabels.has(expected)) {
        problems.push(`Expected label "${expected}" not found in schema`);
      }
    }

    if (problems.length > 0) {
      console.error('Schema verification failed:');
      for (const p of problems) {
        console.error(`  - ${p}`);
      }
      throw new Error(`Schema verification failed: ${problems.join('; ')}`);
    }

    console.log(`✓ Schema labels verified: ${[...foundLabels].join(', ')}`);
  }

  /**
   * Delete connection (cleanup)
   */
  async deleteConnection(): Promise<void> {
    try {
      await this.betaClient.api(`/external/connections/${this.connectionId}`).delete();
      console.log(`✓ Deleted connection: ${this.connectionId}`);
    } catch (error: any) {
      if (error.statusCode === 404) {
        console.log('Connection already deleted');
        return;
      }
      throw error;
    }
  }

  /**
   * Register the Graph Connector as a profile source
   * This fixes the "unknownFutureValue" label issue by:
   * 1. Registering the connection as a profile source
   * 2. Adding it to the prioritized sources list
   *
   * IMPORTANT: Uses beta client as /admin/people/profileSources is a beta-only endpoint
   * Requires PeopleSettings.ReadWrite.All application permission
   *
   * @returns true if registration succeeded, false if it failed
   */
  async registerAsProfileSource(displayName: string, webUrl: string): Promise<boolean> {
    console.log('📋 Registering connection as profile source (beta API)...');

    let registrationSucceeded = false;

    try {
      // Step 1: Register as a profile source (beta endpoint)
      // IMPORTANT: 'kind' property is REQUIRED per Microsoft docs
      // https://learn.microsoft.com/graph/api/peopleadminsettings-post-profilesources
      const profileSourcePayload = {
        sourceId: this.connectionId,
        displayName,
        kind: 'Connector',
        webUrl,
      };

      try {
        // Must use beta client - this endpoint is not available in v1.0
        await this.betaClient.api('/admin/people/profileSources').post(profileSourcePayload);
        console.log('✓ Registered as profile source');
        registrationSucceeded = true;
      } catch (error: any) {
        // 409 Conflict means already registered
        if (error.statusCode === 409) {
          console.log('✓ Already registered as profile source');
          registrationSucceeded = true;
        } else if (error.statusCode === 403) {
          console.warn('⚠ Permission denied (403) for profile source registration');
          console.warn('  Missing PeopleSettings.ReadWrite.All application permission');
          console.warn('  To fix: Azure Portal > App Registrations > Your App > API Permissions');
          console.warn('          Add Microsoft Graph > Application > PeopleSettings.ReadWrite.All');
          console.warn('          Then grant admin consent');
          throw error;
        } else if (error.statusCode === 400) {
          console.warn('⚠ Bad request (400) for profile source registration');
          console.warn(`  Details: ${error.body?.error?.message || error.message}`);
          console.warn('  This may be a tenant configuration issue');
          throw error;
        } else {
          throw error;
        }
      }

      // Step 2: Add to prioritized sources in profile property settings
      // This ensures our connector data appears in user profiles
      // Reference: https://learn.microsoft.com/en-us/graph/profilepriority-configure-profilepropertysetting
      const sourceUrl = `https://graph.microsoft.com/beta/admin/people/profileSources(sourceId='${this.connectionId}')`;

      try {
        // Get current settings (returns a collection with 'value' array)
        const response = await this.betaClient.api('/admin/people/profilePropertySettings').get();
        const settings = response.value?.[0];

        if (!settings) {
          console.log('⚠ No profile property settings found - skipping prioritization');
          console.log('  This is normal for new tenants; prioritization will be set up automatically');
          // Not a failure, just not applicable
          return registrationSucceeded;
        }

        const settingsId = settings.id;
        const currentSources: string[] = settings.prioritizedSourceUrls || [];

        // Check if already in prioritized list (match by connection ID, not exact URL)
        const alreadyPrioritized = currentSources.some(url => url.includes(this.connectionId));

        if (!alreadyPrioritized) {
          // Add to front of priority list (highest priority = index 0)
          const updatedSources = [sourceUrl, ...currentSources];

          // PATCH the specific settings item by ID
          await this.betaClient.api(`/admin/people/profilePropertySettings/${settingsId}`).patch({
            prioritizedSourceUrls: updatedSources,
          });
          console.log('✓ Added to prioritized profile sources (highest priority)');
          // Prioritization succeeded
        } else {
          console.log('✓ Already in prioritized profile sources');
          // Already prioritized
        }
      } catch (error: any) {
        // Profile property settings might not exist yet in some tenants
        if (error.statusCode === 404) {
          console.log('⚠ Profile property settings not found - skipping prioritization');
          // Not a failure, just not applicable
        } else if (error.statusCode === 403) {
          console.warn('⚠ Permission denied (403) for profile property settings');
          // Don't throw - registration succeeded, prioritization is optional
        } else {
          console.warn(`⚠ Failed to update prioritization: ${error.message}`);
          // Don't throw - registration succeeded, prioritization is optional
        }
      }

      return registrationSucceeded;
    } catch (error: any) {
      // Log detailed error info for debugging
      console.warn(`⚠ Profile source registration failed: ${error.message}`);
      if (error.statusCode) {
        console.warn(`  Status code: ${error.statusCode}`);
      }
      if (error.body?.error?.message) {
        console.warn(`  Details: ${error.body.error.message}`);
      }
      console.warn('');
      console.warn('  To diagnose: node tools/debug/check-profile-source.mjs');
      console.warn('  To fix manually: node tools/admin/register-profile-source.mjs');
      console.warn('');
      console.warn('  Common issues:');
      console.warn('  - Missing PeopleSettings.ReadWrite.All application permission');
      console.warn('  - People data connectors may require tenant opt-in (preview feature)');
      console.warn('  - Connection must have contentCategory: "people"');

      return false;
    }
  }
}
