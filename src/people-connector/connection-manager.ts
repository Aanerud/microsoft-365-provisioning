import { Client } from '@microsoft/microsoft-graph-client';

export interface PeopleConnectionConfig {
  connectionId: string;
  name: string;
  description: string;
}

export class PeopleConnectionManager {
  private client: Client;
  private betaClient: Client;
  private connectionId: string;

  constructor(client: Client, betaClient: Client, connectionId: string) {
    this.client = client;
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

    await this.client.api('/external/connections').post(connectionRequest);
    console.log(`✓ Created connection: ${this.connectionId}`);
  }

  /**
   * Get connection status
   */
  async getConnection(): Promise<any> {
    return await this.client.api(`/external/connections/${this.connectionId}`).get();
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
      await this.client
        .api(`/external/connections/${this.connectionId}/schema`)
        .post(schema);

      console.log('✓ Schema registration initiated');

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
        return;
      }
      throw error;
    }
  }

  /**
   * Wait for schema to be in 'ready' state
   */
  private async waitForSchemaReady(maxWaitMs: number = 60000): Promise<void> {
    const startTime = Date.now();

    while (Date.now() - startTime < maxWaitMs) {
      const connection = await this.getConnection();

      if (connection.state === 'ready') {
        console.log('✓ Schema is ready');
        return;
      } else if (connection.state === 'failed') {
        throw new Error(`Schema registration failed: ${connection.failureReason}`);
      }

      console.log(`  Schema state: ${connection.state}, waiting...`);
      await new Promise(resolve => setTimeout(resolve, 3000));
    }

    throw new Error('Schema registration timeout');
  }

  /**
   * Delete connection (cleanup)
   */
  async deleteConnection(): Promise<void> {
    try {
      await this.client.api(`/external/connections/${this.connectionId}`).delete();
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
   */
  async registerAsProfileSource(displayName: string, webUrl: string): Promise<void> {
    console.log('📋 Registering connection as profile source (beta API)...');

    try {
      // Step 1: Register as a profile source (beta endpoint)
      const profileSourcePayload = {
        sourceId: this.connectionId,
        displayName,
        webUrl,
      };

      try {
        // Must use beta client - this endpoint is not available in v1.0
        await this.betaClient.api('/admin/people/profileSources').post(profileSourcePayload);
        console.log('✓ Registered as profile source');
      } catch (error: any) {
        // 409 Conflict means already registered
        if (error.statusCode === 409) {
          console.log('✓ Already registered as profile source');
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
          return;
        }

        const settingsId = settings.id;
        const currentSources: string[] = settings.prioritizedSourceUrls || [];

        // Check if already in prioritized list
        if (!currentSources.includes(sourceUrl)) {
          // Add to front of priority list (highest priority = index 0)
          const updatedSources = [sourceUrl, ...currentSources];

          // PATCH the specific settings item by ID
          await this.betaClient.api(`/admin/people/profilePropertySettings/${settingsId}`).patch({
            prioritizedSourceUrls: updatedSources,
          });
          console.log('✓ Added to prioritized profile sources (highest priority)');
        } else {
          console.log('✓ Already in prioritized profile sources');
        }
      } catch (error: any) {
        // Profile property settings might not exist yet in some tenants
        if (error.statusCode === 404) {
          console.log('⚠ Profile property settings not found - skipping prioritization');
        } else {
          throw error;
        }
      }
    } catch (error: any) {
      // Log detailed error info for debugging
      console.warn(`⚠ Profile source registration failed: ${error.message}`);
      if (error.statusCode) {
        console.warn(`  Status code: ${error.statusCode}`);
      }
      if (error.body?.error?.message) {
        console.warn(`  Details: ${error.body.error.message}`);
      }
      console.warn('  This may cause labels to show as "unknownFutureValue"');
      console.warn('  Ensure PeopleSettings.ReadWrite.All application permission is granted');
      console.warn('  Note: People data connectors may require tenant opt-in (preview feature)');
    }
  }
}
