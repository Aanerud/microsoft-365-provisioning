import { Client } from '@microsoft/microsoft-graph-client';

export interface PeopleConnectionConfig {
  connectionId: string;
  name: string;
  description: string;
}

export class PeopleConnectionManager {
  private client: Client;
  private connectionId: string;

  constructor(client: Client, _betaClient: Client, connectionId: string) {
    this.client = client;
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
      if (error.statusCode === 409) {
        console.log('✓ Schema already registered');
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
}
