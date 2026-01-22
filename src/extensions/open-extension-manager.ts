/**
 * Open Extension Manager
 *
 * Manages custom user properties using Microsoft Graph Open Extensions.
 * Open extensions allow storing arbitrary JSON data on user objects without
 * modifying the schema.
 *
 * Documentation: https://learn.microsoft.com/en-us/graph/extensibility-open-users
 */

import { Client } from '@microsoft/microsoft-graph-client';
import { getCustomProperties } from '../schema/user-property-schema.js';

export interface OpenExtension {
  '@odata.type': string;
  extensionName: string;
  id?: string;
  [key: string]: any;
}

export class OpenExtensionManager {
  private betaClient: Client;
  private extensionName = 'com.m365provision.customFields';

  constructor(_client: Client, betaClient: Client) {
    this.betaClient = betaClient;
  }

  /**
   * Create open extension with custom properties for a user
   */
  async createExtension(
    userId: string,
    customProperties: Record<string, any>
  ): Promise<void> {
    if (Object.keys(customProperties).length === 0) {
      return; // Nothing to create
    }

    try {
      await this.betaClient.api(`/users/${userId}/extensions`).post({
        '@odata.type': 'microsoft.graph.openTypeExtension',
        extensionName: this.extensionName,
        ...customProperties,
      });
      console.log(`✓ Created extension with ${Object.keys(customProperties).length} custom properties for user ${userId}`);
    } catch (error: any) {
      // If extension already exists, try updating instead
      if (error.statusCode === 409 || error.code === 'Request_ResourceAlreadyExists') {
        console.warn(`⚠ Extension already exists for user ${userId}, updating instead`);
        await this.updateExtension(userId, customProperties);
      } else {
        console.error(`✗ Failed to create extension for user ${userId}: ${error.message}`);
        throw error;
      }
    }
  }

  /**
   * Update existing open extension
   */
  async updateExtension(
    userId: string,
    customProperties: Record<string, any>
  ): Promise<void> {
    try {
      await this.betaClient.api(`/users/${userId}/extensions/${this.extensionName}`).patch({
        ...customProperties,
      });
      console.log(`✓ Updated extension for user ${userId}`);
    } catch (error: any) {
      if (error.statusCode === 404) {
        // Extension doesn't exist, create it
        console.warn(`⚠ Extension not found for user ${userId}, creating instead`);
        await this.createExtension(userId, customProperties);
      } else {
        console.error(`✗ Failed to update extension for user ${userId}: ${error.message}`);
        throw error;
      }
    }
  }

  /**
   * Get extension properties for a user
   */
  async getExtension(userId: string): Promise<Record<string, any> | null> {
    try {
      const result = await this.betaClient
        .api(`/users/${userId}/extensions/${this.extensionName}`)
        .get();

      // Remove metadata fields
      const { '@odata.type': _odataType, extensionName, id, ...customProps } = result;
      return customProps;
    } catch (error: any) {
      if (error.statusCode === 404) {
        return null; // Extension doesn't exist
      }
      console.error(`✗ Failed to get extension for user ${userId}: ${error.message}`);
      throw error;
    }
  }

  /**
   * Delete extension from user
   */
  async deleteExtension(userId: string): Promise<void> {
    try {
      await this.betaClient
        .api(`/users/${userId}/extensions/${this.extensionName}`)
        .delete();
      console.log(`✓ Deleted extension for user ${userId}`);
    } catch (error: any) {
      if (error.statusCode === 404) {
        return; // Already deleted or never existed
      }
      console.error(`✗ Failed to delete extension for user ${userId}: ${error.message}`);
      throw error;
    }
  }

  /**
   * Check if user has custom properties extension
   */
  async hasExtension(userId: string): Promise<boolean> {
    const extension = await this.getExtension(userId);
    return extension !== null;
  }

  /**
   * Extract custom properties from CSV row
   * Returns only properties that are not in the standard schema
   */
  extractCustomProperties(
    csvRow: Record<string, any>,
    csvColumns: string[]
  ): Record<string, any> {
    const customProps: Record<string, any> = {};
    const customColumns = getCustomProperties(csvColumns);

    for (const col of customColumns) {
      const value = csvRow[col];
      if (value !== undefined && value !== '' && value !== null) {
        customProps[col] = value;
      }
    }

    return customProps;
  }

  /**
   * Batch create extensions for multiple users
   * More efficient than creating extensions one by one
   */
  async createExtensionsBatch(
    userExtensions: Array<{ userId: string; customProperties: Record<string, any> }>
  ): Promise<{
    successful: string[];
    failed: Array<{ userId: string; error: string }>;
  }> {
    const successful: string[] = [];
    const failed: Array<{ userId: string; error: string }> = [];

    // Filter out users with no custom properties
    const usersToProcess = userExtensions.filter(
      ue => Object.keys(ue.customProperties).length > 0
    );

    if (usersToProcess.length === 0) {
      return { successful, failed };
    }

    // Microsoft Graph batch limit is 20 requests per batch
    const BATCH_SIZE = 20;

    for (let i = 0; i < usersToProcess.length; i += BATCH_SIZE) {
      const batch = usersToProcess.slice(i, i + BATCH_SIZE);

      // Build batch request
      const batchRequests = batch.map((ue, index) => ({
        id: `${i + index}`,
        method: 'POST',
        url: `/users/${ue.userId}/extensions`,
        body: {
          '@odata.type': 'microsoft.graph.openTypeExtension',
          extensionName: this.extensionName,
          ...ue.customProperties,
        },
        headers: {
          'Content-Type': 'application/json',
        },
      }));

      try {
        const batchResponse = await this.betaClient.api('/$batch').post({
          requests: batchRequests,
        });

        // Process batch response
        for (let j = 0; j < batchResponse.responses.length; j++) {
          const response = batchResponse.responses[j];
          const { userId } = batch[j];

          if (response.status >= 200 && response.status < 300) {
            successful.push(userId);
          } else {
            const error = response.body?.error?.message || `HTTP ${response.status}`;
            failed.push({ userId, error });
          }
        }

        // Small delay between batches to avoid rate limiting
        if (i + BATCH_SIZE < usersToProcess.length) {
          await new Promise(resolve => setTimeout(resolve, 500));
        }
      } catch (error: any) {
        // If batch fails entirely, mark all users in batch as failed
        batch.forEach(({ userId }) => {
          failed.push({
            userId,
            error: error.message || String(error),
          });
        });
      }
    }

    return { successful, failed };
  }

  /**
   * Batch update extensions for multiple users
   */
  async updateExtensionsBatch(
    userExtensions: Array<{ userId: string; customProperties: Record<string, any> }>
  ): Promise<{
    successful: string[];
    failed: Array<{ userId: string; error: string }>;
  }> {
    const successful: string[] = [];
    const failed: Array<{ userId: string; error: string }> = [];

    const BATCH_SIZE = 20;

    for (let i = 0; i < userExtensions.length; i += BATCH_SIZE) {
      const batch = userExtensions.slice(i, i + BATCH_SIZE);

      const batchRequests = batch.map((ue, index) => ({
        id: `${i + index}`,
        method: 'PATCH',
        url: `/users/${ue.userId}/extensions/${this.extensionName}`,
        body: ue.customProperties,
        headers: {
          'Content-Type': 'application/json',
        },
      }));

      try {
        const batchResponse = await this.betaClient.api('/$batch').post({
          requests: batchRequests,
        });

        for (let j = 0; j < batchResponse.responses.length; j++) {
          const response = batchResponse.responses[j];
          const { userId } = batch[j];

          if (response.status >= 200 && response.status < 300) {
            successful.push(userId);
          } else {
            const error = response.body?.error?.message || `HTTP ${response.status}`;
            failed.push({ userId, error });
          }
        }

        if (i + BATCH_SIZE < userExtensions.length) {
          await new Promise(resolve => setTimeout(resolve, 500));
        }
      } catch (error: any) {
        batch.forEach(({ userId }) => {
          failed.push({
            userId,
            error: error.message || String(error),
          });
        });
      }
    }

    return { successful, failed };
  }

  /**
   * Get extensions for multiple users in batch
   */
  async getExtensionsBatch(userIds: string[]): Promise<Map<string, Record<string, any> | null>> {
    const results = new Map<string, Record<string, any> | null>();

    const BATCH_SIZE = 20;

    for (let i = 0; i < userIds.length; i += BATCH_SIZE) {
      const batch = userIds.slice(i, i + BATCH_SIZE);

      const batchRequests = batch.map((userId, index) => ({
        id: `${i + index}`,
        method: 'GET',
        url: `/users/${userId}/extensions/${this.extensionName}`,
      }));

      try {
        const batchResponse = await this.betaClient.api('/$batch').post({
          requests: batchRequests,
        });

        for (let j = 0; j < batchResponse.responses.length; j++) {
          const response = batchResponse.responses[j];
          const userId = batch[j];

          if (response.status === 200) {
            const { '@odata.type': _odataType, extensionName, id, ...customProps } = response.body;
            results.set(userId, customProps);
          } else if (response.status === 404) {
            results.set(userId, null); // No extension
          } else {
            results.set(userId, null); // Error fetching
          }
        }

        if (i + BATCH_SIZE < userIds.length) {
          await new Promise(resolve => setTimeout(resolve, 500));
        }
      } catch (error: any) {
        // Mark all users in failed batch as having no extension
        batch.forEach(userId => {
          results.set(userId, null);
        });
      }
    }

    return results;
  }
}
