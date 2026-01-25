import { Client } from '@microsoft/microsoft-graph-client';
import { getOptionBProperties, getCustomProperties } from '../schema/user-property-schema.js';
import { Logger } from '../utils/logger.js';
import fs from 'fs/promises';
import path from 'path';

// Retry configuration (based on cocogen best practices)
const MAX_RETRIES = 6;
const BASE_DELAY_MS = 1000;
const MAX_DELAY_MS = 30000;
const RETRYABLE_STATUS_CODES = new Set([429, 500, 502, 503, 504]);

export interface PeopleItem {
  email: string;
  [key: string]: any;
}

export class PeopleItemIngester {
  private betaClient: Client;
  private connectionId: string;
  private logger: Logger;

  constructor(_client: Client, betaClient: Client, connectionId: string, logger: Logger) {
    this.betaClient = betaClient;
    this.connectionId = connectionId;
    this.logger = logger;
  }

  /**
   * Convert CSV row to external item
   * Includes Option B standard properties + custom organization properties
   */
  createExternalItem(csvRow: any, csvColumns: string[]): any {
    const email = csvRow.email;
    const itemId = `person-${email.replace(/@/g, '-').replace(/\./g, '-')}`;
    const optionBProps = getOptionBProperties();
    const customProps = getCustomProperties(csvColumns);

    // Build content from all enrichment data
    const contentParts = [];
    if (csvRow.aboutMe) contentParts.push(csvRow.aboutMe);
    if (csvRow.skills) contentParts.push(`Skills: ${Array.isArray(csvRow.skills) ? csvRow.skills.join(', ') : csvRow.skills}`);
    if (csvRow.interests) contentParts.push(`Interests: ${Array.isArray(csvRow.interests) ? csvRow.interests.join(', ') : csvRow.interests}`);

    // Build properties
    const properties: any = {
      // REQUIRED: Link to Entra ID user
      accountInformation: JSON.stringify({
        userPrincipalName: email
      })
    };

    // Add Option B standard properties
    for (const prop of optionBProps) {
      const value = csvRow[prop.name];
      if (!value || value === '') continue;

      if (prop.type === 'array') {
        // Array properties need @odata.type annotation
        properties[`${prop.name}@odata.type`] = 'Collection(String)';

        // Parse array value if it's a string
        let arrayValue: any[];
        if (typeof value === 'string') {
          try {
            // Try parsing JSON array
            const parsed = JSON.parse(value.replace(/'/g, '"'));
            arrayValue = Array.isArray(parsed) ? parsed : [value];
          } catch {
            // Fallback to comma-separated
            arrayValue = value.split(',').map((v: string) => v.trim());
          }
        } else if (Array.isArray(value)) {
          arrayValue = value;
        } else {
          arrayValue = [value];
        }

        // Handle special complex types
        if (prop.peopleDataLabel === 'personLanguage') {
          // Languages with proficiency - format: {"displayName":"Norwegian","proficiency":"nativeOrBilingual"}
          // Supported input formats:
          //   - JSON: [{"language":"Norwegian","proficiency":"native"}]
          //   - Colon: ["Norwegian:native"]
          //   - Parenthetical: ["Italian (Native)", "English (Fluent)"]
          //   - Plain: ["Norwegian"]
          properties[prop.name] = arrayValue.map((item: any) => {
            // Map friendly proficiency names to Microsoft Graph values
            const mapProficiency = (prof: string): string => {
              const normalized = prof.toLowerCase().trim();
              if (normalized.includes('native') || normalized.includes('bilingual')) return 'nativeOrBilingual';
              if (normalized.includes('fluent') || normalized === 'full' || normalized === 'fullprofessional') return 'fullProfessional';
              if (normalized.includes('professional') || normalized.includes('working')) return 'professionalWorking';
              if (normalized.includes('conversational') || normalized.includes('intermediate') || normalized.includes('limited')) return 'limitedWorking';
              if (normalized.includes('basic') || normalized.includes('beginner') || normalized.includes('elementary')) return 'elementary';
              return 'professionalWorking'; // default
            };

            if (typeof item === 'object' && item.language) {
              // Already parsed object with language/proficiency
              return JSON.stringify({
                displayName: item.language,
                proficiency: mapProficiency(item.proficiency || 'professionalWorking')
              });
            } else if (typeof item === 'string') {
              // Check for parenthetical format: "Italian (Native)"
              const parenMatch = item.match(/^(.+?)\s*\(([^)]+)\)$/);
              if (parenMatch) {
                const [, lang, prof] = parenMatch;
                return JSON.stringify({
                  displayName: lang.trim(),
                  proficiency: mapProficiency(prof)
                });
              }
              // Check for colon format: "Norwegian:native"
              if (item.includes(':')) {
                const [lang, prof] = item.split(':').map((s: string) => s.trim());
                return JSON.stringify({
                  displayName: lang,
                  proficiency: mapProficiency(prof)
                });
              }
              // Simple language name without proficiency
              return JSON.stringify({
                displayName: String(item).trim(),
                proficiency: 'professionalWorking'
              });
            } else {
              return JSON.stringify({
                displayName: String(item),
                proficiency: 'professionalWorking'
              });
            }
          });
        } else if (prop.peopleDataLabel) {
          // Other labeled array properties use: {"displayName":"..."}
          properties[prop.name] = arrayValue.map((item: string) =>
            JSON.stringify({ displayName: item })
          );
        } else {
          // Custom properties without labels can be plain strings
          properties[prop.name] = arrayValue;
        }
      } else if (prop.peopleDataLabel === 'personNote') {
        // personAnnotation entity requires: {"detail":"..."}
        properties[prop.name] = JSON.stringify({ detail: value });
      } else if (prop.peopleDataLabel) {
        // Other single-value labels (e.g., personWebSite) use: {"displayName":"..."}
        properties[prop.name] = JSON.stringify({ displayName: value });
      } else {
        // Custom properties without labels (interests, responsibilities, schools)
        properties[prop.name] = value;
      }
    }

    // Add custom organization properties (VTeam, BenefitPlan, CostCenter, etc.)
    for (const customProp of customProps) {
      const value = csvRow[customProp];
      if (value && value !== '') {
        properties[customProp] = String(value); // Store as string

        // Add to content for searchability
        contentParts.push(`${customProp}: ${value}`);
      }
    }

    return {
      id: itemId,
      content: {
        value: contentParts.join('. '),
        type: 'text'
      },
      properties,
      acl: [
        {
          type: 'everyone',
          value: 'everyone',
          accessType: 'grant'
        }
      ]
    };
  }

  /**
   * Ingest single item (using beta endpoint for People Data)
   */
  async ingestItem(item: any): Promise<void> {
    await this.betaClient
      .api(`/external/connections/${this.connectionId}/items/${item.id}`)
      .put(item);

    this.logger.success(`Ingested: ${item.id}`);
  }

  /**
   * Batch ingest items with retry logic (based on cocogen best practices)
   * Note: Using individual requests instead of batch due to auth token issues with $batch
   */
  async batchIngestItems(items: any[]): Promise<{
    successful: string[];
    failed: Array<{ id: string; error: string }>;
  }> {
    const successful: string[] = [];
    const failed: Array<{ id: string; error: string }> = [];

    // Process items individually using beta endpoint for People Data
    for (const item of items) {
      try {
        await this.putItemWithRetry(item);
        successful.push(item.id);
        this.logger.success(`Ingested: ${item.id}`);

        // Small delay to avoid rate limiting
        await new Promise(resolve => setTimeout(resolve, 200));

      } catch (error: any) {
        const errorMsg = error.message || error.code || 'Unknown error';
        failed.push({ id: item.id, error: errorMsg });
        this.logger.error(`Failed: ${item.id}`, {
          status: error.statusCode,
          error: errorMsg,
          detail: error.body || ''
        });
      }
    }

    return { successful, failed };
  }

  /**
   * PUT item with exponential backoff retry (based on cocogen pattern)
   * Retries on 429, 500, 502, 503, 504 up to MAX_RETRIES times
   */
  private async putItemWithRetry(item: any): Promise<void> {
    let lastError: any;

    for (let attempt = 0; attempt <= MAX_RETRIES; attempt++) {
      try {
        await this.betaClient
          .api(`/external/connections/${this.connectionId}/items/${item.id}`)
          .put(item);
        return; // Success!

      } catch (error: any) {
        lastError = error;
        const statusCode = error.statusCode || error.code;

        // Check if this is a retryable error
        if (!this.shouldRetry(statusCode) || attempt === MAX_RETRIES) {
          throw error; // Non-retryable or max retries reached
        }

        // Parse Retry-After header if present
        const retryAfterMs = this.parseRetryAfter(error.headers);
        const delay = this.computeDelay(attempt, retryAfterMs);

        this.logger.warn(`Throttled (${statusCode}) for ${item.id}, retrying in ${delay}ms (attempt ${attempt + 1}/${MAX_RETRIES})`);
        await this.sleep(delay);
      }
    }

    throw lastError;
  }

  /**
   * Check if status code is retryable
   */
  private shouldRetry(statusCode: number | string): boolean {
    const code = typeof statusCode === 'string' ? parseInt(statusCode, 10) : statusCode;
    return RETRYABLE_STATUS_CODES.has(code);
  }

  /**
   * Parse Retry-After header value (seconds or HTTP date)
   */
  private parseRetryAfter(headers: any): number | null {
    if (!headers) return null;
    const value = headers['retry-after'] || headers['Retry-After'];
    if (!value) return null;

    const seconds = Number(value);
    if (!Number.isNaN(seconds) && Number.isFinite(seconds)) {
      return Math.max(0, seconds * 1000);
    }

    const dateMs = Date.parse(value);
    if (!Number.isNaN(dateMs)) {
      return Math.max(0, dateMs - Date.now());
    }

    return null;
  }

  /**
   * Compute delay with exponential backoff + jitter (cocogen pattern)
   */
  private computeDelay(attempt: number, retryAfter: number | null): number {
    if (retryAfter !== null) {
      return Math.min(retryAfter, MAX_DELAY_MS);
    }
    const exp = Math.min(MAX_DELAY_MS, BASE_DELAY_MS * Math.pow(2, attempt));
    const jitter = Math.floor(Math.random() * 250);
    return Math.min(MAX_DELAY_MS, exp + jitter);
  }

  /**
   * Sleep helper
   */
  private sleep(ms: number): Promise<void> {
    return new Promise(resolve => setTimeout(resolve, ms));
  }

  /**
   * Delete external items that are not in the current CSV (orphaned items)
   * Maintains a state file to track all created items
   */
  async deleteOrphanedItems(currentEmails: Set<string>): Promise<string[]> {
    const stateFilePath = path.join(process.cwd(), 'state', 'external-items-state.json');
    const deletedItems: string[] = [];

    try {
      // Ensure state directory exists
      await fs.mkdir(path.dirname(stateFilePath), { recursive: true });

      // Load previous state
      let previousItems: string[] = [];
      try {
        const stateData = await fs.readFile(stateFilePath, 'utf-8');
        const state = JSON.parse(stateData);
        previousItems = state.items || [];
      } catch (error) {
        // State file doesn't exist yet, first run
        this.logger.info('No previous state found, skipping orphan cleanup');
      }

      // Generate current item IDs from CSV emails
      const currentItemIds = Array.from(currentEmails).map(email =>
        `person-${email.replace(/@/g, '-').replace(/\./g, '-')}`
      );

      // Find orphaned items (in previous state but not in current CSV)
      const orphanedItems = previousItems.filter(itemId => !currentItemIds.includes(itemId));

      // Delete orphaned items
      if (orphanedItems.length > 0) {
        for (const itemId of orphanedItems) {
          try {
            await this.betaClient
              .api(`/external/connections/${this.connectionId}/items/${itemId}`)
              .delete();

            deletedItems.push(itemId);
            this.logger.success(`Deleted orphaned item: ${itemId}`);

            // Small delay to avoid rate limiting
            await new Promise(resolve => setTimeout(resolve, 100));

          } catch (error: any) {
            if (error.statusCode === 404) {
              // Item already deleted, that's fine
              this.logger.info(`Item already deleted: ${itemId}`);
              deletedItems.push(itemId);
            } else {
              this.logger.error(`Failed to delete orphaned item: ${itemId}`, {
                error: error.message
              });
            }
          }
        }
      }

      // Save current state
      await fs.writeFile(stateFilePath, JSON.stringify({
        items: currentItemIds,
        lastUpdated: new Date().toISOString()
      }, null, 2));

    } catch (error: any) {
      this.logger.error('Error during orphan cleanup', {
        error: error.message
      });
    }

    return deletedItems;
  }
}
