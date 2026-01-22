import fs from 'fs';
import path from 'path';
import os from 'os';

export interface CachedToken {
  accessToken: string;
  expiresOn: number; // Unix timestamp
  account: {
    username: string;
    name: string;
    tenantId: string;
  };
}

export class TokenCache {
  private cacheDir: string;
  private cacheFile: string;

  constructor() {
    this.cacheDir = path.join(os.homedir(), '.m365-provision');
    this.cacheFile = path.join(this.cacheDir, 'token-cache.json');
    this.ensureCacheDir();
  }

  private ensureCacheDir(): void {
    if (!fs.existsSync(this.cacheDir)) {
      fs.mkdirSync(this.cacheDir, { recursive: true, mode: 0o700 });
    }
  }

  /**
   * Save token to cache
   */
  save(token: CachedToken): void {
    try {
      fs.writeFileSync(this.cacheFile, JSON.stringify(token, null, 2), {
        mode: 0o600,
      });
      console.log('✓ Token cached for future use');
    } catch (error: any) {
      console.warn(`⚠ Failed to cache token: ${error.message}`);
    }
  }

  /**
   * Load token from cache
   */
  load(): CachedToken | null {
    try {
      if (!fs.existsSync(this.cacheFile)) {
        return null;
      }

      const content = fs.readFileSync(this.cacheFile, 'utf8');
      const token = JSON.parse(content) as CachedToken;

      // Check if token is still valid (with 5 minute buffer)
      const now = Date.now();
      const bufferMs = 5 * 60 * 1000; // 5 minutes

      if (token.expiresOn * 1000 < now + bufferMs) {
        console.log('⚠ Cached token expired');
        return null;
      }

      return token;
    } catch (error: any) {
      console.warn(`⚠ Failed to load cached token: ${error.message}`);
      return null;
    }
  }

  /**
   * Clear cached token
   */
  clear(): void {
    try {
      if (fs.existsSync(this.cacheFile)) {
        fs.unlinkSync(this.cacheFile);
        console.log('✓ Token cache cleared');
      }
    } catch (error: any) {
      console.warn(`⚠ Failed to clear cache: ${error.message}`);
    }
  }

  /**
   * Get cache status
   */
  getStatus(): string {
    const token = this.load();
    if (!token) {
      return 'No cached token';
    }

    const expiresIn = Math.floor((token.expiresOn * 1000 - Date.now()) / 1000 / 60);
    return `Token valid for ${expiresIn} minutes (${token.account.username})`;
  }
}
