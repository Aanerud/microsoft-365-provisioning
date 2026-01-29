import express from 'express';
import { Server } from 'http';
import fs from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';
import { TokenCache } from './token-cache.js';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

export interface BrowserAuthConfig {
  tenantId: string;
  clientId: string;
  port?: number;
  scopes?: string[];
  forceRefresh?: boolean;
}

export interface AuthResult {
  accessToken: string;
  expiresOn?: number;
  account: any;
}

export class BrowserAuthServer {
  private app: express.Application;
  private server?: Server;
  private authTimeout?: ReturnType<typeof setTimeout>;
  private config: Required<BrowserAuthConfig> & { forceRefresh: boolean };
  private resolveAuth?: (result: AuthResult) => void;
  private rejectAuth?: (error: Error) => void;
  private tokenCache: TokenCache;

  constructor(config: BrowserAuthConfig) {
    this.config = {
      ...config,
      port: config.port || 5544,
      scopes: config.scopes || [
        'User.ReadWrite.All',
        'Directory.ReadWrite.All',
        'Organization.Read.All',
        'offline_access',
      ],
      forceRefresh: config.forceRefresh || false,
    };

    this.tokenCache = new TokenCache();
    this.app = express();
    this.app.use(express.json());
    this.setupRoutes();
  }

  private setupRoutes(): void {
    // Serve the authentication page
    this.app.get('/', (_req, res) => {
      const htmlPath = path.join(__dirname, '../../public/auth.html');
      let htmlContent = fs.readFileSync(htmlPath, 'utf8');

      // Inject environment variables
      htmlContent = htmlContent.replace(/{{TENANT_ID}}/g, this.config.tenantId);
      htmlContent = htmlContent.replace(/{{CLIENT_ID}}/g, this.config.clientId);
      htmlContent = htmlContent.replace(/{{REDIRECT_URI}}/g, `http://localhost:${this.config.port}`);
      htmlContent = htmlContent.replace(/{{SCOPES}}/g, JSON.stringify(this.config.scopes));

      res.send(htmlContent);
    });

    // Receive token from browser
    this.app.post('/auth/callback', (req, res) => {
      const { accessToken, account, error } = req.body;

      if (error) {
        res.json({ success: false, message: 'Authentication failed' });
        if (this.rejectAuth) {
          this.rejectAuth(new Error(error));
        }
        this.stopServer();
        return;
      }

      if (!accessToken) {
        res.json({ success: false, message: 'No access token received' });
        if (this.rejectAuth) {
          this.rejectAuth(new Error('No access token received'));
        }
        this.stopServer();
        return;
      }

      res.json({ success: true, message: 'Authentication successful! You can close this window.' });

      if (this.resolveAuth) {
        // Cache the token
        const expiresOn = Math.floor(Date.now() / 1000) + 3600; // 1 hour from now
        this.tokenCache.save({
          accessToken,
          expiresOn,
          account,
        });

        this.resolveAuth({ accessToken, expiresOn, account });
      }

      // Close server after successful auth
      setTimeout(() => this.stopServer(), 1000);
    });
  }

  /**
   * Start the authentication server and wait for user to authenticate
   * Checks cache first unless forceRefresh is true
   */
  async authenticate(): Promise<AuthResult> {
    // Check cache first
    if (!this.config.forceRefresh) {
      const cachedToken = this.tokenCache.load();
      if (cachedToken) {
        console.log(`\n✅ Using cached token (${cachedToken.account.username})`);
        const expiresIn = Math.floor((cachedToken.expiresOn * 1000 - Date.now()) / 1000 / 60);
        console.log(`   Token valid for ${expiresIn} minutes\n`);
        return {
          accessToken: cachedToken.accessToken,
          expiresOn: cachedToken.expiresOn,
          account: cachedToken.account,
        };
      }
    }

    return new Promise((resolve, reject) => {
      this.resolveAuth = resolve;
      this.rejectAuth = reject;

      this.server = this.app.listen(this.config.port, () => {
        const url = `http://localhost:${this.config.port}`;
        console.log(`\n🔐 Authentication Required\n`);
        console.log(`Opening browser for Microsoft authentication...`);
        console.log(`If browser doesn't open, visit: ${url}\n`);

        // Try to open browser automatically
        this.openBrowser(url);
      });

      // Timeout after 5 minutes
      this.authTimeout = setTimeout(() => {
        if (this.server) {
          reject(new Error('Authentication timeout after 5 minutes'));
          this.stopServer();
        }
      }, 5 * 60 * 1000);
    });
  }

  /**
   * Clear cached token
   */
  clearCache(): void {
    this.tokenCache.clear();
  }

  private openBrowser(url: string): void {
    const open = async (url: string) => {
      const { default: openModule } = await import('open');
      await openModule(url);
    };

    open(url).catch((error) => {
      console.warn(`Could not open browser automatically: ${error.message}`);
      console.log(`Please open manually: ${url}`);
    });
  }

  private stopServer(): void {
    if (this.server) {
      this.server.close();
      this.server = undefined;
    }
    if (this.authTimeout) {
      clearTimeout(this.authTimeout);
      this.authTimeout = undefined;
    }
  }
}
