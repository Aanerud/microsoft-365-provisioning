#!/usr/bin/env node
import { Client } from '@microsoft/microsoft-graph-client';
import fs from 'fs';
import dotenv from 'dotenv';

dotenv.config();

const cachePath = `${process.env.HOME}/.m365-provision/token-cache.json`;
const cacheData = JSON.parse(fs.readFileSync(cachePath, 'utf-8'));

const client = Client.init({
  authProvider: (done) => {
    done(null, cacheData.accessToken);
  }
});

console.log('🗑️  Deleting connection m365provisionpeople...');

try {
  await client.api('/external/connections/m365provisionpeople').delete();
  console.log('✅ Connection deleted successfully');
} catch (error) {
  console.log(`❌ Failed: ${error.statusCode} - ${error.message}`);
}
