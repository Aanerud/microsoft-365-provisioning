#!/usr/bin/env node
import fs from 'fs';
import dotenv from 'dotenv';

dotenv.config();

const cachePath = `${process.env.HOME}/.m365-provision/token-cache.json`;
const cacheData = JSON.parse(fs.readFileSync(cachePath, 'utf-8'));

console.log('\n🔍 Token Information:\n');
console.log('Account:', cacheData.account?.username);
console.log('Token expires:', new Date(cacheData.expiresOn).toLocaleString());
console.log('\nScopes requested:', cacheData.scopes);

// Decode JWT to see actual scopes
const token = cacheData.accessToken;
const parts = token.split('.');
if (parts.length === 3) {
  const payload = JSON.parse(Buffer.from(parts[1], 'base64').toString());
  console.log('\nActual scopes in token:');
  if (payload.scp) {
    const scopes = payload.scp.split(' ');
    scopes.forEach(scope => console.log(`  - ${scope}`));
  } else {
    console.log('  No scopes found in token');
  }

  console.log('\nToken audience:', payload.aud);
  console.log('Token issuer:', payload.iss);
  console.log('App ID:', payload.appid);
}
