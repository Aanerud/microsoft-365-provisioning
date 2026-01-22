#!/usr/bin/env node

/**
 * Test script to verify open extensions are working
 * Queries a user and checks if custom properties (VTeam, BenefitPlan, etc.) are stored
 */

import { Client } from '@microsoft/microsoft-graph-client';
import { TokenCache } from './dist/auth/token-cache.js';
import dotenv from 'dotenv';

dotenv.config();

async function testExtensions() {
  console.log('🔍 Testing Open Extensions\n');

  // Load token from cache
  const tokenCache = new TokenCache();
  const cachedToken = await tokenCache.load();

  if (!cachedToken) {
    console.error('❌ No cached token found. Run: npm run provision first');
    process.exit(1);
  }

  // Initialize Graph client
  const client = Client.init({
    authProvider: (done) => {
      done(null, cachedToken.accessToken);
    },
  });

  // Test user
  const testEmail = 'ingrid.johansen@a830edad9050849coep9vqp9bog.onmicrosoft.com';
  console.log(`Testing user: ${testEmail}\n`);

  try {
    // 1. Get user basic info
    console.log('1️⃣  Fetching user basic info...');
    const user = await client.api(`/users/${testEmail}`).get();
    console.log(`   ✓ User found: ${user.displayName} (${user.id})`);
    console.log('');

    // 2. Get ALL extensions for this user
    console.log('2️⃣  Fetching ALL extensions...');
    const extensions = await client.api(`/users/${user.id}/extensions`).get();

    if (extensions.value && extensions.value.length > 0) {
      console.log(`   ✓ Found ${extensions.value.length} extension(s):\n`);
      extensions.value.forEach((ext, index) => {
        console.log(`   Extension ${index + 1}:`);
        console.log(`     Extension Name: ${ext.extensionName || ext.id}`);
        console.log(`     Type: ${ext['@odata.type']}`);
        console.log(`     Properties:`, JSON.stringify(ext, null, 6));
        console.log('');
      });
    } else {
      console.log('   ❌ No extensions found');
    }

    // 3. Try to get specific extension
    console.log('3️⃣  Fetching specific extension: com.m365provision.customFields');
    try {
      const customExt = await client
        .api(`/users/${user.id}/extensions/com.m365provision.customFields`)
        .get();

      console.log('   ✓ Found custom fields extension:\n');

      // Extract custom properties (remove metadata)
      const { '@odata.type': _type, extensionName, id, ...customProps } = customExt;

      console.log('   Custom Properties:');
      Object.entries(customProps).forEach(([key, value]) => {
        console.log(`     ${key}: ${value}`);
      });
      console.log('');

      // Verify expected properties
      const expectedProps = ['VTeam', 'BenefitPlan', 'CostCenter', 'BuildingAccess', 'ProjectCode'];
      const foundProps = Object.keys(customProps);
      const missingProps = expectedProps.filter(p => !foundProps.includes(p));

      if (missingProps.length === 0) {
        console.log('   ✅ All expected custom properties found!');
      } else {
        console.log(`   ⚠️  Missing properties: ${missingProps.join(', ')}`);
      }

    } catch (error) {
      if (error.statusCode === 404) {
        console.log('   ❌ Extension "com.m365provision.customFields" not found');
        console.log('   This means open extensions are NOT working for this user');
      } else {
        throw error;
      }
    }

  } catch (error) {
    console.error('\n❌ Error:', error.message);
    if (error.statusCode) {
      console.error(`   Status: ${error.statusCode}`);
    }
    process.exit(1);
  }
}

testExtensions().catch(console.error);
