#!/usr/bin/env node
import fs from 'fs/promises';

const testEmail = 'ingrid.johansen@a830edad9050849coep9vqp9bog.onmicrosoft.com';
const TOKEN_CACHE_PATH = `${process.env.HOME}/.m365-provision/token-cache.json`;

async function queryExtensions() {
  // Load token
  const tokenData = JSON.parse(await fs.readFile(TOKEN_CACHE_PATH, 'utf-8'));
  const token = tokenData.accessToken;

  console.log(`🔍 Testing Extensions for: ${testEmail}\n`);

  // Query user with extensions
  const response = await fetch(
    `https://graph.microsoft.com/beta/users/${testEmail}?$expand=extensions`,
    {
      headers: { Authorization: `Bearer ${token}` },
    }
  );

  if (!response.ok) {
    throw new Error(`HTTP ${response.status}: ${await response.text()}`);
  }

  const user = await response.json();

  console.log('✅ User found:', user.displayName);
  console.log('   ID:', user.id);
  console.log('');

  // Check extensions
  if (user.extensions && user.extensions.length > 0) {
    console.log(`✅ Found ${user.extensions.length} extension(s):\n`);

    user.extensions.forEach((ext, i) => {
      console.log(`Extension ${i + 1}:`);
      console.log('  Extension Name:', ext.extensionName || ext.id);
      console.log('  Type:', ext['@odata.type']);

      // Show custom properties
      const { '@odata.type': _, extensionName, id, ...props } = ext;
      console.log('  Custom Properties:');
      Object.entries(props).forEach(([key, val]) => {
        console.log(`    ${key}: ${val}`);
      });
      console.log('');
    });
  } else {
    console.log('❌ NO EXTENSIONS FOUND');
    console.log('   This means open extensions are NOT working!\n');
  }
}

queryExtensions().catch(err => {
  console.error('❌ Error:', err.message);
  process.exit(1);
});
