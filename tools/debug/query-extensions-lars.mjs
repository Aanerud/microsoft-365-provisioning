#!/usr/bin/env node
import fs from 'fs/promises';

const testEmail = 'lars.hansen@a830edad9050849coep9vqp9bog.onmicrosoft.com';
const TOKEN_CACHE_PATH = `${process.env.HOME}/.m365-provision/token-cache.json`;

async function queryExtensions() {
  const tokenData = JSON.parse(await fs.readFile(TOKEN_CACHE_PATH, 'utf-8'));
  const token = tokenData.accessToken;

  console.log(`🔍 Testing Extensions for: ${testEmail}\n`);

  const response = await fetch(
    `https://graph.microsoft.com/beta/users/${testEmail}?$expand=extensions`,
    { headers: { Authorization: `Bearer ${token}` } }
  );

  const user = await response.json();

  console.log('✅ User:', user.displayName);
  console.log('');

  if (user.extensions && user.extensions.length > 0) {
    console.log(`✅ Found ${user.extensions.length} extension(s):\n`);

    user.extensions.forEach((ext) => {
      const { '@odata.type': _, extensionName, id, ...props } = ext;
      console.log('Custom Properties:');
      Object.entries(props).forEach(([key, val]) => {
        console.log(`  ${key}: ${val}`);
      });
    });
  } else {
    console.log('❌ NO EXTENSIONS');
  }
}

queryExtensions().catch(console.error);
