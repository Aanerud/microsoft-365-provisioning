#!/usr/bin/env node
import { ClientSecretCredential } from '@azure/identity';
import dotenv from 'dotenv';
dotenv.config();

async function diagnose() {
  console.log('🔍 Graph Connector Service Diagnosis');
  console.log('=====================================\n');

  const credential = new ClientSecretCredential(
    process.env.AZURE_TENANT_ID,
    process.env.AZURE_CLIENT_ID,
    process.env.AZURE_CLIENT_SECRET
  );

  const token = await credential.getToken('https://graph.microsoft.com/.default');

  const endpoints = [
    { name: 'v1.0 list connections', url: 'https://graph.microsoft.com/v1.0/external/connections' },
    { name: 'beta list connections', url: 'https://graph.microsoft.com/beta/external/connections' },
  ];

  for (const ep of endpoints) {
    console.log('Testing:', ep.name);
    try {
      const start = Date.now();
      const response = await fetch(ep.url, {
        headers: { 'Authorization': 'Bearer ' + token.token }
      });
      const elapsed = Date.now() - start;

      console.log('  Status:', response.status, '(' + elapsed + 'ms)');

      if (response.status !== 200) {
        const body = await response.text();
        console.log('  Error:', body.substring(0, 200));
      } else {
        const data = await response.json();
        console.log('  Connections found:', data.value?.length || 0);
      }
    } catch (error) {
      console.log('  Network error:', error.message);
    }
    console.log('');
  }

  // Try /organization to confirm auth works
  console.log('Testing: /organization (to confirm auth)');
  try {
    const response = await fetch('https://graph.microsoft.com/v1.0/organization', {
      headers: { 'Authorization': 'Bearer ' + token.token }
    });
    console.log('  Status:', response.status);
    if (response.status === 200) {
      const data = await response.json();
      console.log('  Org:', data.value?.[0]?.displayName);
    }
  } catch (error) {
    console.log('  Error:', error.message);
  }
}

diagnose().catch(console.error);
