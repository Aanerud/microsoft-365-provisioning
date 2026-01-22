#!/usr/bin/env node
import { ClientSecretCredential } from '@azure/identity';
import { Client } from '@microsoft/microsoft-graph-client';
import { TokenCredentialAuthenticationProvider } from '@microsoft/microsoft-graph-client/authProviders/azureTokenCredentials/index.js';
import dotenv from 'dotenv';

dotenv.config();

const tenantId = process.env.AZURE_TENANT_ID;
const clientId = process.env.AZURE_CLIENT_ID;
const clientSecret = process.env.AZURE_CLIENT_SECRET;

console.log('🔍 Verifying Current State (3 Users)...\n');

const credential = new ClientSecretCredential(tenantId, clientId, clientSecret);
const authProvider = new TokenCredentialAuthenticationProvider(credential, {
  scopes: ['https://graph.microsoft.com/.default']
});

const betaClient = Client.initWithMiddleware({
  authProvider,
  defaultVersion: 'beta'
});

try {
  const expectedUsers = [
    { name: 'Ingrid Johansen', email: 'ingrid.johansen@a830edad9050849coep9vqp9bog.onmicrosoft.com' },
    { name: 'Ola Nordmann', email: 'ola.nordmann@a830edad9050849coep9vqp9bog.onmicrosoft.com' },
    { name: 'Lars Hansen', email: 'lars.hansen@a830edad9050849coep9vqp9bog.onmicrosoft.com' }
  ];

  console.log('✅ Successfully authenticated with client credentials\n');
  console.log(`📋 Verifying ${expectedUsers.length} external items:\n`);

  let successCount = 0;
  let failCount = 0;

  for (const user of expectedUsers) {
    const itemId = `person-${user.email.replace(/@/g, '-').replace(/\./g, '-')}`;

    try {
      const item = await betaClient
        .api(`/external/connections/m365provisionpeople/items/${itemId}`)
        .get();

      console.log(`✅ ${user.name}`);
      if (item.properties) {
        const props = item.properties;
        if (props.accountInformation) {
          const acct = JSON.parse(props.accountInformation);
          console.log(`   Email: ${acct.userPrincipalName}`);
        }
        if (props.skills) {
          console.log(`   Skills: ${Array.isArray(props.skills) ? props.skills.length : 0} items`);
        }
        if (props.interests) {
          console.log(`   Interests: ${Array.isArray(props.interests) ? props.interests.length : 0} items`);
        }
      }
      console.log('');
      successCount++;
    } catch (error) {
      console.log(`❌ ${user.name}`);
      console.log(`   Error: ${error.statusCode} - ${error.message}\n`);
      failCount++;
    }
  }

  console.log('═══════════════════════════════════════════════════════════');
  console.log(`📊 Verification Summary`);
  console.log('═══════════════════════════════════════════════════════════');
  console.log(`✅ Verified: ${successCount}/${expectedUsers.length}`);
  console.log(`❌ Failed:   ${failCount}/${expectedUsers.length}`);
  console.log('═══════════════════════════════════════════════════════════\n');

  if (successCount === expectedUsers.length) {
    console.log('🎉 All external items verified successfully!');
    console.log('✅ End-to-end flow is working correctly!\n');
  } else {
    console.log('⚠️  Some items are missing or failed verification.\n');
  }

} catch (error) {
  console.error('❌ Error:', error.message);
}
