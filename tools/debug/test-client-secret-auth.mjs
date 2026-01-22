#!/usr/bin/env node
import { ClientSecretCredential } from '@azure/identity';
import dotenv from 'dotenv';

dotenv.config();

const tenantId = process.env.AZURE_TENANT_ID;
const clientId = process.env.AZURE_CLIENT_ID;
const clientSecret = process.env.AZURE_CLIENT_SECRET;

console.log('🔍 Testing Client Secret Authentication...\n');
console.log('Tenant ID:', tenantId);
console.log('Client ID:', clientId);
console.log('Secret length:', clientSecret?.length);
console.log('Secret first 10 chars:', clientSecret?.substring(0, 10));
console.log('');

const credential = new ClientSecretCredential(tenantId, clientId, clientSecret);

try {
  console.log('🔐 Attempting to acquire token...');
  const token = await credential.getToken('https://graph.microsoft.com/.default');

  console.log('✅ SUCCESS! Token acquired');
  console.log('Token expires:', new Date(token.expiresOnTimestamp).toLocaleString());
  console.log('\n✅ Client secret is VALID and authentication works!');

} catch (error) {
  console.log('❌ FAILED! Authentication error:');
  console.log('Error:', error.message);
  console.log('\n❌ The client secret is INVALID or doesn\'t match the app registration.');
  console.log('\n📋 To fix:');
  console.log('1. Go to Azure Portal > App registrations > Your app');
  console.log('2. Go to: Certificates & secrets');
  console.log('3. Create a NEW client secret');
  console.log('4. Copy the VALUE (not the ID!)');
  console.log('5. Update .env file: AZURE_CLIENT_SECRET=<new-value>');
}
