#!/usr/bin/env node
import dotenv from 'dotenv';

dotenv.config();

const secret = process.env.AZURE_CLIENT_SECRET;

console.log('🔍 Testing .env loading...\n');
console.log('Secret exists:', !!secret);
console.log('Secret length:', secret?.length);
console.log('Secret first 10 chars:', secret?.substring(0, 10));
console.log('Secret last 10 chars:', secret?.substring(secret.length - 10));
console.log('Secret has spaces:', secret?.includes(' '));
console.log('Secret has newlines:', secret?.includes('\n'));
console.log('Secret has carriage returns:', secret?.includes('\r'));
console.log('\nFull secret (for verification):');
console.log(`"${secret}"`);
