#!/usr/bin/env node
/**
 * Fetch complete user profile from Microsoft Graph
 * Uses cached delegated token from browser auth
 *
 * Usage: node tools/debug/fetch-user-profile.mjs [email]
 *
 * If token is expired, run: npm run test-connection (to refresh token)
 */

import fs from 'fs/promises';

const TOKEN_CACHE_PATH = `${process.env.HOME}/.m365-provision/token-cache.json`;

async function getToken() {
  try {
    const tokenData = JSON.parse(await fs.readFile(TOKEN_CACHE_PATH, 'utf-8'));
    // expiresOn is Unix timestamp in seconds
    const expiresOn = new Date(tokenData.expiresOn * 1000);

    if (expiresOn < new Date()) {
      console.error('Token expired at:', expiresOn.toLocaleString());
      console.error('Run: npm run test-connection   (to refresh token)');
      process.exit(1);
    }

    console.log('Token expires:', expiresOn.toLocaleString());
    return tokenData.accessToken;
  } catch (err) {
    console.error('No cached token found. Run: npm run test-connection');
    process.exit(1);
  }
}

async function fetchUserProfile(email, token) {
  const results = {};

  // 1. Fetch user with beta properties
  console.log('\n1. GET /beta/users/{email}');
  const userResponse = await fetch(
    `https://graph.microsoft.com/beta/users/${email}`,
    { headers: { Authorization: `Bearer ${token}` } }
  );
  console.log(`   Status: ${userResponse.status}`);

  if (!userResponse.ok) {
    const error = await userResponse.text();
    throw new Error(`Failed to fetch user: ${userResponse.status} - ${error}`);
  }
  results.user = await userResponse.json();

  // 2. Fetch extensions (open extensions / schema extensions)
  console.log('\n2. GET /beta/users/{email}/extensions');
  const extResponse = await fetch(
    `https://graph.microsoft.com/beta/users/${email}/extensions`,
    { headers: { Authorization: `Bearer ${token}` } }
  );
  console.log(`   Status: ${extResponse.status}`);
  if (extResponse.ok) {
    results.extensions = await extResponse.json();
  }

  // 3. Fetch profile (People API - skills, languages, interests)
  console.log('\n3. GET /beta/users/{email}/profile');
  const profileResponse = await fetch(
    `https://graph.microsoft.com/beta/users/${email}/profile`,
    { headers: { Authorization: `Bearer ${token}` } }
  );
  console.log(`   Status: ${profileResponse.status}`);
  if (profileResponse.ok) {
    results.profile = await profileResponse.json();
  } else {
    const err = await profileResponse.text();
    console.log(`   Error: ${err.substring(0, 200)}`);
  }

  // 4. Fetch skills specifically
  console.log('\n4. GET /beta/users/{email}/profile/skills');
  const skillsResponse = await fetch(
    `https://graph.microsoft.com/beta/users/${email}/profile/skills`,
    { headers: { Authorization: `Bearer ${token}` } }
  );
  console.log(`   Status: ${skillsResponse.status}`);
  if (skillsResponse.ok) {
    results.skills = await skillsResponse.json();
  }

  // 5. Fetch languages
  console.log('\n5. GET /beta/users/{email}/profile/languages');
  const langResponse = await fetch(
    `https://graph.microsoft.com/beta/users/${email}/profile/languages`,
    { headers: { Authorization: `Bearer ${token}` } }
  );
  console.log(`   Status: ${langResponse.status}`);
  if (langResponse.ok) {
    results.languages = await langResponse.json();
  }

  return results;
}

// Main
const email = process.argv[2];
if (!email) {
  console.log('Usage: node tools/debug/fetch-user-profile.mjs <email>');
  console.log('Example: node tools/debug/fetch-user-profile.mjs nora.d@yourdomain.onmicrosoft.com');
  process.exit(1);
}
console.log(`\nFetching profile for: ${email}`);
console.log('─'.repeat(60));

try {
  const token = await getToken();
  const results = await fetchUserProfile(email, token);

  console.log('\n' + '='.repeat(80));
  console.log('COMPLETE USER PROFILE OUTPUT (BETA)');
  console.log('='.repeat(80));

  console.log('\n--- USER (BETA) ---');
  console.log(JSON.stringify(results.user, null, 2));

  if (results.extensions?.value?.length > 0) {
    console.log('\n--- EXTENSIONS ---');
    console.log(JSON.stringify(results.extensions, null, 2));
  } else {
    console.log('\n--- EXTENSIONS: None ---');
  }

  if (results.profile) {
    console.log('\n--- PROFILE (People API) ---');
    console.log(JSON.stringify(results.profile, null, 2));
  } else {
    console.log('\n--- PROFILE: Not accessible ---');
  }

  if (results.skills?.value?.length > 0) {
    console.log('\n--- SKILLS ---');
    console.log(JSON.stringify(results.skills, null, 2));
  } else {
    console.log('\n--- SKILLS: None ---');
  }

  if (results.languages?.value?.length > 0) {
    console.log('\n--- LANGUAGES ---');
    console.log(JSON.stringify(results.languages, null, 2));
  } else {
    console.log('\n--- LANGUAGES: None ---');
  }

} catch (error) {
  console.error('\nError:', error.message);
  process.exit(1);
}
