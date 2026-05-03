#!/usr/bin/env node
/**
 * Fixture-driven tests for renderProducts() in src/json-loader.ts.
 *
 * 13 cases covering the spec in /Users/aaanerud/.claude/plans/sunny-rolling-minsky.md:
 *   - full data, partial fields, falsy names, empty/undefined/non-array inputs
 *   - idempotency (pre-rendered string input)
 *   - special-character sanitization (newlines replaced, quotes preserved)
 *   - regression: anders.n (5 products) — IRON RULE
 *   - regression: GTIN casing artifact (gTIN) — IRON RULE (outside-voice finding 1)
 *   - size cap: >8KB truncated
 *   - camelCase input bypass path
 *
 * Usage: node tools/debug/test-products-renderer.mjs
 * Exit: 0 on all pass, 1 on any failure.
 */

import { strict as assert } from 'assert';
import { renderCollectionProperties } from '../../dist/json-loader.js';

let passed = 0;
let failed = 0;
const failures = [];

function test(name, fn) {
  try {
    fn();
    console.log(`  ✓ ${name}`);
    passed++;
  } catch (err) {
    console.log(`  ✗ ${name}`);
    console.log(`      ${err.message}`);
    failures.push({ name, error: err });
    failed++;
  }
}

function render(input, fieldName = 'products') {
  const r = { [fieldName]: input };
  renderCollectionProperties(r, [fieldName]);
  return r[fieldName];
}

console.log('\nrenderProducts fixture tests\n');

// ─── Case 1: Full data (name + model + gtin), single item ───
test('1. full data single item (YAML)', () => {
  const out = render([{ name: 'Echo Vault', model: 'Model A', gtin: '96128' }]);
  assert.equal(out, '- name: Echo Vault\n  model: Model A\n  gtin: 96128');
});

// ─── Case 2: Name only (no model, no gtin) ───
test('2. name only, no model or gtin', () => {
  const out = render([{ name: 'Echo Vault' }]);
  assert.equal(out, '- name: Echo Vault');
});

// ─── Case 3: Name + model, no gtin ───
test('3. name + model, no gtin', () => {
  const out = render([{ name: 'Echo Vault', model: 'Model A' }]);
  assert.equal(out, '- name: Echo Vault\n  model: Model A');
});

// ─── Case 4: All entries have falsy names → empty string ───
test('4. entries with only empty/null values → filtered; one with model only → renders', () => {
  const out = render([{ name: '' }, { name: null }, { model: 'X' }]);
  assert.equal(out, '- model: X');
});

// ─── Case 5: Empty array → empty string ───
test('5. empty array → empty string', () => {
  const out = render([]);
  assert.equal(out, '');
});

// ─── Case 6: undefined → products stays undefined (no mutation) ───
test('6. undefined → no mutation', () => {
  const r = {};
  renderCollectionProperties(r, ['products']);
  assert.equal(r.products, undefined);
});

// ─── Case 7: Already a string → idempotent, no mutation ───
test('7. idempotent on pre-rendered string', () => {
  const existing = '- name: X\n  model: Y\n  gtin: 123';
  const r = { products: existing };
  renderCollectionProperties(r, ['products']);
  assert.equal(r.products, existing);
});

// ─── Case 8: Object (not array) → no mutation ───
test('8. non-array object → no mutation', () => {
  const r = { products: { not: 'array' } };
  renderCollectionProperties(r, ['products']);
  assert.deepEqual(r.products, { not: 'array' });
});

// ─── Case 9: Special chars — quotes/backslashes preserved, newlines sanitized ───
test('9. special chars: quotes/backslashes passthrough, newlines → space', () => {
  const out = render([
    { name: 'Quote"Test', model: 'Back\\slash', gtin: 'Line\nBreak' },
  ]);
  // Quotes and backslashes preserved verbatim; \n becomes a single space
  assert.ok(out.includes('Quote"Test'), 'quotes should be preserved');
  assert.ok(out.includes('Back\\slash'), 'backslashes should be preserved');
  assert.ok(out.includes('Line Break'), 'newline should be replaced with single space');
  assert.ok(!out.includes('Line\nBreak'), 'raw newline should NOT appear in field value');
});

// ─── Case 10: [CRITICAL REGRESSION] anders.n fixture — 5 full products ───
test('10. [REGRESSION] anders.n fixture preserves Name + Model + GTIN', () => {
  const andersFixture = [
    { name: 'Echo Vault', model: 'Model A', gtin: '96128715782278' },
    { name: 'Falcon Ridge', model: 'Model Y', gtin: '95441323966177' },
    { name: 'Nova Pulse', model: 'Model D', gtin: '25101503687956' },
    { name: 'Orbit Stack', model: 'Model F', gtin: '10819205373202' },
    { name: 'Vertex Signal', model: 'Model A', gtin: '90282932596751' },
  ];
  const out = render(andersFixture);

  // YAML format: each product is a "- name: X\n  model: Y\n  gtin: Z" block
  const nameLines = [...out.matchAll(/^- name: (.+)$/gm)];
  assert.equal(nameLines.length, 5, `expected 5 YAML items, got ${nameLines.length}`);

  for (const p of andersFixture) {
    assert.ok(out.includes(p.name), `missing name: ${p.name}`);
    assert.ok(out.includes(p.model), `missing model: ${p.model}`);
    assert.ok(out.includes(p.gtin), `missing gtin: ${p.gtin}`);
  }
});

// ─── Case 11: [CRITICAL IRON RULE] GTIN casing artifact from pascalToCamel ───
// pascalToCamel('GTIN') returns 'gTIN' (only first char lowercased).
// The renderer MUST tolerate this via case-insensitive field lookup, otherwise
// real PascalCase configs (using `"GTIN": ...`) silently reproduce the exact
// regression we're fixing. This is the headline IRON RULE test.
test('11. [IRON RULE] GTIN casing: gTIN artifact resolved via case-insensitive lookup', () => {
  const out = render([
    { name: 'Echo Vault', model: 'Model A', gTIN: '96128715782278' }, // post-pascalToCamel
  ]);
  assert.ok(
    out.includes('gtin: 96128715782278'),
    `expected lowercase "gtin: 96128715782278" in output (case-insensitive key lookup); got: ${JSON.stringify(out)}`
  );
});

// Case 11b: All-caps GTIN direct from input (never normalized)
test('11b. all-caps GTIN direct (no normalization)', () => {
  const out = render([{ name: 'Echo Vault', model: 'Model A', GTIN: '123' }]);
  assert.ok(out.includes('gtin: 123'), `expected gtin in output; got: ${out}`);
});

// Case 11c: TitleCase fields from untouched PascalCase input
test('11c. TitleCase fields (Name/Model/GTIN) from unnormalized input', () => {
  const out = render([{ Name: 'Echo Vault', Model: 'Model A', GTIN: '123' }]);
  assert.equal(out, '- name: Echo Vault\n  model: Model A\n  gtin: 123');
});

// ─── Case 12: [SIZE CAP] output capped at 8KB with [TRUNCATED] marker ───
test('12. [SIZE CAP] 500 products → truncated at ~8KB with marker', () => {
  const products = [];
  for (let i = 0; i < 500; i++) {
    products.push({
      name: `Product Number ${i}`,
      model: `Model Variant ${i}`,
      gtin: `${1000000000000 + i}`,
    });
  }
  const out = render(products);
  assert.ok(
    Buffer.byteLength(out, 'utf-8') <= 8192,
    `output exceeded 8KB cap (was ${Buffer.byteLength(out, 'utf-8')} bytes)`
  );
  assert.ok(out.endsWith('[TRUNCATED]'), `expected [TRUNCATED] suffix; got end: ${out.slice(-50)}`);
  assert.ok(out.startsWith('- name:'), 'first YAML item preserved');
});

// ─── Case 13: camelCase input path — renderer runs regardless of PascalCase detection ───
// Verifies that calling renderProducts directly on a camelCase-shaped record
// (no PascalCase keys anywhere) still produces the rendered string. This guards
// against the bug where the renderer lived inside normalizePascalRecord and
// was skipped entirely for camelCase input.
test('13. camelCase input path: renderer works standalone', () => {
  const r = {
    email: 'user@example.com',
    products: [{ name: 'Alpha', model: 'Beta', gtin: '42' }],
  };
  renderCollectionProperties(r, ['products']);
  assert.equal(typeof r.products, 'string');
  assert.ok(r.products.includes('Alpha'), 'value preserved as-is');
  assert.ok(r.products.includes('Beta'), 'value preserved as-is');
  assert.ok(r.products.includes('42'));
});

// ─── Case 14: Generic collection — non-Products field renders the same way ───
test('14. generic collection: achievements renders as YAML', () => {
  const out = render([
    { Title: 'Best Paper Award', Year: '2024', Conference: 'ICSE' },
    { Title: 'Patent Filed', Year: '2025' },
  ], 'achievements');
  assert.ok(out.includes('- title: Best Paper Award'), 'first item name');
  assert.ok(out.includes('  year: 2024'), 'first item year indented');
  assert.ok(out.includes('  conference: ICSE'), 'first item conf indented');
  assert.ok(out.includes('- title: Patent Filed'), 'second item');
  assert.ok(out.includes('  year: 2025'), 'second item year');
});

// ─── Summary ───────────────────────────────────────────────────────────────
console.log(`\n${passed} passed, ${failed} failed`);
if (failed > 0) {
  console.log('\nFailed:');
  for (const f of failures) {
    console.log(`  - ${f.name}`);
    console.log(`    ${f.error.stack}`);
  }
  process.exit(1);
}
process.exit(0);
