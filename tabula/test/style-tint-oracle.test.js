import { test } from 'node:test';
import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import { applyTint } from '../src/stylepresets.js';

const oracle = JSON.parse(readFileSync(new URL('./fixtures/excel-tint-oracle.json', import.meta.url), 'utf8'));

test('theme tints match independently measured Excel colors, including nonstandard and boundary values', () => {
  let count = 0;
  for (const { base, expected } of oracle.colors) {
    assert.equal(expected.length, oracle.tints.length);
    for (let i = 0; i < oracle.tints.length; i++) {
      assert.equal(applyTint(base, oracle.tints[i]), expected[i], `${base} tint ${oracle.tints[i]}`);
      assert.equal(applyTint(base, oracle.inputTints[i]), expected[i], `${base} decimal tint ${oracle.inputTints[i]}`);
      count++;
    }
  }
  assert.equal(count, 384);
});

test('light and dark imported theme fills retain Excel channel values', () => {
  assert.equal(applyTint('4472C4', 0.7999816888943144), 'D9E1F2');
  assert.equal(applyTint('4472C4', 0.8), 'D9E1F2');
  assert.equal(applyTint('5B9BD5', -0.499984740745262), '1F4E78');
  // A single floor of the combined positive expression would produce 000000.
  assert.equal(applyTint('000000', 0.00009155552842799158), '010101');
});
