import test from 'node:test';
import assert from 'node:assert/strict';
import { Workbook } from '../src/workbook.js';
import { numberRaw, readXlsx, writeXlsx } from '../src/xlsx.js';
import { parseInput } from '../src/format.js';

const values = [0.1, 0.25, 1 / 3, 0.12345678901234567, 1e-16, 1.234567890123456e-10, -0.9876543210987654, 1.234567890123456e15];

test('percentage import retains exact numeric values, including fractions smaller than tolerance', () => {
  for (const value of values) {
    assert.ok(Object.is(parseInput(numberRaw(value, { numFmt: 'percent' })).value, value), String(value));
  }
  assert.equal(numberRaw(0.25, { numFmt: 'percent' }), '25%');
});

test('percentage values and formats survive two XLSX save/reopen cycles without rounding', () => {
  const wb = new Workbook();
  wb.transact(() => values.forEach((value, r) => {
    wb.setInput(0, r, 0, String(value));
    wb.setStyle(0, r, 0, { numFmt: 'percent', decimals: 8 });
  }));
  let back = wb;
  for (let pass = 0; pass < 2; pass++) {
    back = new Workbook(readXlsx(writeXlsx(back)).data);
    values.forEach((value, r) => {
      assert.ok(Object.is(back.getValue(0, r, 0), value), `${pass}:${r}`);
      assert.equal(back.styleAt(0, r, 0).numFmt, 'percent');
      assert.equal(back.styleAt(0, r, 0).decimals, 8);
    });
  }
});
