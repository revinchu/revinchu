import test from 'node:test';
import assert from 'node:assert/strict';
import { nativeOpenEvidence } from '../tools/native-open-evidence.mjs';
const complete = () => ({ mode: 'normal', opened: true, readOnly: true, sourceUnchanged: true, errors: [],
  sheets: [{ name: 'Report', pivots: [], shapes: [], sample: [{ address: 'A1', value: 12 }] }] });
test('explicit normal load requires verified actual workbook identity and completed structure', () => {
  const record = { ...complete(), corruptLoad: 0, readinessVerified: true, fullNameMatches: true, readyWorksheetCount: 1, repairMode: null };
  assert.equal(nativeOpenEvidence(record).policy, 'explicit-xlNormalLoad');
  for (const change of [{ fullNameMatches: false }, { readinessVerified: false }, { readyWorksheetCount: 0 }, { sheets: [] }, { readOnly: null }, { sourceUnchanged: false }, { errors: ['busy'] }])
    assert.equal(nativeOpenEvidence({ ...record, ...change }).passed, false);
});
test('legacy normal Open with completed inspection does not depend on nonexistent RepairMode', () => {
  for (const repairMode of [undefined, null, false]) {
    const result = nativeOpenEvidence({ ...complete(), repairMode });
    assert.equal(result.passed, true); assert.equal(result.policy, 'legacy-default-xlNormalLoad');
    assert.equal(result.repairModePropertyIgnored, true);
  }
});
test('repair diagnostics and extraction never become a normal-open pass', () => {
  for (const change of [{ mode: 'repair-diagnostic-only' }, { mode: 'extract-diagnostic-only' }, { repairDiagnostic: true }, { corruptLoad: 1 }, { corruptLoad: 2 }, { repairMode: true }, { repairLogs: ['diagnostic.xml'] }])
    assert.equal(nativeOpenEvidence({ ...complete(), repairMode: false, ...change }).passed, false);
});
test('a false RepairMode or an Open return alone cannot certify a workbook', () => {
  assert.equal(nativeOpenEvidence({ opened: true, readOnly: true, repairMode: false }).passed, false);
  assert.equal(nativeOpenEvidence({ ...complete(), mode: undefined }).passed, false);
  assert.equal(nativeOpenEvidence({ ...complete(), sheets: [{ name: 'Report' }] }).passed, false);
});
