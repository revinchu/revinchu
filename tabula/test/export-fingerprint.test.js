import test from 'node:test';
import assert from 'node:assert/strict';
import {createCellFingerprint} from '../tools/export-fingerprint.mjs';
import {makeCell} from '../src/workbook.js';
const fingerprint = (data, content = true) => {const f=createCellFingerprint(content);data.forEach((v,r)=>f.add(makeCell(v),r,0));return f.result();};
test('export integrity compares formula source, not the formula flag',()=>{
  const a=fingerprint([{raw:'=A2+1',cached:5}]),b=fingerprint([{raw:'=A2+2',cached:5}]);
  assert.equal(a.formulas,b.formulas);assert.notEqual(a.formulaHash,b.formulaHash);assert.equal(a.cacheHash,b.cacheHash);
});
test('scalar integrity distinguishes input escapes from literal apostrophes and type changes',()=>{
  const a=fingerprint([{raw:"'2025년 10월"}]),b=fingerprint([{raw:'2025년 10월'}]),literal=fingerprint([{raw:"'2025년 10월",inputType:'text'}]);
  assert.notEqual(a.hash,b.hash);assert.equal(a.valueHash,b.valueHash);assert.notEqual(a.valueHash,literal.valueHash);
  assert.notEqual(fingerprint([{raw:"'123"}]).valueHash,fingerprint([{raw:'123'}]).valueHash);
  assert.notEqual(fingerprint([{raw:"'TRUE"}]).valueHash,fingerprint([{raw:'TRUE'}]).valueHash);
  assert.notEqual(fingerprint([{raw:"'#N/A"}]).valueHash,fingerprint([{raw:'#N/A'}]).valueHash);
});
test('cache integrity distinguishes NaN, null, negative zero and stale values',()=>{
  const hashes=[NaN,null,0,-0].map(cached=>fingerprint([{raw:'=1',cached}]).cacheHash);
  assert.equal(new Set(hashes).size,4);
  assert.notEqual(fingerprint([{raw:'=1',cached:1}]).cacheHash,fingerprint([{raw:'=1',staleCached:1}]).cacheHash);
});
test('performance fingerprint does not inspect raw characters or literal values',()=>{
  const f=createCellFingerprint(false),raw={toString(){assert.fail('premature string flatten');}};
  f.add({raw,formula:true,get v(){assert.fail('premature value evaluation');}},2,3);
  assert.deepEqual(f.result(),{formulas:1,nonempty:1});
});
