import { test } from 'node:test';
import assert from 'node:assert/strict';
import { normalizeFormula, FUNCTION_NAMES } from '../src/formula.js';
import { makeSeries } from '../src/series.js';
import { parseDelimited, toDelimited, guessDelimiter } from '../src/csv.js';
import { FUNC_INFO } from '../src/funcinfo.js';

test('수식 정리', () => {
  assert.equal(normalizeFormula('=sum(a1:b2'), '=SUM(A1:B2)');
  assert.equal(normalizeFormula('=if(a1>0,"abc(","x"'), '=IF(A1>0,"abc(","x")');
  assert.equal(normalizeFormula("='my sheet'!a1+true"), "='my sheet'!A1+TRUE");
  assert.equal(normalizeFormula('="hi'), '="hi"');
});

const seq = (raws, values, style) => raws.map((raw, i) => ({ data: raw === null ? null : { raw, style }, value: values[i], pos: i }));
const run = (s, n, forward = true, axis = 'row') => {
  const gen = makeSeries(s, forward, axis);
  return Array.from({ length: n }, (_, k) => gen(k, forward ? s.length + k : -1 - k)?.raw ?? null);
};

test('채우기: 숫자 계열', () => {
  assert.deepEqual(run(seq(['1', '3'], [1, 3]), 3), ['5', '7', '9']);
  assert.deepEqual(run(seq(['5'], [5]), 2), ['5', '5']);
  assert.deepEqual(run(seq(['0.1', '0.2'], [0.1, 0.2]), 2), ['0.3', '0.4']);
  assert.deepEqual(run(seq(['10', '20'], [10, 20]), 2, false), ['0', '-10']);
  assert.deepEqual(run(seq(['45000'], [45000], { numFmt: 'date' }), 2), ['45001', '45002']);
});

test('채우기: 목록과 텍스트 번호', () => {
  assert.deepEqual(run(seq(['월'], ['월']), 3), ['화', '수', '목']);
  assert.deepEqual(run(seq(['11월'], ['11월']), 3), ['12월', '1월', '2월']);
  assert.deepEqual(run(seq(['항목1'], ['항목1']), 2), ['항목2', '항목3']);
  assert.deepEqual(run(seq(['A-009'], ['A-009']), 2), ['A-010', 'A-011']);
  assert.deepEqual(run(seq(['사과', '배'], ['사과', '배']), 3), ['사과', '배', '사과']);
});

test('채우기: 수식 복사', () => {
  const s = [{ data: { raw: '=A1*2' }, value: 2, pos: 0 }];
  assert.deepEqual(run(s, 2), ['=A2*2', '=A3*2']);
  assert.deepEqual(run(s, 1, true, 'col'), ['=B1*2']);
});

test('CSV/TSV', () => {
  assert.deepEqual(parseDelimited('a,b\n"c,d","e ""q"""\n'), [['a', 'b'], ['c,d', 'e "q"']]);
  assert.deepEqual(parseDelimited('x\t"멀티\n라인"\ty', '\t'), [['x', '멀티\n라인', 'y']]);
  assert.equal(toDelimited([['a', 'b,c'], ['"q"', '']]), 'a,"b,c"\r\n"""q""",');
  assert.equal(guessDelimiter('a\tb\tc\n1\t2\t3'), '\t');
  assert.equal(guessDelimiter('a,b\n1,2'), ',');
});

test('모든 함수에 설명이 있음', () => {
  const missing = FUNCTION_NAMES.filter((n) => !FUNC_INFO[n]);
  assert.deepEqual(missing, []);
});

test('한 파일 배포본 빌드', async () => {
  const { execFileSync } = await import('node:child_process');
  const { mkdtempSync, readFileSync } = await import('node:fs');
  const { tmpdir } = await import('node:os');
  const { join } = await import('node:path');
  const { fileURLToPath } = await import('node:url');
  const out = mkdtempSync(join(tmpdir(), 'tabula-'));
  execFileSync(process.execPath, [fileURLToPath(new URL('../build.mjs', import.meta.url)), out]);
  const html = readFileSync(join(out, 'index.html'), 'utf8');
  assert.ok(html.includes('__mods["src/app.js"]'));
  assert.ok(!/^\s*(import|export)\s/m.test(html.slice(html.indexOf('const __mods'))));
  assert.ok(!html.includes('href="styles.css"'));
});
