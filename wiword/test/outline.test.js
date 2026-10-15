import { test } from 'node:test';
import assert from 'node:assert/strict';
import { parseOutline, buildFromOutline, numberLabel, stripNumber, outlineFromDoc } from '../src/outline.js';
import { paraPlain } from '../src/model.js';

const heads = (doc) => doc.body.filter((b) => b.type === 'p' && /^Heading\d$/.test(b.style ?? '')).map((b) => [b.style, paraPlain(b)]);

test('numberLabel schemes', () => {
  assert.equal(numberLabel('decimal', [2]), '2.');
  assert.equal(numberLabel('decimal', [2, 3]), '2.3');
  assert.equal(numberLabel('plan', [3]), 'Ⅲ.');
  assert.equal(numberLabel('plan', [1, 2, 3]), '다.');
  assert.equal(numberLabel('gongmun', [1, 2]), '나.');
  assert.equal(numberLabel('gongmun', [1, 1, 1, 2]), '나)');
  assert.equal(numberLabel('none', [1]), '');
  assert.equal(stripNumber('1.2 시장 동향'), '시장 동향');
  assert.equal(stripNumber('가. 일정'), '일정');
});

test('parseOutline reads # headings and numbered lines', () => {
  const a = parseOutline('# 제목\n## 장\n본문\n### 절');
  assert.deepEqual(a.map((x) => x.kind + (x.level ?? '')), ['h1', 'h2', 'md', 'h3']);
  const b = parseOutline('1. 개요\n1.1 배경\n이 문장은 본문입니다.\n2. 결론');
  assert.deepEqual(b.filter((x) => x.kind === 'h').map((x) => [x.level, x.text]), [[1, '개요'], [2, '배경'], [1, '결론']]);
});

test('buildFromOutline: title, numbering, cover, toc, placeholders', () => {
  const doc = buildFromOutline('# 사업 계획\n## 개요\n## 현황\n### 시장\n### 경쟁\n## 계획', { kind: 'report' });
  assert.equal(doc.props.title, '사업 계획');
  assert.deepEqual(heads(doc), [['Heading1', '1. 개요'], ['Heading1', '2. 현황'], ['Heading2', '2.1 시장'], ['Heading2', '2.2 경쟁'], ['Heading1', '3. 계획']]);
  assert.equal(doc.body.find((b) => b.style === 'Title') && paraPlain(doc.body.find((b) => b.style === 'Title')), '사업 계획');
  const text = doc.body.filter((b) => b.type === 'p').map(paraPlain);
  assert.equal(text.filter((t) => t === '(여기에 내용을 쓰세요)').length, 4);
  assert.ok(doc.body.some((b) => /^TOC/.test(b.style ?? '') || b.type === 'sdt' || b.toc), 'has table of contents');
  const plan = buildFromOutline('# 기획\n## 배경\n### 문제\n## 실행', { kind: 'plan', placeholders: false });
  assert.deepEqual(heads(plan).map((h) => h[1]), ['Ⅰ. 배경', '1. 문제', 'Ⅱ. 실행']);
  assert.ok(plan.body.find((b) => b.style === 'Heading1' && paraPlain(b) === 'Ⅱ. 실행').pPr.pageBreakBefore);
  const plain = buildFromOutline('## 1. 이미 번호', { kind: 'plain' });
  assert.deepEqual(heads(plain), [['Heading2', '1. 이미 번호']]);
});

test('outlineFromDoc round-trips headings', () => {
  const doc = buildFromOutline('# T\n## 하나\n첫 문장입니다. 둘째 문장.\n## 둘', { kind: 'plain', placeholders: false });
  const md = outlineFromDoc(doc);
  assert.match(md, /^# 하나\n- 첫 문장입니다\.\n\n# 둘$/);
});
