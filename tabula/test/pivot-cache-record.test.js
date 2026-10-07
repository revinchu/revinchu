import test from 'node:test';
import assert from 'node:assert/strict';
import { parseXml, unx } from '../src/xml.js';
import { scanXmlChildren } from '../src/xml-stream.js';
import { scanPivotCacheRecord, readPivotSnapshotXml, pivotCacheDate, pivotSnapshotValue } from '../src/pivot-cache-data.js';

const enc = new TextEncoder();
const field = i => ({ name: 'Field' + i, db: true, shared: [null, '', 0, true, false, { error: '#N/A' }, 'shared &\n'] });
const fields = Array.from({ length: 59 }, (_, i) => field(i));
// This is the pre-change parseXml conversion, independent of the fast scanner.
function oracle(xml, db, date1904 = false) {
  const node = parseXml(xml), row = new Array(db.length).fill(null);
  for (let j = 0; j < Math.min(node.children.length, db.length); j++) {
    const cell = node.children[j], value = cell.attrs.v;
    row[j] = cell.name === 'x' ? db[j].shared[Number(value)] ?? null : cell.name === 'n' ? Number(value) : cell.name === 'd' ? pivotCacheDate(value, date1904) : cell.name === 'b' ? value === '1' || value === 'true' : cell.name === 'm' ? null : cell.name === 'e' ? { error: value ?? '#N/A' } : unx(value ?? '');
  }
  return row;
}
function drain(gen) { const progress = []; for (;;) { const step = gen.next(); if (step.done) return { snapshot: step.value, progress }; progress.push(step.value); } }
function rowsOf(snapshot) { return Array.from({ length: snapshot.n }, (_, r) => snapshot.header.map((_, c) => pivotSnapshotValue(snapshot, r, c))); }
function* chunks(xml, size) { const bytes = enc.encode(xml); for (let at = 0; at < bytes.length; at += size) yield bytes.subarray(at, at + size); }
const container = records => '<p:pivotCacheRecords xmlns:p="urn:test">' + records.join('') + '</p:pivotCacheRecords>';

for (const date1904 of [false, true]) test('fast cache record handles every typed cell and shared-item type, date1904=' + date1904, () => {
  const xml = '<r><x v="0"/><x v="1"/><x v="2"/><x v="3"/><x v="4"/><x v="5"/><x v="6"/><n v="-2.5e2"/><b v="true"/><b v="0"/><m/><s v=""/><s v="한글😀"/><e v="#DIV/0!"/><d v="1904-01-01T12:34:56.123"/></r>';
  const db = fields.slice(0, 15), row = scanPivotCacheRecord(xml, db, date1904);
  assert.ok(row); assert.deepEqual(row, oracle(xml, db, date1904));
  assert.deepEqual(row.slice(0, 7), db[0].shared); assert.equal(row[7], -250); assert.equal(row[8], true); assert.equal(row[9], false);
  assert.equal(row[10], null); assert.equal(row[11], ''); assert.equal(row[12], '한글😀'); assert.deepEqual(row[13], { error: '#DIV/0!' });
  assert.ok(Math.abs(row[14] - (date1904 ? 0 : 1462) - (12 * 3600 + 34 * 60 + 56.123) / 86400) < 1e-10);
});

test('prefix names, both quote forms, XML whitespace and namespace declarations stay on the fast path', () => {
  const records = [
    "<r><n v='1'/><s v='single'/></r>",
    '<p:r xmlns:p="urn:test"><p:n v = "1" /><other:s v="text" /></p:r>',
    " \t\r\n<_cache:r xmlns='urn:default' xmlns:_cache='urn:test'>\r\n<p-1:n\tv\r\n= '1'\t/>\n<s v='2 > 1'/>\t</_cache:r >\n",
    '<r/>', '<r />', '<p:r xmlns:p="urn:test" />', '<r> \r\n\t </r>',
  ];
  for (const xml of records) { const row = scanPivotCacheRecord(xml, fields.slice(0, 3)); assert.ok(row, xml); assert.deepEqual(row, oracle(xml, fields.slice(0, 3))); }
});

test('entities, XML line normalization and Excel escapes match the old attributes decoder', () => {
  const xml = '<r><s v="&lt;&gt;&amp;&quot;&apos;&#13;&#xA;_x000A_"/><s v="literal\r\nline\rbreak_x005F_x000A_"/><b v="&#49;"/><n v="&#x32;"/><e v="&#35;N/A"/><x v="&#54;"/></r>';
  const db = fields.slice(0, 6), row = scanPivotCacheRecord(xml, db);
  assert.ok(row); assert.deepEqual(row, oracle(xml, db));
  assert.equal(row[0], '<>&"\'\r\n\n'); assert.equal(row[1], 'literal\nline\nbreak_x000A_'); assert.equal(row[2], true); assert.equal(row[3], 2);
  assert.deepEqual(row[4], { error: '#N/A' }); assert.equal(row[5], 'shared &\n');
  const unknown = '<r><s v="&unknown; &AMP;"/></r>'; assert.deepEqual(scanPivotCacheRecord(unknown, [field(0)]), oracle(unknown, [field(0)]));
});

test('missing attrs, invalid shared indices and extra/fewer cells keep database field positions', () => {
  const xml = '<r><m/><x/><x v="-1"/><x v="999"/><s/><e/><b/><d/><n v=""/><n v="12"/></r>';
  const db = fields.slice(0, 11), row = scanPivotCacheRecord(xml, db);
  assert.deepEqual(row, [null, null, null, null, '', { error: '#N/A' }, false, '', 0, 12, null]);
  assert.deepEqual(row, oracle(xml, db)); assert.deepEqual(scanPivotCacheRecord(xml, []), []);
  assert.deepEqual(scanPivotCacheRecord('<r><n v="1"/><n v="2"/></r>', [field(0)]), [1]);
});

test('comments, CDATA, PI, nested/non-self-closing cells and unfamiliar attrs use whole-record fallback', () => {
  const records = [
    '<r><!--before--><n v="3"/><!--after--></r>',
    '<r><![CDATA[ignored]]><n v="3"/></r>',
    '<r><?custom ignored?><n v="3"/></r>',
    '<r><n v="3"></n><s v="tail"/></r>',
    '<r><n v="3"><extLst><ext uri="custom"/></extLst></n><s v="tail"/></r>',
    '<r><s v="value"><![CDATA[ignored]]></s><n v="3"/></r>',
    '<r><unknown v="_x0041_"><n v="99"/></unknown><n v="3"/></r>',
    '<r><n v="3" custom="ignored"/><s v="tail"/></r>',
    '<r custom="ignored"><n v="3"/><s v="tail"/></r>',
    '<r><p:n xmlns:p="urn:test" v="3"/></r>',
    '<r>non-whitespace text<n v="3"/></r>',
    '<r><n v="3"/><unknown v="extra"/></r>',
  ];
  for (const xml of records) {
    assert.equal(scanPivotCacheRecord(xml, fields.slice(0, 2)), null, xml);
    const { snapshot } = drain(readPivotSnapshotXml(enc.encode(container([xml])), fields.slice(0, 2)));
    assert.deepEqual(rowsOf(snapshot), [oracle(xml, fields.slice(0, 2))], xml);
  }
});

test('malformed/unfamiliar complete records preserve the legacy parse result rather than partial fast values', () => {
  const records = [
    '<r><n v="1" v="2"/><s v="end"/></r>',
    '<r><n v=3/><s v="end"/></r>',
    '<r><s v="raw < text"/><s v="end"/></r>',
    '<r><n v="3" junk/><s v="end"/></r>',
    '<r><n v="3"/><broken text/><s v="end"/></r>',
    '<r><a:b:n v="3"/><s v="end"/></r>',
    '<r><n v="3"/ > <s v="end"/></r>',
  ];
  for (const xml of records) {
    assert.equal(scanPivotCacheRecord(xml, fields.slice(0, 2)), null, xml);
    const { snapshot } = drain(readPivotSnapshotXml(enc.encode(container([xml])), fields.slice(0, 2)));
    assert.deepEqual(rowsOf(snapshot), [oracle(xml, fields.slice(0, 2))], xml);
  }
  for (const xml of ['<r><n v="1"/></wrong>', '<r><n v="1"/></r>junk', '<r><n v="1"/>', '<r><n v="1"/></r><r/>']) assert.equal(scanPivotCacheRecord(xml, [field(0)]), null);
});

test('59-field type/quote/prefix combinations agree with the XML tree oracle', () => {
  let seed = 78123; const next = () => { seed ^= seed << 13; seed ^= seed >>> 17; seed ^= seed << 5; return seed >>> 0; };
  const attrs = [['n', '0'], ['n', '-19.25'], ['m', null], ['s', ''], ['s', '한글😀 &amp; _x000A_'], ['x', '0'], ['x', '5'], ['b', 'false'], ['b', 'true'], ['e', '#REF!'], ['d', '2020-10-01T00:00:00']];
  for (let r = 0; r < 120; r++) {
    const prefix = r % 2 ? 'p:' : '', quote = r % 3 ? '"' : "'"; let body = '';
    for (let c = 0; c < 59; c++) { const [type, v] = attrs[next() % attrs.length]; body += '<' + prefix + type + (v == null ? '' : ' v=' + quote + v + quote) + '/>' + (c % 7 ? '' : '\n'); }
    const xml = '<' + prefix + 'r>' + body + '</' + prefix + 'r>', row = scanPivotCacheRecord(xml, fields, !!(r % 2));
    assert.ok(row); assert.deepEqual(row, oracle(xml, fields, !!(r % 2)));
  }
});

test('streamed UTF8 records preserve field order, db filtering, row count and progress at tiny byte boundaries', () => {
  const input = [field(0), { ...field(99), db: false }, field(1)];
  const record = '<p:r><p:s v="한글😀 &amp; _x000A_"/><p:n v="0"/></p:r>';
  const xml = container(Array.from({ length: 2050 }, () => record));
  for (const size of [1, 17, 65536]) {
    const { snapshot, progress } = drain(readPivotSnapshotXml(chunks(xml, size), input, false, 2050));
    assert.equal(snapshot.n, 2050); assert.deepEqual(snapshot.header, ['Field0', 'Field1']); assert.deepEqual(progress, [2048]);
    for (const r of [0, 2048, 2049]) assert.deepEqual(snapshot.header.map((_, c) => pivotSnapshotValue(snapshot, r, c)), ['한글😀 & \n', 0]);
  }
  assert.throws(() => drain(readPivotSnapshotXml(chunks(xml, 17), input, false, 2051)), /행 수/);
  assert.throws(() => [...scanXmlChildren(chunks('<pivotCacheRecords><r><s v="incomplete"/>', 7), 'r')], /끝까지/);
});

test('finite-number and entity failures are not swallowed by fast or fallback record parsing', () => {
  for (const value of ['Infinity', '-Infinity', 'NaN', 'not-number']) for (const comment of ['', '<!--fallback-->']) {
    const xml = '<r>' + comment + '<n v="' + value + '"/></r>';
    assert.throws(() => drain(readPivotSnapshotXml(enc.encode(container([xml])), [field(0)])), /올바르지 않은 숫자/);
  }
  assert.throws(() => drain(readPivotSnapshotXml(enc.encode(container(['<r><n/></r>'])), [field(0)])), /올바르지 않은 숫자/);
  for (const comment of ['', '<!--fallback-->']) {
    const xml = '<r>' + comment + '<s v="&#x110000;"/></r>';
    assert.throws(() => oracle(xml, [field(0)]), RangeError);
    assert.throws(() => drain(readPivotSnapshotXml(enc.encode(container([xml])), [field(0)])), RangeError);
  }
  // Even ignored extra cells must still validate attribute entities like parseXml.
  assert.throws(() => scanPivotCacheRecord('<r><m/><s v="&#x110000;"/></r>', [field(0)]), RangeError);
});
