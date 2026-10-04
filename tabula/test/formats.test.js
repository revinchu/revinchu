// XLSB 수식 해석 · ODS 읽기/쓰기 · 피벗 'Σ 값' 위치
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { decodeFormula } from '../src/xlsb.js';
import { readOds, writeOds, toOdfFormula, fromOdfFormula } from '../src/ods.js';
import { computePivot } from '../src/pivot.js';
import { Workbook } from '../src/workbook.js';
import { readXlsx, writeXlsx } from '../src/xlsx.js';
import { unzip, textOf } from '../src/zip.js';

/** rgce 바이트 배열 → CellParsedFormula (cce + rgce + cb=0) */
const fmla = (bytes, extra = []) => new Uint8Array([bytes.length, 0, 0, 0, ...bytes, extra.length, 0, 0, 0, ...extra]);
const i32 = (n) => [n & 255, (n >> 8) & 255, (n >> 16) & 255, (n >>> 24) & 255];
const u16 = (n) => [n & 255, (n >> 8) & 255];
const env = { sheets: ['Sheet1', '데이터 시트'], names: [{ name: '_xlfn.XLOOKUP' }, { name: '세율' }], xti: [{ self: true, first: 1, last: 1 }, { self: true, first: -2, last: -2 }], tables: new Map([[1, { name: '표1', cols: ['날짜', '비용'] }]]) };

test('XLSB 수식: 참조 · 연산자 · 함수 · 문자열 · 3D 참조 · 이름 · 미래 함수 · 표 참조', () => {
  // =SUM(A1:B2)*2 → PtgArea(0x25) · PtgInt · PtgMul · PtgFuncVar SUM(4) 순서: area, SUM(1개), int, mul
  const area = [0x25, ...i32(0), ...i32(1), ...u16(0xC000), ...u16(0xC001)];
  const sum = [0x22, 1, ...u16(4)];
  const t1 = decodeFormula(fmla([...area, ...sum, 0x1E, ...u16(2), 0x05]), 0, env, { r: 5, c: 5 });
  assert.equal(t1.text, 'SUM(A1:B2)*2');
  // 절대 참조 $C$3 & "원"
  const ref = [0x24, ...i32(2), ...u16(2)];
  const str = [0x17, ...u16(1), 0xD0, 0xC6];
  assert.equal(decodeFormula(fmla([...ref, ...str, 0x08]), 0, env, { r: 0, c: 0 }).text, '$C$3&"원"');
  // 3D: '데이터 시트'!A1
  assert.equal(decodeFormula(fmla([0x3A, ...u16(0), ...i32(0), ...u16(0xC000)]), 0, env, { r: 0, c: 0 }).text, "'데이터 시트'!A1");
  // 이름 · 미래 함수 (PtgNameX + FuncVar 255)
  assert.equal(decodeFormula(fmla([0x23, ...i32(2)]), 0, env, { r: 0, c: 0 }).text, '세율');
  const xl = [0x39, ...u16(1), ...i32(1), ...ref, ...ref, ...ref, 0x22, 4, ...u16(255)];
  assert.equal(decodeFormula(fmla(xl), 0, env, { r: 0, c: 0 }).text, '_xlfn.XLOOKUP($C$3,$C$3,$C$3)');
  // 표 참조: 표1[[#This Row],[비용]]
  const list = [0x18, 0x19, 0, 0, ...u16(0x41), ...u16(1), 0, 0, ...u16(1), ...u16(1)];
  assert.equal(decodeFormula(fmla(list), 0, env, { r: 3, c: 3 }).text, '표1[[#This Row],[비용]]');
  // 공유 수식의 상대 참조(PtgRefN): 한 행 위
  const refN = [0x2C, ...i32(0xFFFFF), ...u16(0xC000)];
  assert.equal(decodeFormula(fmla(refN), 0, env, { r: 9, c: 2 }).text, 'C9');
  // 모르는 토큰 → null (값만 가져옴)
  assert.equal(decodeFormula(fmla([0x02]), 0, env, { r: 0, c: 0 }).text, null);
});

test('ODS: 수식 변환 · 값 · 서식 · 병합 왕복', () => {
  const f = "=SUM(A1:B2)+'시트 2'!C3*IF(A1>0,\"a,b\",Sheet1!$D$4)";
  assert.equal(toOdfFormula(f), 'of:=SUM([.A1:.B2])+[$\'시트 2\'.C3]*IF([.A1]>0;"a,b";[$Sheet1.$D$4])');
  assert.equal(`=${fromOdfFormula(toOdfFormula(f))}`, f);
  const wb = new Workbook();
  wb.transact(() => {
    [['이름', '값'], ['A', '12.5'], ['', '=B2*2'], ['', '2024-03-15'], ['', '15%']].forEach((row, r) => row.forEach((v, c) => { if (v) wb.setInput(0, r, c, v); }));
    for (const c of [0, 1]) wb.setStyle(0, 0, c, { bold: true, fill: '#ffee00' });
  });
  const bytes = writeOds([{ si: 0, name: 'Sheet1', merges: [{ r1: 5, c1: 0, r2: 5, c2: 1 }] }], {
    raw: (s, r, c) => wb.getRaw(s, r, c), value: (s, r, c) => wb.getValue(s, r, c), style: (s, r, c) => wb.styleAt(s, r, c),
    used: (s) => wb.usedRange(s), colWidth: (s, c) => wb.colWidth(s, c), rowHeight: () => null,
  });
  assert.equal(textOf(unzip(bytes).mimetype), 'application/vnd.oasis.opendocument.spreadsheet');
  const back = new Workbook(readOds(bytes).data);
  assert.equal(back.getRaw(0, 2, 1), '=B2*2');
  assert.equal(back.getValue(0, 2, 1), 25);
  assert.equal(back.getValue(0, 3, 1), 45366);
  assert.equal(back.styleAt(0, 4, 1).numFmt, 'percent');
  assert.equal(back.styleAt(0, 0, 0).bold, true);
  assert.equal(back.styleAt(0, 0, 0).fill, '#ffee00');
  assert.deepEqual(back.sheets[0].merges, [{ r1: 5, c1: 0, r2: 5, c2: 1 }]);
});

test("피벗 'Σ 값' 위치: 맨 안쪽(월별) / 맨 바깥(지표별로 4월·5월 나란히) + xlsx 왕복", () => {
  const rows = [['월', '매체', '비용', '클릭'], ['4월', 'A', 10, 1], ['5월', 'A', 20, 2], ['4월', 'B', 30, 3], ['5월', 'B', 40, 4]];
  const base = { rows: ['매체'], cols: ['월'], pages: [], filters: {}, layout: 'tabular', grandRows: true, grandCols: false, style: 'None', values: [{ field: '비용', agg: 'sum' }, { field: '클릭', agg: 'sum' }] };
  const txt = (grid, r) => grid[r].map((c) => String(c?.raw ?? '').replace(/^'/, ''));
  const inner = computePivot(rows, base).grid;
  assert.deepEqual(txt(inner, 3).slice(1), ['10', '1', '20', '2']);
  const outer = computePivot(rows, { ...base, valuesPos: 0 }).grid;
  assert.deepEqual(txt(outer, 1).slice(1), ['합계 : 비용', '', '합계 : 클릭', '']);
  assert.deepEqual(txt(outer, 2).slice(1), ['4월', '5월', '4월', '5월']);
  assert.deepEqual(txt(outer, 3).slice(1), ['10', '20', '1', '2']);
  // xlsx: colFields 에 -2 가 앞에
  const wb = new Workbook();
  wb.transact(() => rows.forEach((row, r) => row.forEach((v, c) => wb.setInput(0, r, c, String(v)))));
  wb.transact(() => wb.setSheetProp(0, 'pivot', { ...base, name: '피벗1', source: 'Sheet1', range: { r1: 0, c1: 0, r2: 4, c2: 3 }, valuesPos: 0, top: 0, left: 6, area: { r1: 0, c1: 6, r2: 5, c2: 10 } }));
  const xml = textOf(unzip(writeXlsx(wb))['xl/pivotTables/pivotTable1.xml']);
  assert.match(xml, /<colFields count="2"><field x="-2"\/><field x="0"\/><\/colFields>/);
  const back = new Workbook(readXlsx(writeXlsx(wb)).data);
  assert.equal(back.sheets[0].pivot.valuesPos, 0);
});

test('큰 데이터: 고유 숫자 15만 개인 피벗 원본 저장 · MAX(열 전체) 가 호출 스택을 넘기지 않음', () => {
  const wb = new Workbook();
  wb.transact(() => { wb.setInput(0, 0, 0, '매체'); wb.setInput(0, 0, 1, '비용'); for (let i = 1; i <= 150000; i++) { wb.setInput(0, i, 0, i % 3 ? 'A' : 'B'); wb.setInput(0, i, 1, String(i + 0.5)); } });
  wb.transact(() => wb.setSheetProp(0, 'pivot', { name: '피벗1', source: 'Sheet1', range: { r1: 0, c1: 0, r2: 150000, c2: 1 }, rows: ['매체'], cols: [], pages: [], filters: {}, values: [{ field: '비용', agg: 'max' }], top: 0, left: 4, area: { r1: 0, c1: 4, r2: 3, c2: 5 } }));
  wb.transact(() => wb.setInput(0, 0, 8, '=MAX(B:B)-MIN(B:B)'));
  assert.equal(wb.getValue(0, 0, 8), 149999);
  const cacheXml = textOf(unzip(writeXlsx(wb))['xl/pivotCache/pivotCacheDefinition1.xml']);
  assert.match(cacheXml, /containsNumber="1"/);
  assert.doesNotMatch(cacheXml, /maxValue=/); // 나열되지 않은 값 필드에는 숫자 자식이 없음
});

test('.xls (엑셀 97-2003) 읽기: 시트 · 값 · 서식 · 수식(공유 · 시트 참조 · 이름 · 추가 기능 함수 · 배열 상수)', async () => {
  const { readFileSync } = await import('node:fs');
  const { readXls } = await import('../src/xls.js');
  const { Workbook } = await import('../src/workbook.js');
  const res = readXls(new Uint8Array(readFileSync(new URL('./fixtures/formulas97.xls', import.meta.url))));
  assert.deepEqual(res.warnings, []);
  const wb = new Workbook(res.data);
  assert.deepEqual(wb.sheets.map((s) => s.name), ['Sheet1', '데이터 2']);
  assert.deepEqual(wb.names.map((n) => [n.name, n.ref]), [['단가', '=Sheet1!$B$1:$B$10']]);
  const f = (r, c) => wb.getCell(0, r, c).raw;
  assert.equal(f(4, 2), '=A5*B5+$A$1'); // 공유 수식의 상대 참조
  assert.equal(f(3, 5), "=SUMIF('데이터 2'!A1:A30,\"항목1\",'데이터 2'!B1:B30)");
  assert.equal(f(11, 5), '=SUM({1,2;3,4})');
  assert.equal(f(16, 5), '=EDATE(DATE(2024,1,31),1)');
  assert.equal(f(12, 5), '=IFERROR(1/0,"오류")');
  assert.equal(f(25, 5), '=A1:A3 A2:B2');
  const cached = [];
  wb.sheets.forEach((s, si) => s.cells.forEachRC((c, r, cc) => { if (c.formula) cached.push([si, r, cc, wb.getValue(si, r, cc)]); }));
  wb.invalidate();
  for (const [si, r, c, v] of cached) {
    const x = wb.getValue(si, r, c);
    assert.ok(JSON.stringify(x) === JSON.stringify(v) || Math.abs(x - v) < 1e-9, `${r},${c} ${JSON.stringify(x)} ≠ ${JSON.stringify(v)}`);
  }
});

/** 작은 EMF 만들기: 머리말 + 붓 · 사각형 + 글꼴 · 글자 + 끝 */
function tinyEmf() {
  const recs = [];
  const rec = (type, body) => { const len = 8 + body.length; recs.push([...i32(type), ...i32(len), ...body]); };
  const header = [...i32(0), ...i32(0), ...i32(99), ...i32(49), ...i32(0), ...i32(0), ...i32(2646), ...i32(1323),
    0x20, 0x45, 0x4d, 0x46, ...i32(0x10000), ...i32(0), ...i32(0), ...u16(0), ...u16(0), ...i32(0), ...i32(0), ...i32(0),
    ...i32(1920), ...i32(1080), ...i32(508), ...i32(285)];
  rec(1, header);
  rec(39, [...i32(1), ...i32(0), 0x44, 0x72, 0xc4, 0, ...i32(0)]); // 붓 #4472c4
  rec(37, i32(1));
  rec(37, i32(0x80000008)); // NULL_PEN
  rec(43, [...i32(10), ...i32(10), ...i32(60), ...i32(40)]);
  const face = [...'맑은 고딕'].flatMap((ch) => u16(ch.charCodeAt(0)));
  rec(82, [...i32(2), ...i32(-12), ...i32(0), ...i32(0), ...i32(0), ...i32(700), 0, 0, 0, 129, 0, 0, 0, 0, ...face, ...new Array(64 - face.length).fill(0)]);
  rec(37, i32(2));
  rec(24, [0xff, 0, 0, 0]); // 글자색 빨강
  const str = [...'매출'].flatMap((ch) => u16(ch.charCodeAt(0)));
  // EMREXTTEXTOUTW: bounds 16 + mode/scale 12 + EMRTEXT 40 (+8 머리) = 76 바이트 뒤에 문자열
  rec(84, [...i32(0), ...i32(0), ...i32(0), ...i32(0), ...i32(1), ...i32(0), ...i32(0),
    ...i32(5), ...i32(45), ...i32(2), ...i32(76), ...i32(0), ...i32(0), ...i32(0), ...i32(-1), ...i32(-1), ...i32(0), ...str]);
  rec(14, [...i32(0), ...i32(0), ...i32(20)]);
  const all = recs.flat();
  all.splice(48, 4, ...i32(all.length)); // nBytes
  all.splice(52, 4, ...i32(recs.length));
  return new Uint8Array(all);
}

for (const ext of ['emf', 'bin']) test(`EMF 그림(${ext})을 SVG 로 그리고 xlsx 로 다시 저장할 때 원본 EMF 를 유지`, async () => {
  const { emfToSvg } = await import('../src/emf.js');
  const { zip } = await import('../src/zip.js');
  const emf = tinyEmf();
  const svg = emfToSvg(emf);
  assert.match(svg, /viewBox="0 0 100 50"/);
  assert.match(svg, /<path d="M10 10H60V40H10Z" fill="#4472c4" stroke="none"/);
  assert.match(svg, /<text x="5" y="45"[^>]*font-size="12" font-weight="bold" fill="#ff0000"[^>]*>매출<\/text>/);
  assert.equal(emfToSvg(new Uint8Array(100)), null);

  // xlsx 안의 EMF → 화면용 SVG + 원본(im.emf), 저장하면 다시 .emf
  const wb = new Workbook();
  wb.setInput(0, 0, 0, '1');
  const png = 'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8DwHwAFBQIAX8jx0gAAAABJRU5ErkJggg==';
  wb.sheets[0].images.push({ id: 'im1', name: '차트', x: 10, y: 10, w: 100, h: 50, src: `data:image/png;base64,${png}` });
  const files = unzip(writeXlsx(wb));
  const media = Object.keys(files).find((k) => k.startsWith('xl/media/'));
  delete files[media];
  files[`xl/media/image1.${ext}`] = emf;
  for (const k of Object.keys(files)) {
    if (/drawing\d+\.xml\.rels$/.test(k)) files[k] = new TextEncoder().encode(textOf(files[k]).replace(/image1\.png/, `image1.${ext}`));
    if (k === '[Content_Types].xml') files[k] = new TextEncoder().encode(textOf(files[k]).replace('Extension="png" ContentType="image/png"', `Extension="${ext}" ContentType="${ext === 'emf' ? 'image/x-emf' : 'image/unknown'}"`));
  }
  const { data, warnings } = readXlsx(zip(files));
  assert.deepEqual(warnings, []);
  const im = new Workbook(data).sheets[0].images[0];
  assert.match(im.src, /^data:image\/svg\+xml;base64,/);
  assert.match(im.emf, /^data:image\/x-emf;base64,/);
  const out = unzip(writeXlsx(new Workbook(data)));
  assert.deepEqual([...out['xl/media/image1.emf']], [...emf]);
  assert.match(textOf(out['[Content_Types].xml']), /Extension="emf" ContentType="image\/x-emf"/);
});
