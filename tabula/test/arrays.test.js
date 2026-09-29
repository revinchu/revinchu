import { test } from 'node:test';
import assert from 'node:assert/strict';
import { Workbook } from '../src/workbook.js';
import { readXlsx, writeXlsx } from '../src/xlsx.js';
import { unzip, zip, textOf } from '../src/zip.js';
import { toFileFormula, fromFileFormula } from '../src/xlfn.js';
import { ERR } from '../src/formula.js';

function book(cells) {
  const wb = new Workbook();
  wb.transact(() => {
    for (const [a, raw] of Object.entries(cells)) {
      const m = /^([A-Z]+)(\d+)$/.exec(a);
      const c = [...m[1]].reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0) - 1;
      wb.setInput(0, Number(m[2]) - 1, c, raw);
    }
  });
  return wb;
}
const v = (wb, a, si = 0) => {
  const m = /^([A-Z]+)(\d+)$/.exec(a);
  const c = [...m[1]].reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0) - 1;
  return wb.getValue(si, Number(m[2]) - 1, c);
};

test('동적 배열: 분산 · #SPILL! · A1# 참조', () => {
  const wb = book({ A1: '=SEQUENCE(3)', B1: '=A1#*10', C1: '=SUM(A1#)' });
  assert.equal(v(wb, 'A3'), 3);
  assert.equal(v(wb, 'B2'), 20);
  assert.equal(v(wb, 'C1'), 6);
  assert.deepEqual(wb.spillRange(0, 0, 0), { r1: 0, c1: 0, r2: 2, c2: 0 });
  assert.deepEqual(wb.spillAnchorOf(0, 2, 1), { r: 0, c: 1 });
  wb.transact(() => wb.setInput(0, 1, 0, 'x'));
  assert.equal(v(wb, 'A1'), ERR.SPILL);
  assert.equal(v(wb, 'A2'), 'x');
  assert.equal(v(wb, 'C1'), ERR.REF);
});

test('배열 함수: FILTER · SORT · UNIQUE · XLOOKUP · LET · LAMBDA', () => {
  const wb = book({
    A1: '서울', A2: '부산', A3: '서울', A4: '대구',
    B1: '10', B2: '20', B3: '30', B4: '5',
    D1: '=UNIQUE(A1:A4)', E1: '=SORT(B1:B4,,-1)', F1: '=FILTER(A1:A4,B1:B4>=20)',
    G1: '=XLOOKUP("부산",A1:A4,B1:B4)', H1: '=LET(x,SUM(B1:B4),x/4)', I1: '=MAP(B1:B2,LAMBDA(v,v*2))',
    J1: '=SUMPRODUCT((A1:A4="서울")*B1:B4)', K1: '=TEXTJOIN("-",TRUE,UNIQUE(A1:A4))',
  });
  assert.deepEqual([v(wb, 'D1'), v(wb, 'D2'), v(wb, 'D3')], ['서울', '부산', '대구']);
  assert.deepEqual([v(wb, 'E1'), v(wb, 'E4')], [30, 5]);
  assert.deepEqual([v(wb, 'F1'), v(wb, 'F2')], ['부산', '서울']);
  assert.equal(v(wb, 'G1'), 20);
  assert.equal(v(wb, 'H1'), 16.25);
  assert.equal(v(wb, 'I2'), 40);
  assert.equal(v(wb, 'J1'), 40);
  assert.equal(v(wb, 'K1'), '서울-부산-대구');
});

test('이름 정의: 범위 · 상수 · LAMBDA · 시트 범위', () => {
  const wb = book({ A1: '1', A2: '2', A3: '3', B1: '=SUM(값)', B2: '=세율*100', B3: '=두배(21)', B4: '=ROWS(값)' });
  wb.transact(() => wb.setNames([
    { name: '값', ref: '=Sheet1!$A$1:$A$3' },
    { name: '세율', ref: '=0.1' },
    { name: '두배', ref: '=LAMBDA(x,x*2)' },
  ]));
  assert.equal(v(wb, 'B1'), 6);
  assert.equal(v(wb, 'B2'), 10);
  assert.equal(v(wb, 'B3'), 42);
  assert.equal(v(wb, 'B4'), 3);
  // 행 삽입 시 이름 참조도 이동
  wb.transact(() => wb.insertRows(0, 0, 1));
  assert.equal(wb.names[0].ref, '=Sheet1!$A$2:$A$4');
  assert.equal(v(wb, 'B2'), 6);
});

test('파일 수식 변환: 접두사 · SINGLE · ANCHORARRAY · 옛 형식 암시적 교차', () => {
  assert.equal(toFileFormula('=XLOOKUP(A1,B:B,C:C)'), '_xlfn.XLOOKUP(A1,B:B,C:C)');
  assert.equal(toFileFormula('=SORT(FILTER(A1:A9,B1:B9>0))'), '_xlfn._xlws.SORT(_xlfn._xlws.FILTER(A1:A9,B1:B9>0))');
  assert.equal(toFileFormula('=LET(x,2,x*x)'), '_xlfn.LET(_xlpm.x,2,_xlpm.x*_xlpm.x)');
  assert.equal(toFileFormula('=SUM(A1#)'), 'SUM(_xlfn.ANCHORARRAY(A1))');
  assert.equal(toFileFormula('=@A1:A9', { dynamic: true }), '_xlfn.SINGLE(A1:A9)');
  assert.equal(toFileFormula('=GROUPBY(A1:A9,B1:B9,SUM)'), '_xlfn.GROUPBY(A1:A9,B1:B9,_xleta.SUM)');
  assert.equal(fromFileFormula('_xlfn.LET(_xlpm.x,2,_xlpm.x*_xlpm.x)'), 'LET(x,2,x*x)');
  assert.equal(fromFileFormula('SUM(_xlfn.ANCHORARRAY(A1))'), 'SUM(A1#)');
  assert.equal(fromFileFormula('A1:A10*2', { legacy: true }), '@A1:A10*2');
  assert.equal(fromFileFormula('SUM(A1:A10)', { legacy: true }), 'SUM(A1:A10)');
});

test('xlsx 왕복: 동적 배열 · 이름 · 지원하지 않는 함수의 저장된 값', () => {
  const wb = book({ A1: '=SEQUENCE(2,2)', D1: '=SUM(A1#)', E1: '=금액*2', F1: '5' });
  wb.transact(() => wb.setNames([{ name: '금액', ref: '=Sheet1!$F$1' }]));
  const bytes = writeXlsx(wb);
  const files = unzip(bytes);
  const sheetXml = textOf(files['xl/worksheets/sheet1.xml']);
  assert.match(sheetXml, /<c r="A1" cm="1"><f t="array" ref="A1:B2" aca="false">_xlfn\.SEQUENCE\(2,2\)<\/f><v>1<\/v><\/c>/);
  assert.match(sheetXml, /<c r="B2"><v>4<\/v><\/c>/);
  assert.match(sheetXml, /_xlfn\.ANCHORARRAY\(A1\)/);
  assert.ok(files['xl/metadata.xml']);
  assert.match(textOf(files['xl/workbook.xml']), /<definedName name="금액">Sheet1!\$F\$1<\/definedName>/);

  const { data } = readXlsx(bytes);
  const back = new Workbook(data);
  assert.equal(back.getRaw(0, 0, 0), '=SEQUENCE(2,2)');
  assert.equal(back.getRaw(0, 1, 1), ''); // 분산 값은 다시 계산됨
  assert.equal(back.getValue(0, 1, 1), 4);
  assert.equal(back.getValue(0, 0, 3), 10);
  assert.equal(back.getValue(0, 0, 4), 10);
  assert.equal(back.names[0].name, '금액');

  // 지원하지 않는 함수: 수식 유지 + 파일의 계산 값 표시
  const xml = sheetXml.replace('<c r="F1"><v>5</v></c>', '<c r="F1"><f>_xlfn.NOSUCHFN(1)</f><v>77</v></c>');
  const patched = { ...files, 'xl/worksheets/sheet1.xml': new TextEncoder().encode(xml) };
  const res = readXlsx(zip(patched));
  const wb2 = new Workbook(res.data);
  assert.equal(wb2.getRaw(0, 0, 5), '=NOSUCHFN(1)');
  assert.equal(wb2.getValue(0, 0, 5), 77);
  assert.ok(res.warnings.some((w) => w.includes('저장된 계산 결과')));
});

test('조건부 서식 xlsx: 여러 범위 · 임계값 · 2010 확장 아이콘 · 표준 편차 · 표시 형식', async () => {
  const { prepareCond, condFormatAt } = await import('../src/condfmt.js');
  const wb = book({ A1: '1', A2: '5', A3: '10', C1: '100', C2: '-20' });
  wb.transact(() => wb.setSheetProp(0, 'cond', [
    { r1: 0, c1: 0, r2: 2, c2: 0, more: [{ r1: 0, c1: 2, r2: 1, c2: 2 }], type: 'top', v1: '1', style: { bold: true } },
    { r1: 0, c1: 0, r2: 2, c2: 0, type: 'icons', icons: '3Stars', cfvo: [{ type: 'percent', v: 0 }, { type: 'num', v: 4 }, { type: 'num', v: 9 }] },
    { r1: 0, c1: 2, r2: 1, c2: 2, type: 'bar', color: '#5b9bd5', negColor: '#ff0000', gradient: false, cfvo: [{ type: 'num', v: -50 }, { type: 'num', v: 100 }] },
    { r1: 0, c1: 0, r2: 2, c2: 0, type: 'aboveAvg', stdDev: 1, equal: true, style: { numFmt: 'custom', code: '0.00"점"' } },
  ]));
  const preps = prepareCond(wb, 0);
  assert.equal(condFormatAt(preps, wb, 0, 0, 2, 100).style?.bold, true); // 두 범위를 합쳐 상위 1개
  assert.equal(condFormatAt(preps, wb, 0, 2, 0, 10).style?.bold, undefined);
  assert.equal(condFormatAt(preps, wb, 0, 1, 0, 5).icon, 'star1');
  assert.equal(condFormatAt(preps, wb, 0, 2, 0, 10).icon, 'star2');
  assert.equal(condFormatAt(preps, wb, 0, 1, 2, -20).bar.neg, true);

  const files = unzip(writeXlsx(wb));
  const xml = textOf(files['xl/worksheets/sheet1.xml']);
  assert.match(xml, /sqref="A1:A3 C1:C2"/);
  assert.match(xml, /x14:iconSet iconSet="3Stars"/);
  assert.match(xml, /stdDev="1"/);
  assert.match(textOf(files['xl/styles.xml']), /<dxf><numFmt numFmtId="\d+" formatCode="0\.00&quot;점&quot;"\/>/);
  const back = new Workbook(readXlsx(writeXlsx(wb)).data);
  const cond = back.sheets[0].cond;
  assert.deepEqual(cond[0].more, [{ r1: 0, c1: 2, r2: 1, c2: 2 }]);
  assert.equal(cond[1].icons, '3Stars');
  assert.deepEqual(cond[1].cfvo.map((c) => [c.type, String(c.v)]), [['percent', '0'], ['num', '4'], ['num', '9']]);
  assert.equal(cond[2].negColor, '#ff0000');
  assert.equal(cond[2].gradient, false);
  assert.equal(cond[3].stdDev, 1);
  assert.equal(cond[3].style.code, '0.00"점"');
});

test('표 xlsx: 계산된 열 · 요약 행 사용자 수식', () => {
  const wb = book({ A1: '품목', B1: '수량', C1: '금액', A2: '가', B2: '2', A3: '나', B3: '3', C2: '=[@수량]*10', C3: '=[@수량]*10', C4: '=SUBTOTAL(109,[금액])*2' });
  wb.transact(() => wb.setSheetProp(0, 'tables', [{ id: 't1', name: '표1', r1: 0, c1: 0, r2: 3, c2: 2, header: true, totals: true, style: 'TableStyleMedium2', totalsFns: {} }]));
  const xml = textOf(unzip(writeXlsx(wb))['xl/tables/table1.xml']);
  assert.match(xml, /<calculatedColumnFormula>표1\[\[#This Row\],\[수량\]\]\*10<\/calculatedColumnFormula>/);
  assert.match(xml, /totalsRowFunction="custom">.*<totalsRowFormula>SUBTOTAL\(109,표1\[금액\]\)\*2<\/totalsRowFormula>/);
});

test('복잡한 피벗 테이블 xlsx 왕복: 행 2개 · 열 · 값 2개 · 보고서 필터 · 테이블 형식 · GETPIVOTDATA', async () => {
  const { computePivot, resolvePivot, pivotSourceData } = await import('../src/pivot.js');
  const rows = [['지역', '구', '분기', '매출', '수량', '연도'], ['서울', '강남', 'Q1', '100', '1', '2024'], ['서울', '강북', 'Q1', '50', '2', '2024'],
    ['부산', '해운대', 'Q2', '30', '3', '2024'], ['서울', '강남', 'Q2', '20', '4', '2023'], ['부산', '해운대', 'Q1', '10', '5', '2024']];
  const cells = {};
  rows.forEach((row, r) => row.forEach((v, c) => { cells[`${r},${c}`] = v; }));
  const wb = book({});
  wb.transact(() => Object.entries(cells).forEach(([k, v]) => { const [r, c] = k.split(',').map(Number); wb.setInput(0, r, c, v); }));
  const at = wb.transact(() => wb.addSheet('피벗'));
  const def = {
    source: 'Sheet1', range: { r1: 0, c1: 0, r2: 5, c2: 5 }, rows: ['지역', '구'], cols: ['분기'],
    values: [{ field: '매출', agg: 'sum' }, { field: '수량', agg: 'average', showAs: 'percentOfTotal' }], pages: ['연도'],
    filters: { 연도: ['2024'] }, layout: 'tabular', subtotals: true, top: 0, left: 0,
  };
  const { def: d, rows: vis } = resolvePivot(pivotSourceData(wb, def).rows, def);
  const { grid } = computePivot(vis, d);
  wb.transact(() => { grid.forEach((row, r) => row.forEach((cd, c) => { if (cd?.raw) wb.setCellData(at, r, c, cd); })); wb.setSheetProp(at, 'pivot', { ...def, area: { r1: 0, c1: 0, r2: grid.length - 1, c2: 8 } }); });
  assert.equal(grid[0][1].raw.replace(/^'/, ''), '2024');
  const files = unzip(writeXlsx(wb));
  const pt = textOf(files['xl/pivotTables/pivotTable1.xml']);
  assert.match(pt, /<rowFields count="2"><field x="0"\/><field x="1"\/><\/rowFields>/);
  assert.match(pt, /<colFields count="2"><field x="2"\/><field x="-2"\/><\/colFields>/);
  assert.match(pt, /<pageFields count="1"><pageField fld="5" item="\d" hier="-1"\/><\/pageFields>/);
  assert.match(pt, /showDataAs="percentOfTotal"/);
  assert.match(pt, /outline="0"/);
  const back = readXlsx(zip(files)).data.sheets[1].pivot;
  assert.deepEqual([back.rows, back.cols, back.pages, back.layout], [['지역', '구'], ['분기'], ['연도'], 'tabular']);
  assert.deepEqual(back.values, [{ field: '매출', agg: 'sum' }, { field: '수량', agg: 'average', showAs: 'percentOfTotal' }]);
  assert.deepEqual(back.filters, { 연도: ['2024'] });
  // GETPIVOTDATA
  wb.transact(() => wb.setInput(at, 30, 0, '=GETPIVOTDATA("매출",A3,"지역","서울","분기","Q1")'));
  assert.equal(wb.getValue(at, 30, 0), 150);
});

test('빠른 채우기 · 하이퍼링크 xlsx 왕복', async () => {
  const { flashFill } = await import('../src/flashfill.js');
  assert.deepEqual(flashFill([{ sources: ['홍길동 <gd.hong@abc.com>'], target: 'gd.hong' }], [['김철수 <cs.kim@xyz.com>']]), ['cs.kim']);
  assert.deepEqual(flashFill([{ sources: ['john', 'smith'], target: 'Smith, John' }], [['mary', 'jones']]), ['Jones, Mary']);
  assert.deepEqual(flashFill([{ sources: ['010-1234-5678'], target: '01012345678' }], [['010-9876-5432']]), ['01098765432']);
  const wb = book({ A1: '회사', A2: '목차' });
  wb.transact(() => {
    wb.setCellData(0, 0, 0, { raw: '회사', link: 'https://example.com/a?b=1&c=2' });
    wb.setCellData(0, 1, 0, { raw: '목차', link: '#Sheet1!C5' });
  });
  const bytes = writeXlsx(wb);
  const files = unzip(bytes);
  assert.match(textOf(files['xl/worksheets/sheet1.xml']), /<hyperlinks><hyperlink ref="A1" r:id="rId\d+"\/><hyperlink ref="A2" location="Sheet1!C5"/);
  assert.match(textOf(files['xl/worksheets/_rels/sheet1.xml.rels']), /TargetMode="External"/);
  const back = new Workbook(readXlsx(bytes).data);
  assert.equal(back.getCell(0, 0, 0).link, 'https://example.com/a?b=1&c=2');
  assert.equal(back.getCell(0, 1, 0).link, '#Sheet1!C5');
});

test('피벗 날짜 그룹: 파생 필드(월2 = 일의 월) · 일 그룹 · 날짜 필터 항목 xlsx 왕복', async () => {
  const { computePivot, resolvePivot, pivotSourceData } = await import('../src/pivot.js');
  const wb = book({});
  // 2025-11-20, 2025-12-01, 2025-12-02, 2026-01-15 (날짜 일련번호)
  const data = [[45981, 5], [45992, 7], [45993, 1], [46037, 2]];
  wb.transact(() => { wb.setInput(0, 0, 0, '일'); wb.setInput(0, 0, 1, '노출'); data.forEach(([d, v], i) => { wb.setInput(0, i + 1, 0, String(d)); wb.setInput(0, i + 1, 1, String(v)); }); });
  const def = {
    source: 'Sheet1', range: { r1: 0, c1: 0, r2: 4, c2: 1 }, rows: ['월2', '일'], cols: [], values: [{ field: '노출', agg: 'sum' }], pages: [], filters: {},
    groups: { 월2: { by: 'months', base: '일' }, 일: { by: 'mdays' } }, collapsed: { 월2: ['12월'] }, top: 0, left: 3,
  };
  const res = resolvePivot(pivotSourceData(wb, def), def);
  const { grid } = computePivot(res, res.def);
  const labels = grid.map((r) => r[0]?.raw?.replace(/^'/, '')).filter(Boolean);
  // 엑셀처럼 월 번호 순서 (연도 구분 없음)
  assert.deepEqual(labels.slice(0, 5), ['행 레이블', '1월', '1월15일', '11월', '11월20일']);
  assert.ok(labels.includes('12월') && !labels.includes('12월1일')); // 12월은 축소
  // 그룹 항목으로 거르기 (12월만)
  const fdef = { ...def, rows: ['월2'], filters: { 월2: ['12월'] } };
  const fres = resolvePivot(pivotSourceData(wb, fdef), fdef);
  assert.deepEqual(computePivot(fres, fres.def).grid.map((r) => r.map((c) => c?.raw?.replace(/^'/, ''))).slice(1), [['12월', '8'], ['총합계', '8']]);
  wb.transact(() => wb.setSheetProp(0, 'pivot', { ...def, area: { r1: 0, c1: 3, r2: grid.length - 1, c2: 4 } }));
  const files = unzip(writeXlsx(wb));
  const cache = textOf(files['xl/pivotCache/pivotCacheDefinition1.xml']);
  assert.match(cache, /<cacheField name="월2" numFmtId="0" databaseField="0"><fieldGroup base="0"><rangePr groupBy="months"/);
  assert.match(cache, /<cacheField name="일" numFmtId="14"><sharedItems [^>]*containsDate="1"[^>]*\/><fieldGroup par="2" base="0"><rangePr groupBy="days"/);
  const back = readXlsx(zip(files)).data.sheets[0].pivot;
  assert.deepEqual(back.rows, ['월2', '일']);
  assert.equal(back.groups.월2.by, 'months');
  assert.equal(back.groups.월2.base, '일');
  assert.equal(back.groups.일.by, 'mdays');
  assert.deepEqual(back.collapsed, { 월2: ['12월'] });
});
