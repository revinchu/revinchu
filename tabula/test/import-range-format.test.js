import test from 'node:test';
import assert from 'node:assert/strict';
import { Workbook } from '../src/workbook.js';
import { Range } from '../src/fxcore.js';
import { NET, importRangeSource, netClear } from '../src/fx-web.js';
import { formatValue } from '../src/format.js';
import { readXlsx, writeXlsx } from '../src/xlsx.js';
import { unzip, textOf } from '../src/zip.js';
import { parseXml, child, kids } from '../src/xml.js';

// Synthetic responses only: no real document ID, URL, downloaded rows, or network requests.
const ID = 'synthetic_format_fixture_12345';
const csv = '금액,비율,상태\n"12,345.60",12.50%,유지\n"1,234.00",0.75%,확인';
const formula = `=IMPORTRANGE("${ID}","A1:C3")`;
const display = (wb, r, c) => formatValue(wb.getValue(0, r, c), wb.styleAt(0, r, c)).text;
function withImport(fn, text = csv, range = 'A1:C3') {
  const old = { ...NET };
  NET.cache = new Map(); NET.authorize = null; NET.fetcher = () => { throw Error('검사 중 네트워크 요청 금지'); }; NET.onDone = null;
  const url = importRangeSource(ID, range).url;
  const setResponse = value => NET.cache.set(url, { state: 'ok', data: value, sheets: new Set(), t: Date.now() });
  setResponse(text);
  try {
    const wb = new Workbook();
    wb.transact(() => wb.setInput(0, 0, 0, `=IMPORTRANGE("${ID}","${range}")`));
    return fn(wb, setResponse);
  } finally { Object.assign(NET, old); }
}

test('IMPORTRANGE 숫자 표시 힌트는 숫자값·SUM·셀 참조를 문자열로 바꾸지 않는다', () => withImport(wb => {
  assert.equal(wb.getValue(0, 1, 0), 12345.6);
  assert.equal(wb.getValue(0, 1, 1), 0.125);
  assert.equal(display(wb, 1, 0), '12,345.60');
  assert.equal(display(wb, 1, 1), '12.50%');
  assert.equal(display(wb, 2, 0), '1,234.00');
  assert.equal(display(wb, 2, 1), '0.75%');
  const sp = wb.spillsOf(0)[0];
  assert.equal(sp.cellFormats[0]?.[0] ?? null, null);
  assert.equal(sp.cellFormats[1][0].numFmt, 'custom');
  assert.equal(wb.styleAt(0, 1, 0).shrink, true);
  assert.equal(wb.sheets[0].cells.getRC(1, 0), undefined, '표시 힌트를 셀 자료로 굳히지 않는다');
  wb.transact(() => {
    wb.setInput(0, 0, 4, '=SUM(A2:A3)');
    wb.setInput(0, 1, 4, '=B2*100');
    wb.setInput(0, 2, 4, `=QUERY(IMPORTRANGE("${ID}","A1:C3"),"select sum(Col1) label sum(Col1) ''",1)`);
  });
  assert.equal(wb.getValue(0, 0, 4), 13579.6);
  assert.equal(wb.getValue(0, 1, 4), 12.5);
  assert.equal(wb.getValue(0, 2, 4), 13579.6);
}));

test('셀·행·열·시트에 명시한 숫자 서식과 일반 서식은 가져오기 힌트보다 우선한다', () => {
  for (const scope of ['cell', 'row', 'col', 'all']) withImport(wb => {
    wb.getValue(0, 1, 0);
    const patch = { numFmt: 'number', decimals: 1 };
    wb.transact(() => scope === 'cell' ? wb.setStyle(0, 1, 0, patch) : wb.setLineStyle(0, scope, scope === 'row' ? 1 : 0, patch));
    assert.equal(wb.styleAt(0, 1, 0).numFmt, 'number', scope);
    assert.equal(display(wb, 1, 0), '12,345.6', scope);
    wb.transact(() => scope === 'cell' ? wb.setStyle(0, 1, 0, { numFmt: 'general', code: undefined }) : wb.setLineStyle(0, scope, scope === 'row' ? 1 : 0, { numFmt: 'general', code: undefined }));
    assert.equal(wb.styleAt(0, 1, 0).numFmt, 'general', scope);
    assert.equal(display(wb, 1, 0), '12345.6', scope);
    assert.equal(wb.getValue(0, 1, 0), 12345.6, scope);
    wb.undo(); assert.equal(display(wb, 1, 0), '12,345.6', scope);
    wb.redo(); assert.equal(display(wb, 1, 0), '12345.6', scope);
  });
});

test('채우기 등 다른 서식은 힌트와 공존하고 QUERY format 기존 우선순위는 유지한다', () => withImport(wb => {
  wb.transact(() => {
    wb.setStyle(0, 1, 0, { fill: '#123456', bold: true });
    wb.setInput(0, 0, 4, '=QUERY(A1:B3,"select A,B format A \'0.0\',B \'0.0%\'",1)');
    wb.setStyle(0, 1, 4, { numFmt: 'general' });
  });
  assert.equal(display(wb, 1, 0), '12,345.60');
  assert.equal(wb.styleAt(0, 1, 0).fill, '#123456');
  assert.equal(wb.styleAt(0, 1, 0).bold, true);
  wb.transact(() => wb.setStyle(0, 1, 0, { shrink: false }));
  assert.equal(wb.styleAt(0, 1, 0).shrink, false, '사용자가 줄임을 끄면 힌트가 덮어쓰지 않는다');
  assert.equal(display(wb, 1, 0), '12,345.60');
  assert.equal(wb.getValue(0, 1, 4), 12345.6);
  assert.equal(display(wb, 1, 4), '12345.6');
  assert.equal(wb.styleAt(0, 1, 4).queryFormat, '0.0');
  assert.equal(display(wb, 1, 5), '12.5%');
}));

test('numFmt 없이 소수 자릿수나 코드를 지정해도 셀·행·열·시트의 선택을 보존한다', () => {
  for (const scope of ['cell', 'row', 'col', 'all']) for (const patch of [{ decimals: 0 }, { code: '0.000' }]) withImport(wb => {
    wb.getValue(0, 1, 0);
    wb.transact(() => scope === 'cell' ? wb.setStyle(0, 1, 0, patch) : wb.setLineStyle(0, scope, scope === 'row' ? 1 : 0, patch));
    const style = wb.styleAt(0, 1, 0);
    assert.equal(style.numFmt, undefined, scope);
    assert.equal(style.decimals, patch.decimals, scope);
    assert.equal(style.code, patch.code, scope);
    assert.equal(style.shrink, undefined, scope);
    assert.equal(wb.getValue(0, 1, 0), 12345.6);
  });
});

test('다시 가져온 짧은 결과와 수식 지우기에는 이전 스필 표시 힌트가 남지 않는다', () => withImport((wb, setResponse) => {
  assert.equal(display(wb, 2, 0), '1,234.00');
  const affected = netClear();
  assert.deepEqual(affected, [0]);
  setResponse('7,8,완료');
  for (const si of affected) wb.invalidate(si);
  assert.equal(wb.getValue(0, 0, 0), 7);
  assert.equal(display(wb, 0, 0), '7');
  assert.equal(wb.getValue(0, 2, 0), null);
  assert.equal(wb.styleAt(0, 2, 0).code, undefined);
  wb.transact(() => wb.clearRange(0, 0, 0, 0, 0, 'contents'));
  assert.equal(wb.getValue(0, 0, 1), null);
  assert.equal(wb.styleAt(0, 0, 0).code, undefined);
  assert.equal(wb.spillsOf(0).length, 0);
}, csv, 'A:C'));

test('1셀 범위의 숫자 표시 힌트도 스칼라 축소로 잃지 않는다', () => withImport(wb => {
  assert.equal(wb.getValue(0, 0, 0), 1234.5);
  assert.equal(display(wb, 0, 0), '1,234.50');
  assert.equal(wb.spillsOf(0)[0].cellFormats[0][0].numFmt, 'custom');
}, '"1,234.50"', 'A1'));

test('불완전 메타데이터와 텍스트 값에는 숫자 표시 힌트를 강제하지 않는다', () => {
  const wb = new Workbook();
  const arr = new Range([['00123', 0.125, 20]]);
  arr.cellFormats = [[{ numFmt: 'custom', code: '#,##0.00' }, { numFmt: 'custom', code: '0.00%' }, null]];
  wb.placeSpill('0:0,0', 0, 0, 0, arr);
  assert.equal(wb.styleAt(0, 0, 0).code, undefined);
  assert.equal(wb.styleAt(0, 0, 1).code, '0.00%');
  assert.equal(wb.styleAt(0, 0, 2).code, undefined);
  assert.equal(wb.styleAt(0, 0, 3).code, undefined);
  assert.equal(wb.styleAt(0, 1, 1).code, undefined);
  assert.equal(arr.rows[0][0], '00123');
});

test('XLSX 표준 숫자 유형·표시 서식으로 저장하고 다시 열어 값과 모양을 보존한다', () => withImport(wb => {
  wb.getValue(0, 0, 0);
  wb.transact(() => {
    wb.setStyle(0, 2, 1, { numFmt: 'general' });
    wb.setStyle(0, 2, 0, { shrink: false });
  });
  const source = wb.serialize();
  const bytes = writeXlsx(wb), files = unzip(bytes);
  const xml = textOf(files['xl/worksheets/sheet1.xml']);
  assert.match(xml, /<c\b[^>]*\br="A2"[^>]*><v>12345\.6<\/v><\/c>/);
  assert.match(xml, /<c\b[^>]*\br="B2"[^>]*><v>0\.125<\/v><\/c>/);
  const back = new Workbook(readXlsx(bytes).data);
  assert.equal(back.getRaw(0, 0, 0), formula);
  assert.equal(back.getValue(0, 1, 0), 12345.6);
  assert.equal(back.getValue(0, 1, 1), 0.125);
  assert.equal(display(back, 1, 0), '12,345.60');
  assert.equal(display(back, 1, 1), '12.50%');
  assert.equal(display(back, 2, 0), '1,234.00');
  assert.equal(display(back, 2, 1), '0.0075');
  assert.equal(back.styleAt(0, 2, 0).shrink, false);
  assert.deepEqual(wb.serialize(), source);
}));

test('XLSX는 암시적 일반을 명시 General로 바꾸지 않고 사용자 General·shrink false는 보존한다', () => {
  const wb = new Workbook();
  wb.transact(() => {
    for (let c = 0; c < 4; c++) wb.setInput(0, 0, c, '1234');
    wb.setStyle(0, 0, 0, { fill: '#123456' });
    wb.setStyle(0, 0, 1, { fill: '#123456', numFmt: 'general' });
    wb.setStyle(0, 0, 2, { numFmt: '', bold: true });
    wb.setStyle(0, 0, 3, { shrink: false });
  });
  const files = unzip(writeXlsx(wb)), back = new Workbook(readXlsx(writeXlsx(wb)).data);
  const xml = parseXml(textOf(files['xl/styles.xml']));
  const xfs = kids(child(xml, 'cellXfs'), 'xf');
  const sheetXml = parseXml(textOf(files['xl/worksheets/sheet1.xml']));
  const cells = kids(child(child(sheetXml, 'sheetData'), 'row'), 'c');
  const xf = ref => xfs[Number(cells.find(c => c.attrs.r === ref).attrs.s ?? 0)];
  assert.equal(xf('A1').attrs.applyNumberFormat, undefined);
  assert.equal(xf('B1').attrs.applyNumberFormat, '1');
  assert.equal(xf('C1').attrs.applyNumberFormat, undefined);
  assert.equal(back.styleAt(0, 0, 0).numFmt, undefined);
  assert.equal(back.styleAt(0, 0, 1).numFmt, 'general');
  assert.equal(back.styleAt(0, 0, 2).numFmt, undefined);
  assert.equal(child(xf('D1'), 'alignment').attrs.shrinkToFit, '0');
  assert.equal(back.styleAt(0, 0, 3).shrink, false);
});

test('이름 스타일의 명시 General 상속·숫자 구성요소 제외와 직접 General 재지정을 구분한다', () => {
  const wb = new Workbook();
  wb.cellStyles = [
    { name: '일반 선택', style: { numFmt: 'general', fill: '#123456' } },
    { name: '색만 선택', style: { fill: '#234567' }, include: { number: false } },
    { name: '기본 숫자', style: { numFmt: 'number', decimals: 2 } },
    { name: '암시적 일반', style: { fill: '#345678' } },
  ];
  const styles = [
    { ...wb.cellStyles[0].style, cellStyleName: '일반 선택' },
    { ...wb.cellStyles[1].style, cellStyleName: '색만 선택', numFmt: 'general' },
    { cellStyleName: '기본 숫자', numFmt: 'general' },
    { ...wb.cellStyles[3].style, cellStyleName: '암시적 일반' },
  ];
  wb.transact(() => styles.forEach((st, c) => { wb.setInput(0, 0, c, '1'); wb.setStyle(0, 0, c, st); }));
  const bytes = writeXlsx(wb), back = new Workbook(readXlsx(bytes).data);
  const xml = parseXml(textOf(unzip(bytes)['xl/styles.xml'])), xfs = kids(child(xml, 'cellXfs'), 'xf');
  assert.equal(xfs.find(x => x.attrs.xfId === '1').attrs.applyNumberFormat, '0', '이름 스타일의 명시 General도 상속은 유지');
  assert.equal(xfs.find(x => x.attrs.xfId === '2').attrs.applyNumberFormat, '1', '제외된 구성요소는 직접 General을 적용');
  assert.equal(xfs.find(x => x.attrs.xfId === '3').attrs.applyNumberFormat, '1', '숫자 부모를 직접 General로 덮어씀');
  for (let c = 0; c < 3; c++) assert.equal(back.styleAt(0, 0, c).numFmt, 'general');
  assert.equal(back.styleAt(0, 0, 3).numFmt, undefined);
  assert.equal(back.cellStyles[0].style.numFmt, 'general');
  assert.equal(back.cellStyles[1].include.number, false);
  assert.equal(back.cellStyles[3].style.numFmt, undefined);
});
