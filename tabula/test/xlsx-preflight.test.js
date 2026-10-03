import { test } from 'node:test';
import assert from 'node:assert/strict';
import { Workbook } from '../src/workbook.js';
import { ColBuilder } from '../src/block.js';
import { EXCEL_MAX_ROWS, MAX_COLS } from '../src/formula.js';
import { readXlsx, writeXlsx, xlsxOverflow, xlsxExportWarnings } from '../src/xlsx.js';
import { zip, unzip, textOf } from '../src/zip.js';
import { parseXml, descendants, child } from '../src/xml.js';

const R = EXCEL_MAX_ROWS, C = MAX_COLS;
function block(r0, c0, columns, perm) {
  const n = columns[0].length;
  return { r0, c0, n, ver: 0, cols: columns.map(values => { const col = new ColBuilder(); values.forEach((v, i) => col.set(i, v)); return col.finish(n); }), ...(perm ? { perm: Int32Array.from(perm) } : {}) };
}
const book = (blocks = [], cells = {}) => new Workbook({ sheets: [{ name: '합성', cells, blocks }] });
const sheetXml = wb => textOf(unzip(writeXlsx(wb))['xl/worksheets/sheet1.xml']);

test('XLSX 경계: 마지막 행·열은 유지하고 다음 행·열만 실제 값 수에 포함', () => {
  const wb = book([block(R - 1, C - 1, [[11, 12], [21, 22]])]);
  assert.equal(xlsxOverflow(wb), 3);
  const xml = sheetXml(wb);
  assert.match(xml, /r="XFD1048576"/); assert.doesNotMatch(xml, /XFE|1048577/);
  const back = new Workbook(readXlsx(writeXlsx(wb)).data);
  assert.equal(back.getValue(0, R - 1, C - 1), 11);
  assert.equal(back.getValue(0, R, C - 1), null);
  assert.match(xml, /dimension ref="A1:XFD1048576"/);
});

test('XLSX 블록 경계: 정렬 perm·0·false·오류를 세고 null/빈값은 제외', () => {
  const wb = book([block(R - 1, 0, [[null, 0, false, {error:'#N/A'}, 'text', '']], [4, 0, 1, 2, 3, 5])]);
  assert.equal(wb.getValue(0, R - 1, 0), 'text');
  assert.equal(xlsxOverflow(wb), 3);
  const empty = book([block(R, 0, [[null, '']])]);
  assert.equal(xlsxOverflow(empty), 0, '순수 빈 블록은 제외');
  empty.sheets[0].blocks[0].cols[0].fmt = { fill:'#ff0000' };
  assert.equal(xlsxOverflow(empty), 2, '명시적 블록 서식도 한도 밖에서는 손실');
});

test('XLSX 블록: 일반 셀의 값·빈칸 override와 겹친 첫 블록 우선 규칙', () => {
  const wb = book([block(R, 0, [[1, 2, null, 4]]), block(R + 1, 0, [[20, 30, 40, 50]])], {
    [`${R},0`]: {raw:''}, [`${R + 1},0`]: {raw:'override'}, [`${R + 3},0`]: {raw:'', style:{bold:true}},
  });
  wb.sheets[0].cells.setRC(R, 0, {raw:'',v:null});
  assert.equal(xlsxOverflow(wb), 3, '값 override와 서식 override 각 1개, 첫 블록에 가리지 않은 마지막 값 1개');
  assert.equal(wb.getValue(0, R + 2, 0), null);
});

test('XLSX 한도 검사는 읽기 전용이며 전체 좌표 Set·재계산 없이 반복 가능', () => {
  const wb = book([block(R - 1, 2, [[10, 20, 30]])], { '0,0':{raw:'=1+1', cached:2} });
  const b = wb.sheets[0].blocks[0], before = JSON.stringify(wb.serialize()), version = wb.version;
  wb.getValue = () => { throw new Error('한도 검사에서 계산 금지'); };
  assert.equal(xlsxOverflow(wb), 2); assert.equal(xlsxOverflow(wb), 2);
  assert.equal(wb.version, version); assert.equal(wb.sheets[0].blocks[0], b); assert.equal(JSON.stringify(wb.serialize()), before);
});

test('XLSX ordinary/Map 셀은 양축·수식·셀 그림을 검사하고 외부 캐시 시트 제외', () => {
  const wb = book([], {[`${R},0`]:{raw:'0'}, [`0,${C}`]:{raw:'FALSE'}, [`${R},${C}`]:{raw:'=1'}, [`${R+1},0`]:{raw:'',image:{src:'data:image/png;base64,AA=='}}, [`${R+2},0`]:{raw:'',style:{bold:true}}});
  wb.sheets[0].cells = new Map(wb.sheets[0].cells);
  assert.equal(xlsxOverflow(wb), 5);
  const normal = book(); normal.sheets.push({state:'veryHidden',external:1,cells:new Map([[`${R},0`,{raw:'cached'}]]), blocks:[]});
  assert.equal(xlsxOverflow(normal), 0);
});

function dataTableFile() {
  const wb = book([], { '0,0':{raw:'2'}, '0,1':{raw:'=A1*2'}, '1,0':{raw:'2'}, '1,1':{raw:'4'}, '2,0':{raw:'3'}, '2,1':{raw:'6'} });
  const files = unzip(writeXlsx(wb)), path = 'xl/worksheets/sheet1.xml';
  files[path] = textOf(files[path]).replace(/(<c\b[^>]*r="B2"[^>]*>)/, '$1<f t="dataTable" ref="B2:B3" r1="A1"/>');
  return zip(files);
}

test('가상 분석 TABLE: 계산 지원으로 오인하지 않고 가져오기·보관 복원·저장 후 경고 유지', () => {
  const parsed = readXlsx(dataTableFile());
  assert.equal(parsed.warnings.filter(w => w.includes('TABLE')).length, 1);
  assert.deepEqual(parsed.data.props.xlsxImportWarnings, ['dataTableValuesOnly']);
  let wb = new Workbook(parsed.data);
  assert.equal(wb.getValue(0, 1, 1), 4); assert.equal(wb.getValue(0, 2, 1), 6);
  assert.equal(wb.getCell(0, 1, 1).formula, undefined);
  wb = new Workbook(wb.serialize());
  assert.match(xlsxExportWarnings(wb)[0], /수식 없이 값/);
  wb.setInput(0, 0, 0, '9');
  assert.equal(wb.getValue(0, 1, 1), 4, 'TABLE 미지원 상태를 숨기지 않음');
  wb.props.markedFinal = true;
  for (let pass = 0; pass < 2; pass++) {
    const bytes = writeXlsx(wb), files = unzip(bytes), loaded = readXlsx(bytes);
    assert.doesNotMatch(textOf(files['xl/worksheets/sheet1.xml']), /t="dataTable"/);
    assert.equal(loaded.warnings.filter(w => w.includes('TABLE')).length, 1);
    assert.ok(loaded.data.props.markedFinal);
    wb = new Workbook(loaded.data);
    assert.equal(xlsxExportWarnings(wb).filter(w => w.includes('TABLE')).length, 1);
  }
});

test('변환 이력은 알려진 code만 수용하며 파일의 임의 경고문/속성은 노출하지 않음', () => {
  const wb = book(); wb.props = {xlsxImportWarnings:['untrusted <img>', 'dataTableValuesOnly', 'dataTableValuesOnly']};
  assert.equal(xlsxExportWarnings(wb).length, 1);
  const bytes = writeXlsx(wb), files = unzip(bytes), xml = textOf(files['docProps/custom.xml']);
  assert.doesNotMatch(xml, /untrusted/);
  files['docProps/custom.xml'] = xml.replace('dataTableValuesOnly', 'untrusted');
  assert.deepEqual(readXlsx(zip(files)).warnings, []);
  assert.deepEqual(xlsxExportWarnings(book()), []);
});

function chartBook(type, axes, secondary = false) {
  const wb = book(); wb.sheets[0].charts = [{id:'c',type,x:0,y:0,w:480,h:320,axes,
    series:[{name:{text:'A'},cache:[1,10,100],xCache:[2,20,200]}, ...(secondary?[{name:{text:'B'},cache:[2,20,200],xCache:[3,30,300]}]:[])],
    ...(secondary ? {seriesFmt:[{}, {axis:'secondary'}]} : {})}]; return wb;
}
function nativeOnly(wb) {
  const files = unzip(writeXlsx(wb));
  files['xl/charts/chart1.xml'] = textOf(files['xl/charts/chart1.xml']).replace(/<c:extLst>[\s\S]*?<\/c:extLst>/g, '');
  return {files, parsed:readXlsx(zip(files))};
}

test('로그 축: 세로/가로막대·선의 주축/보조축 표준 logBase를 확장 없이 왕복', () => {
  for (const type of ['column', 'bar', 'line']) {
    const wb = chartBook(type, {y:{logBase:10,min:1,max:1000},y2:{logBase:2}}, true);
    const {files,parsed} = nativeOnly(wb), axes = parsed.data.sheets[0].charts[0].axes;
    assert.equal(axes.y.logBase, 10, type); assert.equal(axes.y2.logBase, 2, type);
    assert.equal(axes.y.min, 1); assert.equal(axes.y.max, 1000);
    assert.equal(parsed.warnings.filter(w => w.includes('로그 축')).length, 1);
    assert.match(xlsxExportWarnings(new Workbook(parsed.data))[0], /선형 축/);
    const scaling = descendants(parseXml(files['xl/charts/chart1.xml']), 'scaling').filter(n => child(n,'logBase'));
    assert.ok(scaling.every(n => n.children[0].name === 'logBase'));
  }
});

test('분산형/거품형: 서로 다른 X/Y 로그 밑·범위·역순을 표준 숫자 축에 왕복', () => {
  for (const type of ['scatter', 'bubble']) {
    const wb = chartBook(type, {x:{logBase:2,min:2,max:256,reverse:true,title:'합성 X'},y:{logBase:10,min:1,max:1000}});
    const {parsed} = nativeOnly(wb), axes = parsed.data.sheets[0].charts[0].axes;
    assert.equal(axes.x.logBase, 2); assert.equal(axes.y.logBase, 10);
    assert.equal(axes.x.min, 2); assert.equal(axes.x.max, 256); assert.equal(axes.x.reverse, true); assert.equal(axes.x.title, '합성 X');
    const twice = nativeOnly(new Workbook(parsed.data)).parsed;
    assert.equal(twice.data.sheets[0].charts[0].axes.x.logBase, 2);
  }
});

test('반복 계산·표시 정밀도 경고는 현재 설정을 조회하고 값·설정을 변경하지 않음', () => {
  const wb = book(); wb.calculation = {mode:'manual', iterate:true, fullPrecision:false};
  const before = JSON.stringify(wb.serialize());
  assert.equal(xlsxExportWarnings(wb).length, 2);
  assert.equal(JSON.stringify(wb.serialize()), before);
});


test('블록 한도 무작위 소형 오라클: 겹침·희소·perm·override를 실제 보이는 값과 대조', () => {
  let seed = 20261003;
  const rand = n => { seed = (Math.imul(seed, 1664525) + 1013904223) >>> 0; return seed % n; };
  for (let t = 0; t < 40; t++) {
    const bs = [];
    for (let k = 0; k < 4; k++) {
      const columns = Array.from({length:3}, () => Array.from({length:6}, () => [null, null, 0, false, 'x', 7][rand(6)]));
      bs.push(block(R - 3 + rand(5), C - 3 + rand(4), columns, rand(2) ? [5,2,0,4,1,3] : null));
    }
    const wb = book(bs);
    for (let i=0;i<8;i++) wb.sheets[0].cells.setRC(R-2+rand(8),C-2+rand(5),rand(2)?{raw:'',v:null}:{raw:'fixed',v:'fixed'});
    let expected=0;
    for(let r=R-3;r<R+10;r++) for(let c=C-3;c<C+5;c++) {
      if(r<R&&c<C) continue;
      const value=wb.getValue(0,r,c); if(value!==null&&value!==undefined&&value!=='') expected++;
    }
    assert.equal(xlsxOverflow(wb),expected,`fixture ${t}`);
  }
});


test('한도 밖 빈칸의 명시적 서식·메모·링크·윗주는 경고하고 순수 blank override는 제외', () => {
  const wb = book([block(R, 0, [[1,2,3,4,5,6]])]);
  const cells = wb.sheets[0].cells;
  cells.setRC(R,0,{raw:'',v:null});
  cells.setRC(R+1,0,{raw:'',v:null,style:{fill:'#aabbcc'}});
  cells.setRC(R+2,0,{raw:'',v:null,comment:{text:'합성 메모'}});
  cells.setRC(R+3,0,{raw:'',v:null,link:'#Sheet1!A1'});
  cells.setRC(R+4,0,{raw:'',v:null,phonetic:{visible:false,runs:[]}});
  cells.setRC(R+5,0,{raw:'',v:null,style:{}});
  assert.equal(xlsxOverflow(wb),4);
  cells.setRC(0,C,{raw:'',style:{numFmt:'text'}});
  assert.equal(xlsxOverflow(wb),5,'열 한도 밖의 명시적 서식 포함');
});
