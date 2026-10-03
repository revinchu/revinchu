import test from 'node:test';
import assert from 'node:assert/strict';
import { readXls } from '../src/xls.js';
import { readXlsx, writeXlsx } from '../src/xlsx.js';
import { Workbook } from '../src/workbook.js';

// Minimal synthetic BIFF8 in a CFB v3 container. No user workbook bytes are used.
const words = (...n) => { const b = Buffer.alloc(n.length * 2); n.forEach((x, i) => b.writeUInt16LE(x & 0xffff, i * 2)); return b; };
const record = (type, data = Buffer.alloc(0)) => Buffer.concat([words(type, data.length), data]);
// Test checksum uses polynomial long division, independent of the importer's bitwise loop.
function extensionChecksum(bytes) {
  let crc = 0n;
  for (const b of bytes) {
    crc ^= BigInt(b) << 24n;
    for (let bit = 0; bit < 8; bit++) { crc <<= 1n; if (crc & 0x100000000n) crc ^= 0x1000000afn; }
  }
  return Number(crc);
}
const extColor = (type, rgb = [0x12, 0x34, 0x56], { colorType = 2, tint = 0, alpha = 255 } = {}) => {
  const d = Buffer.alloc(20); d.writeUInt16LE(type, 0); d.writeUInt16LE(20, 2); d.writeUInt16LE(colorType, 4); d.writeInt16LE(tint, 6); d.set([...rgb, alpha], 8); return d;
};
const xfExtension = (ix, props) => { const h = Buffer.alloc(20); h.writeUInt16LE(0x087d, 0); h.writeUInt16LE(ix, 14); h.writeUInt16LE(props.length, 18); return Buffer.concat([h, ...props]); };
function colorStyles() {
  return Array.from({ length: 16 }, () => {
    const d = Buffer.alloc(20); d.writeUInt16LE(1, 4); d[6] = 0x10;
    // Solid indexed red fill, black thin borders, both diagonals; every XF shares the FONT.
    d.writeUInt32LE(0xc0001111, 10); d.writeUInt32LE(0x06200000, 14); d.writeUInt16LE(10 | (9 << 7), 18); return d;
  });
}
function fixture({ zoom = [23, 20], rows = [], defaultTwips = 270, xfs, extensions = [], checksum = true, sheetRecords = [], globalRecords = [] } = {}) {
  const bof = type => record(0x0809, words(0x0600, type, 0x0dbb, 0x07cc, 0, 0, 6, 0));
  const font = Buffer.alloc(16); font.writeUInt16LE(220, 0); font.writeUInt16LE(0x7fff, 4); font.writeUInt16LE(400, 6); font[14] = 7;
  const xf = Buffer.alloc(20); xf.writeUInt16LE(1, 4); xf[6] = 0x10;
  const bound = Buffer.alloc(9); bound[6] = 1; bound[8] = 83;
  const styles = xfs || [xf], check = Buffer.alloc(20); check.writeUInt16LE(0x087c, 0); check.writeUInt16LE(styles.length, 14);
  check.writeUInt32LE((extensionChecksum(Buffer.concat(styles)) ^ (checksum === 'stale' ? 1 : 0)) >>> 0, 16);
  const globals = () => Buffer.concat([bof(5), record(0x0031, Buffer.concat([font, Buffer.from('Calibri')])), ...styles.map(x => record(0x00e0, x)),
    ...(extensions.length && checksum ? [record(0x087c, check)] : []), ...extensions.map(x => record(0x087d, x)), ...globalRecords, record(0x0085, bound), record(10)]);
  bound.writeUInt32LE(globals().length, 0);
  const rowRecords = rows.map(({ r, twips, manual = false, hidden = false, standardBit = false }) => {
    const d = Buffer.alloc(16); d.writeUInt16LE(r, 0); d.writeUInt16LE(1, 4); d.writeUInt16LE(twips | (standardBit ? 0x8000 : 0), 6);
    d.writeUInt16LE(0x0100 | (manual ? 0x40 : 0) | (hidden ? 0x20 : 0), 12); return record(0x0208, d);
  });
  const cells = styles.map((_, c) => { const cell = Buffer.alloc(14); cell.writeUInt16LE(c, 2); cell.writeUInt16LE(c, 4); cell.writeDoubleLE(42 + c, 6); return record(0x0203, cell); });
  const data = Buffer.concat([globals(), bof(0x0010), record(0x0225, words(0, defaultTwips)), ...rowRecords, ...cells, ...sheetRecords, ...(zoom ? [record(0x00a0, words(...zoom))] : []), record(10)]);
  assert.ok(data.length <= 4096);
  const file = Buffer.alloc(512 * 11); Buffer.from([0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1]).copy(file);
  file.writeUInt16LE(0x003e, 24); file.writeUInt16LE(3, 26); file.writeUInt16LE(0xfffe, 28); file.writeUInt16LE(9, 30); file.writeUInt16LE(6, 32);
  file.writeUInt32LE(1, 44); file.writeUInt32LE(8, 48); file.writeUInt32LE(4096, 56); file.writeUInt32LE(0xfffffffe, 60); file.writeUInt32LE(0xfffffffe, 68);
  for (let i = 0; i < 109; i++) file.writeUInt32LE(i === 0 ? 9 : 0xffffffff, 76 + i * 4);
  data.copy(file, 512);
  const directory = (slot, name, type, start, size) => {
    const at = 512 * 9 + slot * 128, text = Buffer.from(name + '\0', 'utf16le'); text.copy(file, at); file.writeUInt16LE(text.length, at + 64); file[at + 66] = type; file[at + 67] = 1;
    for (const offset of [68, 72, 76]) file.writeUInt32LE(0xffffffff, at + offset);
    file.writeUInt32LE(start, at + 116); file.writeUInt32LE(size, at + 120);
  };
  directory(0, 'Root Entry', 5, 0xfffffffe, 0); file.writeUInt32LE(1, 512 * 9 + 76); directory(1, 'Workbook', 2, 0, 4096);
  for (let i = 0; i < 128; i++) file.writeUInt32LE(i < 7 ? i + 1 : i === 7 || i === 8 ? 0xfffffffe : i === 9 ? 0xfffffffd : 0xffffffff, 512 * 10 + i * 4);
  return new Uint8Array(file);
}

test('XLS SCL ratios become model percentages and survive standard XLSX export', () => {
  for (const [num, den, percent] of [[9, 10, 90], [23, 20, 115], [5, 4, 125], [1, 4, 25], [1, 10, 10], [4, 1, 400], [2, 3, 67]]) {
    const wb = new Workbook(readXls(fixture({ zoom: [num, den] })).data);
    assert.equal(wb.sheets[0].zoom, percent, `${num}/${den}`);
    assert.equal(new Workbook(readXlsx(writeXlsx(wb)).data).sheets[0].zoom, percent);
    assert.equal(wb.getValue(0, 0, 0), 42);
  }
});

test('XLS default/equal/invalid SCL keeps 100% and invalid ratios are bounded', () => {
  for (const zoom of [null, [1, 1], [0, 1], [1, 0], [-1, 10], [1, -10]]) assert.equal(readXls(fixture({ zoom })).data.sheets[0].zoom, undefined);
  assert.equal(readXls(fixture({ zoom: [1, 100] })).data.sheets[0].zoom, 10);
  assert.equal(readXls(fixture({ zoom: [20, 1] })).data.sheets[0].zoom, 400);
});

test('XLS WINDOW2 and PLV distinguish normal, page-break preview and page-layout modes', () => {
  for (const [flags, plv, mode] of [[6, false, 'normal'], [0x0806, false, 'pageBreakPreview'], [6, true, 'pageLayout']]) {
    const window = Buffer.alloc(18); window.writeUInt16LE(flags);
    const page = Buffer.alloc(16); page.writeUInt16LE(0x088b); page.writeUInt16LE(100, 12); page.writeUInt16LE(1, 14);
    const wb = new Workbook(readXls(fixture({ sheetRecords: [record(0x023e, window), ...(plv ? [record(0x088b, page)] : [])] })).data);
    assert.equal(wb.sheets[0].view.mode, mode);
    assert.equal(new Workbook(readXlsx(writeXlsx(wb)).data).sheets[0].view.mode, mode);
    assert.equal(wb.getValue(0, 0, 0), 42);
  }
});

test('XLS manual page breaks use first following row/column indexes and tolerate truncated records', () => {
  const breaks = values => Buffer.concat([words(values.length), ...values.map(id => words(id, 0, 255))]);
  const sheetRecords = [record(0x001b, breaks([7, 19, 7, 0])), record(0x001a, breaks([3, 7, 256])), record(0x088b, words(1))];
  const wb = new Workbook(readXls(fixture({ sheetRecords })).data);
  assert.deepEqual(wb.sheets[0].page.rowBreaks, [7, 19]);
  assert.deepEqual(wb.sheets[0].page.colBreaks, [3, 7]);
  const back = new Workbook(readXlsx(writeXlsx(wb)).data);
  assert.deepEqual(back.sheets[0].page.rowBreaks, [7, 19]);
  assert.deepEqual(back.sheets[0].page.colBreaks, [3, 7]);
});

test('XLS builtin Print_Area unions and Print_Titles whole axes preserve all regions', () => {
  const area = (r1, r2, c1, c2) => Buffer.concat([Buffer.from([0x25]), words(r1, r2, c1, c2)]);
  const name = (id, formula) => { const d = Buffer.alloc(16); d.writeUInt16LE(0x20); d[3] = 1; d.writeUInt16LE(formula.length, 4); d.writeUInt16LE(1, 8); d[15] = id; return record(0x0018, Buffer.concat([d, formula])); };
  const globalRecords = [name(6, Buffer.concat([area(0, 19, 0, 2), area(0, 7, 4, 5), Buffer.from([0x10])])),
    name(7, Buffer.concat([area(0, 1, 0, 255), area(0, 65535, 0, 0), Buffer.from([0x10])]))];
  const wb = new Workbook(readXls(fixture({ globalRecords })).data), p = wb.sheets[0].page;
  assert.deepEqual(p.areas, [{ r1: 0, c1: 0, r2: 19, c2: 2 }, { r1: 0, c1: 4, r2: 7, c2: 5 }]);
  assert.deepEqual(p.area, p.areas[0]); assert.deepEqual(p.titleRows, [0, 1]); assert.deepEqual(p.titleCols, [0, 0]);
  const back = new Workbook(readXlsx(writeXlsx(wb)).data).sheets[0].page;
  assert.deepEqual(back.areas, p.areas); assert.deepEqual(back.titleRows, p.titleRows); assert.deepEqual(back.titleCols, p.titleCols);
});

test('XLS Setup/WSBOOL and margins keep print scaling separate from display zoom', () => {
  const setup = Buffer.alloc(34); setup.writeUInt16LE(9); setup.writeUInt16LE(85, 2); setup.writeUInt16LE(1, 6); setup.writeUInt16LE(2, 10); setup.writeDoubleLE(.3, 16); setup.writeDoubleLE(.4, 24);
  const margin = value => { const b = Buffer.alloc(8); b.writeDoubleLE(value); return b; };
  for (const fit of [false, true]) {
    const sheetRecords = [record(0x00a1, setup), record(0x0081, words(fit ? 0x100 : 0)), record(0x0026, margin(.15)), record(0x0027, margin(.2)), record(0x0028, margin(.25)), record(0x0029, margin(.1)), record(0x0083, words(1)), record(0x0084, words(0)), record(0x002a, words(1)), record(0x002b, words(1))];
    const wb = new Workbook(readXls(fixture({ zoom: [9, 10], sheetRecords })).data), s = wb.sheets[0], p = s.page;
    assert.equal(s.zoom, 90); assert.equal(p.scale, 85); assert.equal(p.paper, 9); assert.equal(p.orientation, 'portrait');
    assert.equal(p.fitW ?? 0, fit ? 1 : 0); assert.equal(p.fitH ?? 0, 0);
    assert.deepEqual(p.margins, { left: .15, right: .2, top: .25, bottom: .1, header: .3, footer: .4 });
    assert.equal(p.hCenter, true); assert.equal(p.vCenter, false); assert.equal(p.headings, true); assert.equal(p.gridlines, true);
    const back = new Workbook(readXlsx(writeXlsx(wb)).data).sheets[0];
    assert.equal(back.zoom, 90); assert.equal(back.page.scale, fit ? 100 : 85); assert.equal(back.page.fitW, fit ? 1 : 0); assert.deepEqual(back.page.margins, p.margins);
  }
});

test('XLS undefined printer fields and short Setup records cannot overwrite defaults', () => {
  const setup = Buffer.alloc(34); setup.writeUInt16LE(9); setup.writeUInt16LE(175, 2); setup.writeUInt16LE(2, 6); setup.writeUInt16LE(1, 8); setup.writeUInt16LE(5, 10); setup.writeDoubleLE(.25, 16); setup.writeDoubleLE(.3, 24);
  const p = readXls(fixture({ sheetRecords: [record(0x0081, words(0x100)), record(0x00a1, setup), record(0x00a1, words(1))] })).data.sheets[0].page;
  assert.equal(p.paper, undefined); assert.equal(p.scale, undefined); assert.equal(p.orientation, undefined);
  assert.equal(p.order, 'overThenDown'); assert.equal(p.fitW, 2); assert.equal(p.fitH, 1); assert.deepEqual(p.margins, { header: .25, footer: .3 });
});

test('XLS auto-fit ROW height is preserved independently of manual-height and hidden flags', () => {
  const wb = new Workbook(readXls(fixture({ rows: [
    { r: 0, twips: 375 }, { r: 1, twips: 270 }, { r: 2, twips: 360, manual: true },
    { r: 3, twips: 375, hidden: true }, { r: 4, twips: 270, manual: true }, { r: 5, twips: 375, standardBit: true },
  ] })).data), s = wb.sheets[0];
  assert.deepEqual(Array.from({ length: 6 }, (_, r) => wb.rowHeight(0, r)), [25, 18, 24, 25, 18, 25]);
  assert.deepEqual(s.rowManual, { 2: true, 4: true });
  assert.deepEqual(s.hiddenRows, { 3: true });
  assert.equal(s.rowHeights[1], undefined, 'default-height rows remain sparse');
  const again = new Workbook(readXlsx(writeXlsx(wb)).data);
  assert.deepEqual(Array.from({ length: 6 }, (_, r) => again.rowHeight(0, r)), [25, 18, 24, 25, 18, 25]);
  assert.deepEqual(again.sheets[0].rowManual, { 2: true, 4: true });
  assert.deepEqual(again.sheets[0].hiddenRows, { 3: true });
});


test('XLS XFExt true RGB overrides palette foreground, text and each border without changing shared fonts or values', () => {
  const props = [extColor(4, [255, 192, 0]), extColor(13, [25, 50, 75]), ...[7, 8, 9, 10, 11].map((type, i) => extColor(type, [i + 1, 20, 30]))];
  const wb = new Workbook(readXls(fixture({ xfs: colorStyles(), extensions: [xfExtension(0, props)] })).data);
  const st = wb.styleAt(0, 0, 0), other = wb.styleAt(0, 0, 1);
  assert.equal(st.fill, '#ffc000'); assert.equal(st.color, '#19324b');
  assert.deepEqual([st.btc, st.bbc, st.blc, st.brc, st.ddc, st.duc], ['#01141e', '#02141e', '#03141e', '#04141e', '#05141e', '#05141e']);
  assert.equal(other.fill, '#ff0000'); assert.equal(other.color, undefined);
  assert.deepEqual(wb.defaultFont, { name: 'Calibri', size: 11 });
  assert.deepEqual(Array.from({ length: 16 }, (_, c) => wb.getValue(0, 0, c)), Array.from({ length: 16 }, (_, c) => 42 + c));
  const again = new Workbook(readXlsx(writeXlsx(wb)).data), saved = again.styleAt(0, 0, 0);
  for (const key of ['fill', 'color', 'btc', 'bbc', 'blc', 'brc', 'ddc', 'duc']) assert.equal(saved[key], st[key], key);
  assert.equal(again.styleAt(0, 0, 1).fill, '#ff0000'); assert.equal(again.getValue(0, 0, 0), 42);
});

test('XLS XFExt uses foreground/background separately for patterns and never adds an absent fill or border', () => {
  const xfs = colorStyles(); xfs[0].writeUInt32LE(0x0a200000, 14); xfs[1].writeUInt32LE(0x02000000, 14); xfs[1].writeUInt32LE(0, 10);
  const props = [extColor(4, [1, 2, 3]), extColor(5, [40, 50, 60]), extColor(7)];
  const wb = new Workbook(readXls(fixture({ xfs, extensions: [xfExtension(0, props), xfExtension(1, props)] })).data);
  assert.ok(wb.styleAt(0, 0, 0).pattern); assert.equal(wb.styleAt(0, 0, 0).patternColor, '#010203'); assert.equal(wb.styleAt(0, 0, 0).fill, '#28323c');
  const st = wb.getCell(0, 0, 1).style; assert.equal(st.fill, undefined); assert.equal(st.bt, undefined); assert.equal(st.btc, undefined);
});

test('XLS XFExt unsupported automatic/indexed/theme/tint/alpha/unknown values retain the base XF color', () => {
  const unsupported = [{ colorType: 0 }, { colorType: 1 }, { colorType: 3 }, { colorType: 4 }, { tint: 1000 }, { tint: -1000 }, { alpha: 0 }, { alpha: 127 }];
  for (const options of unsupported) {
    const wb = new Workbook(readXls(fixture({ xfs: colorStyles(), extensions: [xfExtension(0, [extColor(4, undefined, options), extColor(13, undefined, options), extColor(7, undefined, options)])] })).data);
    assert.equal(wb.styleAt(0, 0, 0).fill, '#ff0000'); assert.equal(wb.styleAt(0, 0, 0).color, undefined); assert.equal(wb.getCell(0, 0, 0).style.btc, undefined);
  }
  const unknown = Buffer.from([0xff, 0xff, 6, 0, 1, 2]);
  const st = new Workbook(readXls(fixture({ xfs: colorStyles(), extensions: [xfExtension(0, [unknown, extColor(4)])] })).data).styleAt(0, 0, 0);
  assert.equal(st.fill, '#123456', 'well-formed unknown properties do not hide supported following colors');
});

test('XLS XFExt rejects malformed records, stale or absent XFCRC, and an XF without extension flag', () => {
  const valid = xfExtension(0, [extColor(4)]), badSize = Buffer.from(valid); badSize.writeUInt16LE(3, 22);
  const shortColor = xfExtension(0, [Buffer.from([4, 0, 4, 0])]);
  const badCount = Buffer.from(valid); badCount.writeUInt16LE(2, 18);
  const partial = xfExtension(0, [extColor(4), Buffer.from([13, 0, 20, 0])]);
  const wrongHeader = Buffer.from(valid); wrongHeader.writeUInt16LE(0, 0);
  for (const d of [valid.subarray(0, 18), valid.subarray(0, 39), badSize, shortColor, badCount, partial, wrongHeader, xfExtension(16, [extColor(4)])]) {
    assert.equal(new Workbook(readXls(fixture({ xfs: colorStyles(), extensions: [d] })).data).styleAt(0, 0, 0).fill, '#ff0000');
  }
  for (const checksum of [false, 'stale']) assert.equal(new Workbook(readXls(fixture({ xfs: colorStyles(), extensions: [valid], checksum })).data).styleAt(0, 0, 0).fill, '#ff0000');
  const xfs = colorStyles(); xfs[0].writeUInt32LE(0x04200000, 14);
  assert.equal(new Workbook(readXls(fixture({ xfs, extensions: [valid] })).data).styleAt(0, 0, 0).fill, '#ff0000');
});

test('XLS 시트 스트림 계산 옵션과 표시 옵션은 XLSX 표준으로 왕복된다', () => {
  const delta=Buffer.alloc(8);delta.writeDoubleLE(.0025);
  const win=Buffer.alloc(18);win.writeUInt16LE(1|2|8,0);win.writeUInt16LE(4,2);win.writeUInt16LE(2,4);win.writeUInt16LE(56,6);
  const pane=Buffer.concat([words(1,2,14,3),Buffer.from([0])]);
  const bytes=fixture({globalRecords:[record(0x000e,words(0)),record(0x0012,words(1)),record(0x0013,words(0xcafe))],sheetRecords:[record(0x000d,words(0)),record(0x000c,words(77)),record(0x000f,words(1)),record(0x0011,words(1)),record(0x0010,delta),record(0x005f,words(0)),record(0x023e,win),record(0x0041,pane)]});
  const a=readXls(bytes),wb=new Workbook(a.data);assert.deepEqual(wb.calculation,{mode:'manual',fullPrecision:false,iterateCount:77,refMode:'A1',iterate:true,iterateDelta:.0025,calcOnSave:false});
  const s=wb.sheets[0];assert.equal(s.view.headers,false);assert.equal(s.view.showFormulas,true);assert.equal(s.view.gridColor,'#003366');assert.equal(s.noZeros,true);assert.deepEqual(s.freeze,{rows:2,cols:1,top:4,left:2});assert.equal(s.view.top,14);assert.equal(s.view.left,3);assert.equal(wb.props.workbookProtection.workbookPassword,'CAFE');
  const back=new Workbook(readXlsx(writeXlsx(wb)).data);assert.deepEqual(back.calculation,wb.calculation);assert.deepEqual(back.sheets[0].view,s.view);assert.deepEqual(back.sheets[0].freeze,s.freeze);
});
test('XLS 자동/자동 데이터 표 제외를 구별하고 기본 시트 보호를 유지한다',()=>{
 for(const [mode,expected]of [[1,'auto'],[0xffff,'autoNoTable']]){const x=readXls(fixture({sheetRecords:[record(0x000d,words(mode)),record(0x0012,words(1)),record(0x0013,words(0xe2bd))]}));assert.equal(x.data.calculation.mode,expected);assert.equal(x.data.sheets[0].protect.on,true);assert.equal(x.data.sheets[0].protect.hash,'E2BD');}
 assert.throws(()=>readXls(fixture({sheetRecords:[record(0x01be,Buffer.alloc(8))]})),/잘렸|개수/);
});


const dvText = s => Buffer.concat([words(s.length), Buffer.from([1]), Buffer.from(s, 'utf16le')]);
const dvInt = n => Buffer.concat([Buffer.from([0x1e]), words(n)]);
function dvRecord({type=1,flags=0,f1=dvInt(1),f2=dvInt(10),ranges=[[0,0,1,1]],strings=['입력','오류','설명','확인']}={}) {
 const head=Buffer.alloc(4);head.writeUInt32LE((type|flags)>>>0);
 return record(0x01be,Buffer.concat([head,...strings.map(dvText),words(f1.length,0),f1,words(f2.length,0),f2,words(ranges.length),...ranges.map(a=>words(...a))]));
}

test('XLS DV는 8유형·8연산자·오류/IME/메시지·목록을 XLSX로 보존한다',()=>{
 const types=['any','whole','decimal','list','date','time','textLength','custom'];
 const ops=['between','notBetween','equal','notEqual','greaterThan','lessThan','greaterThanOrEqual','lessThanOrEqual'];
 for(let i=0;i<8;i++){
  const strings=['입력','오류','설명','확인'], op=[0,3,7].includes(i)?0:i, fl=(op<<20)|((i%3)<<4)|256|(i===3?512:0)|0x40000|0x80000|(10<<10);
  const f1=i===3?Buffer.from([0x17,5,0,65,0,66,0,67]):dvInt(1);
  const a=readXls(fixture({sheetRecords:[dvRecord({type:i,flags:fl|(i===3?128:0),f1,f2:op<=1&&![0,3,7].includes(i)?dvInt(10):Buffer.alloc(0),strings})]}));
  const rule=a.data.sheets[0].validations[0];assert.equal(rule.type,types[i]);assert.equal(rule.op,ops[op]);assert.equal(rule.errorStyle,['stop','warning','info'][i%3]);assert.equal(rule.imeMode,'halfHangul');assert.equal(rule.showDropdown,i!==3);assert.equal(rule.allowBlank,true);
  for(const k of ['showPrompt','showError'])assert.equal(rule[k],true);
  for(const [j,k]of ['promptTitle','errorTitle','prompt','error'].entries())assert.equal(rule[k],strings[j]);
  if(i===3)assert.equal(rule.f1,'"A,B,C"');
  const back=readXlsx(writeXlsx(new Workbook(a.data))).data.sheets[0].validations[0];
  for(const key of Object.keys(rule))assert.deepEqual(back[key],rule[key],`type ${i} ${key}`);
 }
});

test('XLS DV 다중 영역은 첫 영역 상대수식을 옮기고 원본 셀 값을 바꾸지 않는다',()=>{
 const ref=Buffer.concat([Buffer.from([0x24]),words(0,0xc000),dvInt(0),Buffer.from([0x0d])]);
 const parsed=readXls(fixture({sheetRecords:[dvRecord({type:7,f1:ref,f2:Buffer.alloc(0),ranges:[[0,0,1,1],[3,4,3,3]]})]}));
 assert.equal(parsed.data.sheets[0].validations[0].f1,'A1>0');assert.equal(parsed.data.sheets[0].validations[1].f1,'C4>0');
 const wb=new Workbook(parsed.data);assert.equal(wb.getValue(0,0,0),42);
 const back=readXlsx(writeXlsx(wb)).data.sheets[0].validations;assert.equal(back.length,2);assert.equal(back[1].f1,'C4>0');
});

test('XLS 보호의 모든 권한과 암호는 표준 XLSX로 왕복하며 보호 없음도 유지한다',()=>{
 const keys=['objects',null,'formatCells','formatColumns','formatRows','insertColumns','insertRows','insertHyperlinks','deleteColumns','deleteRows','selectLocked','sort','autoFilter','pivotTables','selectUnlocked'];
 for(const bit of [0,2,3,4,5,6,7,8,9,10,11,12,13,14]){
  const feat=Buffer.alloc(23);feat.writeUInt16LE(0x867);feat.writeUInt16LE(2,12);feat[14]=1;feat.writeUInt32LE(0xffffffff,15);feat.writeUInt32LE(1<<bit,19);
  const wb=new Workbook(readXls(fixture({sheetRecords:[record(0x12,words(1)),record(0x13,words(0xe2bd)),record(0x867,feat)]})).data),p=wb.sheets[0].protect;
  assert.equal(p.hash,'E2BD');for(const key of keys.filter(Boolean))assert.equal(p.allow[key],key===keys[bit],`bit ${bit}/${key}`);
  assert.deepEqual(new Workbook(readXlsx(writeXlsx(wb)).data).sheets[0].protect,p);
 }
 assert.equal(readXls(fixture({sheetRecords:[record(0x12,words(0)),record(0x13,words(0xe2bd))]})).data.sheets[0].protect,undefined);
});

test('XLS 잘린 보호·잘못된 유효성·암호화는 무보호 또는 무규칙으로 열리지 않는다',()=>{
 for(const type of [0x12,0x13,0x63,0xdd])assert.throws(()=>readXls(fixture({sheetRecords:[record(type,Buffer.alloc(1))]})),/잘렸/);
 const full=dvRecord();for(const cut of [4,14,full.length-1])assert.throws(()=>readXls(fixture({sheetRecords:[record(0x1be,full.subarray(4,cut))]})),/잘렸|개수/);
 assert.throws(()=>readXls(fixture({sheetRecords:[dvRecord({type:15})]})),/지원하지 않는/);
 assert.throws(()=>readXls(fixture({sheetRecords:[dvRecord({ranges:[[2,1,0,0]]})]})),/범위/);
 assert.throws(()=>readXls(fixture({globalRecords:[record(0x2f,words(0,0,0))]})),/암호/);
});


test('XLS 숫자 유효성의 8개 비교 연산자는 전부 보존된다',()=>{
 for(const [i,op] of ['between','notBetween','equal','notEqual','greaterThan','lessThan','greaterThanOrEqual','lessThanOrEqual'].entries()){
  const wb=new Workbook(readXls(fixture({sheetRecords:[dvRecord({type:1,flags:i<<20})]})).data);
  assert.equal(wb.sheets[0].validations[0].op,op);assert.equal(readXlsx(writeXlsx(wb)).data.sheets[0].validations[0].op,op);
 }
});
