import { test } from 'node:test';
import assert from 'node:assert/strict';
import { Workbook } from '../src/workbook.js';
import { writeXlsx, readXlsx } from '../src/xlsx.js';
import { convertXlsb } from '../src/xlsb.js';
import { parseXml, child } from '../src/xml.js';
import { readOds, writeOds } from '../src/ods.js';
import { zip, unzip, textOf } from '../src/zip.js';

function xmlFixture({ calc = '', protection = '', views = '' } = {}) {
  const wb = new Workbook({ sheets: [{ name: '설정', fileValues: true, cells: { '0,0': { raw: '3' }, '0,1': { raw: '=SUM(A1,2)', cached: 5 }, '0,2': { raw: '=NO_SUCH_FUNCTION()', cached: 99 } } }] });
  const files = unzip(writeXlsx(wb));
  files['xl/workbook.xml'] = textOf(files['xl/workbook.xml']).replace(/<calcPr[^>]*\/>/, calc).replace('<bookViews>', `${protection}<bookViews>`);
  if (views) files['xl/worksheets/sheet1.xml'] = textOf(files['xl/worksheets/sheet1.xml']).replace(/<sheetViews>[\s\S]*?<\/sheetViews>/, `<sheetViews>${views}</sheetViews>`);
  return { files, parsed: () => readXlsx(zip(files)) };
}
const roundtrip = wb => new Workbook(readXlsx(writeXlsx(wb)).data);

test('계산 설정을 표준 calcPr에서 읽고 JSON·XLSX·Undo에서 보존한다', () => {
  const { data, warnings } = xmlFixture({ calc: '<calcPr calcId="123456" calcMode="manual" refMode="R1C1" iterate="1" iterateCount="77" iterateDelta="0.0025" fullPrecision="0" calcOnSave="0" fullCalcOnLoad="0" forceFullCalc="1" concurrentCalc="0" concurrentManualCount="3" calcCompleted="1"/>' }).parsed();
  const wb = new Workbook(data), expected = { mode: 'manual', fullCalcOnLoad: false, iterate: true, fullPrecision: false, calcCompleted: true, calcOnSave: false, concurrentCalc: false, forceFullCalc: true, calcId: 123456, iterateCount: 77, concurrentManualCount: 3, iterateDelta: .0025, refMode: 'R1C1' };
  assert.deepEqual(wb.calculation, expected); assert.equal(wb.manualCalc, true);
  assert.equal(warnings.filter(s => /반복 계산|표시된 정밀도/.test(s)).length, 2);
  assert.deepEqual(roundtrip(wb).calculation, expected);
  assert.deepEqual(new Workbook(wb.serialize()).calculation, expected);
  wb.transact(() => wb.setBookProp('calculation', { ...expected, mode: 'autoNoTable' }));
  assert.equal(wb.manualCalc, false); wb.undo(); assert.equal(wb.manualCalc, true); wb.redo(); assert.equal(wb.manualCalc, false);
  assert.equal(roundtrip(wb).calculation.mode, 'autoNoTable');
  wb.load({ sheets: [{ name: '새 문서', cells: {} }] }); assert.equal(wb.manualCalc, false); assert.equal(wb.calculation, null);
});

test('수동 계산은 편집 결과를 F9까지 지연하고 지원하지 않는 무관한 저장값은 유지한다', () => {
  const wb = new Workbook(xmlFixture({ calc: '<calcPr calcMode="manual"/>' }).parsed().data);
  assert.equal(wb.getValue(0, 0, 1), 5); assert.equal(wb.getValue(0, 0, 2), 99);
  wb.transact(() => wb.setInput(0, 0, 0, '8'));
  assert.equal(wb.getValue(0, 0, 1), 5); assert.equal(wb.needsCalc, true);
  wb.calculateNow(); wb.invalidate();
  assert.equal(wb.getValue(0, 0, 1), 10); assert.equal(wb.getValue(0, 0, 2), 99); assert.equal(wb.manualCalc, true);
});

test('구조 보호 암호·창·수정 기록 보호를 표준 속성으로 왕복하고 임의 속성은 제외한다', () => {
  const protection = '<workbookProtection workbookPassword="CAFE" revisionsPassword="BEEF" lockStructure="1" lockWindows="1" lockRevision="1" workbookAlgorithmName="SHA-512" workbookHashValue="AA==" workbookSaltValue="AQ==" workbookSpinCount="100000" revisionsAlgorithmName="SHA-512" revisionsHashValue="Ag==" revisionsSaltValue="Aw==" revisionsSpinCount="200000" invented="lost"/>';
  const wb = new Workbook(xmlFixture({ protection }).parsed().data);
  assert.equal(wb.props.lockStructure, true); assert.equal(wb.props.workbookProtection.invented, undefined);
  assert.deepEqual(roundtrip(wb).props.workbookProtection, wb.props.workbookProtection);
  const copy = wb.serialize(); copy.props.workbookProtection.workbookPassword = '0000'; assert.equal(wb.props.workbookProtection.workbookPassword, 'CAFE');
  const properties = { ...wb.props, lockStructure: false, workbookProtection: { ...wb.props.workbookProtection } };
  for (const k of Object.keys(properties.workbookProtection)) if (k.startsWith('workbook')) delete properties.workbookProtection[k];
  wb.transact(() => wb.setBookProp('props', properties));
  const back = roundtrip(wb); assert.equal(back.props.lockStructure, undefined); assert.equal(back.props.workbookProtection.lockWindows, '1'); assert.equal(back.props.workbookProtection.revisionsPassword, 'BEEF');
});

test('각 시트의 머리글·수식 표시·눈금선 색을 좌표/고정창과 함께 표준 저장한다', () => {
  const f = xmlFixture({ views: '<sheetView workbookViewId="1" showRowColHeaders="1"/><sheetView workbookViewId="0" showRowColHeaders="0" showFormulas="1" defaultGridColor="0" colorId="10" topLeftCell="A5"><pane ySplit="2" topLeftCell="A20" state="frozen" activePane="bottomLeft"/><selection pane="bottomLeft" activeCell="C22"/></sheetView>' });
  let styles = textOf(f.files['xl/styles.xml']);
  styles = styles.replace('</styleSheet>', '<colors><indexedColors>' + Array.from({ length: 64 }, (_, i) => `<rgbColor rgb="FF${i === 10 ? '123456' : '000000'}"/>`).join('') + '</indexedColors></colors></styleSheet>');
  f.files['xl/styles.xml'] = styles;
  const wb = new Workbook(f.parsed().data);
  assert.deepEqual(wb.sheets[0].view, { top: 19, left: 0, r: 21, c: 2, activePane: 'bottomLeft', headers: false, showFormulas: true, gridColor: '#ff0000' });
  wb.transact(() => { wb.addSheet('반대 설정'); wb.setSheetProp(1, 'view', { headers: true, showFormulas: false, gridColor: '#abcdef' }); });
  const back = roundtrip(wb);
  assert.deepEqual(back.sheets[0].view, wb.sheets[0].view);
  for (const k of ['headers', 'showFormulas', 'gridColor']) assert.equal(back.sheets[1].view[k], wb.sheets[1].view[k]);
  assert.deepEqual(back.sheets[0].freeze, wb.sheets[0].freeze);
  assert.equal(back.getValue(0, 0, 1), 5);
});

test('생략된 표시/계산 옵션은 새로 명시하지 않고 자동 눈금선 색을 강제색으로 바꾸지 않는다', () => {
  const wb = new Workbook(xmlFixture({ views: '<sheetView workbookViewId="0" colorId="64" defaultGridColor="1"/>' }).parsed().data);
  assert.equal(wb.calculation, null); assert.equal(wb.sheets[0].view?.headers, undefined); assert.equal(wb.sheets[0].view?.gridColor, undefined);
  const xml = textOf(unzip(writeXlsx(wb))['xl/worksheets/sheet1.xml']);
  assert.doesNotMatch(xml, /showFormulas|showRowColHeaders|defaultGridColor/);
});


test('calcOnSave=false로 저장하면 수동 계산의 이전 결과를 강제로 계산하지 않는다', () => {
  const wb = new Workbook(xmlFixture({ calc: '<calcPr calcMode="manual" calcOnSave="0"/>' }).parsed().data);
  wb.transact(() => wb.setInput(0, 0, 0, '8'));
  const before = roundtrip(wb);
  assert.equal(before.getValue(0, 0, 0), 8); assert.equal(before.getValue(0, 0, 1), 5);
  assert.equal(before.calculation.calcOnSave, false); assert.equal(before.manualCalc, true);
  wb.calculateNow(); const after = roundtrip(wb); assert.equal(after.getValue(0, 0, 1), 10);
});

test('ODS 자체 내보내기/읽기는 숨긴 시트·빈 행·열 및 마지막 숨김 범위를 유지한다', () => {
  const wb = new Workbook({ sheets: [{ name: '보임', cells: { '0,0': { raw: '합성' } }, hiddenRows: { 1: true, 3: true }, hiddenCols: { 1: true, 4: true } }, { name: '숨김', state: 'hidden', cells: { '0,0': { raw: '42' } } }] });
  const bytes = writeOds(wb.sheets.map((s, si) => ({ ...s, si, hidden: s.state === 'hidden' })), {
    raw: (si, r, c) => wb.getRaw(si, r, c), value: (si, r, c) => wb.getValue(si, r, c), style: (si, r, c) => wb.styleAt(si, r, c),
    used: si => wb.usedRange(si), colWidth: (si, c) => wb.colWidth(si, c), rowHeight: (si, r) => wb.rowHeight(si, r),
  });
  const back = new Workbook(readOds(bytes).data);
  assert.equal(back.sheets[1].state, 'hidden'); assert.deepEqual(back.sheets[0].hiddenRows, { 1: true, 3: true }); assert.deepEqual(back.sheets[0].hiddenCols, { 1: true, 4: true });
  assert.equal(back.getValue(0, 0, 0), '합성'); assert.equal(back.getValue(1, 0, 0), 42);
});

const uint = (...xs) => { const b = Buffer.alloc(xs.length * 4); xs.forEach((v, i) => b.writeUInt32LE(v >>> 0, i * 4)); return b; };
const variable = x => { const a = []; do { const n = x & 127; x >>>= 7; a.push(n | (x ? 128 : 0)); } while (x); return Buffer.from(a); };
const record = (id, data) => Buffer.concat([variable(id), variable(data.length), data]);
const binaryXml = (book = [], sheet = []) => {
  const f = { 'xl/workbook.bin': Buffer.concat(book), 'xl/worksheets/sheet1.bin': Buffer.concat(sheet),
    'xl/_rels/workbook.bin.rels': Buffer.from('<Relationships><Relationship Id="r1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.bin"/></Relationships>') };
  for (const _ of convertXlsb(f)) { /* consume */ }
  return { book: parseXml(textOf(f['xl/workbook.bin'])), sheet: parseXml(textOf(f['xl/worksheets/sheet1.bin'])) };
};
test('XLSB 계산/보호 레코드를 표준 XML로 옮기며 권한 플래그 방향을 맞춘다', () => {
  const calc = Buffer.alloc(26); calc.writeUInt32LE(191029); calc.writeUInt32LE(2, 4); calc.writeUInt32LE(77, 8); calc.writeDoubleLE(.0025, 12); calc.writeUInt32LE(4, 20); calc.writeUInt16LE(0x184, 24);
  const book = Buffer.alloc(6); book.writeUInt16LE(0xCAFE); book.writeUInt16LE(0xBEEF, 2); book.writeUInt16LE(7, 4);
  const protection = Buffer.concat([Buffer.from([0x34, 0x12]), uint(1, 0, 1, 1, 0, 1, 0, 1, 0, 1, 0, 0, 1, 1, 0, 1)]);
  const x = binaryXml([record(157, calc), record(534, book)], [record(535, protection)]);
  assert.deepEqual(child(x.book, 'calcPr').attrs, { calcId: '191029', calcMode: 'autoNoTable', iterateCount: '77', iterateDelta: '0.0025', fullCalcOnLoad: '0', refMode: 'R1C1', iterate: '1', fullPrecision: '0', calcCompleted: '1', calcOnSave: '0', concurrentCalc: '0', concurrentManualCount: '4', forceFullCalc: '1' });
  assert.deepEqual(child(x.book, 'workbookProtection').attrs, { workbookPassword: 'CAFE', revisionsPassword: 'BEEF', lockStructure: '1', lockWindows: '1', lockRevision: '1' });
  const a = child(x.sheet, 'sheetProtection').attrs;
  assert.equal(a.password, '1234'); assert.equal(a.sheet, '1'); assert.equal(a.objects, '1'); assert.equal(a.formatCells, '0'); assert.equal(a.formatColumns, '1'); assert.equal(a.selectLockedCells, '1'); assert.equal(a.selectUnlockedCells, '0');
});
test('XLSB 보호 해제는 유지하고 잘린 보호 설정은 무보호로 열리지 않는다', () => {
  const unlocked = Buffer.alloc(66), x = binaryXml([record(157, Buffer.alloc(12))], [record(535, unlocked)]);
  assert.equal(child(x.book, 'calcPr'), null); assert.equal(child(x.book, 'workbookProtection'), null); assert.equal(child(x.sheet, 'sheetProtection').attrs.sheet, '0');
  assert.throws(() => binaryXml([record(534, Buffer.alloc(4))]), /보호 설정이 잘렸/);
  assert.throws(() => binaryXml([], [record(535, Buffer.alloc(4))]), /보호 설정이 잘렸/);
});

const wide = s => Buffer.concat([uint(s.length), Buffer.from(s, 'utf16le')]);
const isoData = (hash, salt, alg) => Buffer.concat([uint(hash.length), Buffer.from(hash), uint(salt.length), Buffer.from(salt), wide(alg)]);
test('XLSB ISO 보호 해시는 뒤따르는 빈 레거시 해시에 의해 사라지지 않는다', () => {
  const bookIso = Buffer.concat([uint(100000, 123), Buffer.from([7, 0]), isoData([1,2,3], [4,5], 'SHA-512'), isoData([6,7], [], 'SHA-256')]);
  const bookLegacy = Buffer.from([0,0,0,0,7,0]);
  const allowed = uint(1,0,1,1,0,1,0,1,0,1,0,0,1,1,0,1);
  const sheetIso = Buffer.concat([uint(100000), allowed, isoData([1,2,3], [4,5], 'SHA-512')]);
  const xml = binaryXml([record(677, bookIso), record(534, bookLegacy)], [record(678, sheetIso), record(535, Buffer.concat([Buffer.alloc(2), allowed]))]);
  const a=child(xml.book,'workbookProtection').attrs,b=child(xml.sheet,'sheetProtection').attrs;
  assert.equal(a.workbookHashValue, 'AQID'); assert.equal(a.workbookSaltValue,'BAU='); assert.equal(a.workbookSpinCount,'100000'); assert.equal(a.revisionsHashValue,'Bgc='); assert.equal(a.revisionsSpinCount,'123'); assert.equal(a.workbookPassword,undefined);
  assert.equal(b.algorithmName,'SHA-512'); assert.equal(b.hashValue,'AQID'); assert.equal(b.formatCells,'0'); assert.equal(b.password,undefined);
  assert.throws(()=>binaryXml([record(677,Buffer.alloc(9))]),/보호/);
  const malformed=Buffer.concat([uint(1,0),Buffer.from([1,0]),uint(100)]); assert.throws(()=>binaryXml([record(677,malformed)]),/보호/);
});
test('XLSB 인쇄 배율·방향·순서·맞춤·머리말과 활성 선택은 표준 요소로 전달한다', () => {
  const setup=Buffer.concat([uint(9,85,600,600,2,4,2,3),Buffer.from([0x9b,0]),uint(0xffffffff)]);
  const view=Buffer.alloc(30);view.writeUInt16LE(4|8|512);view.writeUInt16LE(75,18);
  const pane=Buffer.alloc(29);pane.writeDoubleLE(1);pane.writeDoubleLE(2,8);pane.writeInt32LE(14,16);pane.writeInt32LE(1,20);pane[28]=1;
  const sel=Buffer.concat([uint(0,19,2,0,1),uint(19,19,2,2)]);
  const header=Buffer.concat([Buffer.from([12,0]),wide('&L제목'),wide('&R&P'),wide(''),wide(''),wide(''),wide('')]);
  const x=binaryXml([],[record(137,view),record(151,pane),record(152,sel),record(478,setup),record(477,Buffer.from([15,0])),record(479,header)]).sheet;
  const p=child(x,'pageSetup').attrs;assert.equal(p.orientation,'landscape');assert.equal(p.scale,'85');assert.equal(p.fitToWidth,'2');assert.equal(p.firstPageNumber,'4');assert.equal(p.pageOrder,'overThenDown');assert.equal(p.draft,'1');assert.equal(p.blackAndWhite,'1');
  assert.equal(child(x,'printOptions').attrs.headings,'1');assert.equal(child(child(x,'headerFooter'),'oddHeader').text,'&L제목');
  assert.equal(child(child(child(x,'sheetViews'),'sheetView'),'selection').attrs.activeCell,'C20');
});
test('XLSB 유효성 유형·연산자·IME·여러 영역과 상대참조를 보존한다', () => {
  const f=n=>Buffer.concat([uint(3),Buffer.from([0x1e,n,0]),uint(0)]);
  const empty=uint(0,0), strings=Buffer.concat([wide('오류'),wide('확인'),wide('입력'),wide('안내')]);
  const dval=(fl,formula1=f(1),formula2=f(10))=>record(64,Buffer.concat([uint(fl,2),uint(3,3,1,1),uint(5,5,1,1),strings,formula1,formula2]));
  const x=binaryXml([],[dval(1|256|512|(9<<10)|0xc0000),dval(7|(2<<4)|(4<<20),Buffer.concat([uint(7),Buffer.from([0x4c,0,0,0,0,0,0xc0]),uint(0)]),empty)]).sheet;
  const list=child(x,'dataValidations').children;assert.equal(list.length,2);
  assert.equal(list[0].attrs.sqref,'B4 B6');assert.equal(list[0].attrs.imeMode,'fullHangul');assert.equal(list[0].attrs.showDropDown,'1');assert.equal(list[0].attrs.allowBlank,'1');assert.equal(child(list[0],'formula1').text,'1');assert.equal(child(list[0],'formula2').text,'10');
  assert.equal(list[1].attrs.type,'custom');assert.equal(list[1].attrs.errorStyle,'information');assert.equal(child(list[1],'formula1').text,'B4');assert.equal(child(list[1],'formula2'),null);
});
test('현재 시트 계산은 다른 시트의 수동 결과를 유지하고 F9는 나중에 모두 반영한다', () => {
  const sheet=name=>({name,fileValues:true,cells:{'0,0':{raw:'3'},'0,1':{raw:'=A1*2',cached:6},'0,2':{raw:'=NOT_SUPPORTED()',cached:42}}});
  const wb=new Workbook({calculation:{mode:'manual'},sheets:[sheet('A'),sheet('B')]});
  wb.transact(()=>{wb.setInput(0,0,0,'5');wb.setInput(1,0,0,'7');});
  assert.equal(wb.getValue(0,0,1),6);assert.equal(wb.getValue(1,0,1),6);
  wb.calculateSheetNow(0);assert.equal(wb.getValue(0,0,1),10);assert.equal(wb.getValue(1,0,1),6);assert.equal(wb.getValue(0,0,2),42);assert.equal(wb.getValue(1,0,2),42);assert.equal(wb.needsCalc,true);
  wb.calculateNow();wb.invalidate();assert.equal(wb.getValue(1,0,1),14);assert.equal(wb.getValue(0,0,2),42);
});
test('부분 무효화의 명시적 preserveImported는 미지원 clean 저장값을 유지한다', () => {
  const wb=new Workbook(xmlFixture().parsed().data);wb.invalidate(0,true);assert.equal(wb.getValue(0,0,2),99);
});
test('임의 눈금선 RGB는 Excel 근사색과 WIXEL 원색으로 구분하며 외부 편집을 덮지 않는다',()=>{
 const wb=new Workbook({sheets:[{name:'색',cells:{},view:{gridColor:'#123456'}}]});const files=unzip(writeXlsx(wb));
 assert.equal(new Workbook(readXlsx(zip(files)).data).sheets[0].view.gridColor,'#123456');
 assert.match(textOf(files['xl/worksheets/sheet1.xml']),/colorId="56"/);
 files['xl/worksheets/sheet1.xml']=textOf(files['xl/worksheets/sheet1.xml']).replace('colorId="56"','colorId="10"');
 assert.equal(new Workbook(readXlsx(zip(files)).data).sheets[0].view.gridColor,'#ff0000');
});
test('분산형 smooth=false와 파선·폭은 WIXEL 메타 없이 표준 왕복된다',()=>{
 const wb=new Workbook({sheets:[{name:'차트',cells:{},charts:[{id:'c',x:0,y:0,w:400,h:300,type:'scatter',scatterStyle:'smoothMarker',series:[{name:{text:'합성'},cache:[1,3,2],xCache:[1,2,3]}],seriesFmt:[{smooth:false,dash:'dash',lineWidth:4}]}]}]});
 const files=unzip(writeXlsx(wb));const path=Object.keys(files).find(k=>/^xl\/charts\/chart\d+\.xml$/.test(k));
 let xml=textOf(files[path]);assert.match(xml,/<c:smooth val="0"\/>/);files[path]=xml.replace(/<c:extLst>[\s\S]*?<\/c:extLst>/g,'');
 const back=new Workbook(readXlsx(zip(files)).data).sheets[0].charts[0];assert.equal(back.seriesFmt[0].smooth,false);assert.equal(back.seriesFmt[0].dash,'dash');assert.equal(back.seriesFmt[0].lineWidth,4);
});


test('보호된 ODS/FODS는 무보호로 열지 않고 미지원 주요 설정은 경고한다',()=>{
 const body=extra=>`<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0"><office:body><office:spreadsheet>${extra}<table:table table:name="S"><table:table-row><table:table-cell office:value-type="float" office:value="7"/></table:table-row></table:table></office:spreadsheet></office:body></office:document-content>`;
 for(const xml of [body('').replace('<office:spreadsheet>','<office:spreadsheet table:structure-protected="true">'),body('').replace('table:name="S"','table:name="S" table:protected="1"')])assert.throws(()=>readOds(xml),/보호된 OpenDocument/);
 const x=readOds(body('<table:content-validations><table:content-validation table:name="v"/></table:content-validations><table:calculation-settings/><table:database-ranges><table:database-range table:name="f"/></table:database-ranges>'));
 assert.equal(new Workbook(x.data).getValue(0,0,0),7);assert.match(x.warnings.join(' '),/데이터 유효성/);assert.match(x.warnings.join(' '),/계산 옵션/);assert.match(x.warnings.join(' '),/필터/);
 assert.deepEqual(readOds(body('')).warnings,[]);
});


test('ODS 마지막 숨김 행 하나는 빈 간격을 압축하며 100만 셀을 조회하지 않는다',()=>{
 let visits=0;const io={used:()=>({rows:1,cols:1}),raw:()=>{visits++;return '7';},value:()=>7,style:()=>null,colWidth:()=>80,rowHeight:()=>20};
 const bytes=writeOds([{si:0,name:'S',hiddenRows:{1048575:true}}],io);assert.equal(visits,1);
 const xml=textOf(unzip(bytes)['content.xml']);assert.match(xml,/number-rows-repeated="1048574"/);assert.ok(xml.length<10000);
 const parsed=readOds(bytes);assert.deepEqual(parsed.data.sheets[0].hiddenRows,{1048575:true});assert.equal(new Workbook(parsed.data).getValue(0,0,0),7);
 const compressed=xml.replace('table:number-rows-repeated="1048574"','table:number-rows-repeated="1048574" table:visibility="collapse"');
 assert.throws(()=>readOds(compressed),/숨김 행이 100,000개/);
});
