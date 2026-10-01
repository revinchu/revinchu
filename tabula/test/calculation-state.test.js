import { test } from 'node:test';
import assert from 'node:assert/strict';
import { Workbook } from '../src/workbook.js';
import { ERR } from '../src/fxcore.js';
import { readXlsx, writeXlsx } from '../src/xlsx.js';

const fixture = () => new Workbook({ sheets: [{ name: '입력', fileValues: true, cells: {
  '0,0': { raw: '2' }, '0,1': { raw: '=NO_SUCH_FN(A1)', cached: 20 },
  '0,2': { raw: '=SUM(B1,5)', cached: 25 }, '0,3': { raw: '=NO_SUCH_FN(99)', cached: 99 },
} }, { name: '요약', fileValues: true, cells: { '0,0': { raw: '=SUM(입력!B1:C1)', cached: 45 } } }] });
const edit = wb => wb.transact(() => wb.setInput(0, 0, 0, '3'));

test('파일을 처음 열면 저장값과 원문을 보존하고 재계산 미지원 상태를 구분한다', () => {
  const wb = fixture();
  assert.equal(wb.getValue(0, 0, 1), 20);
  assert.equal(wb.getValue(0, 0, 2), 25);
  assert.equal(wb.getRaw(0, 0, 1), '=NO_SUCH_FN(A1)');
  const state = wb.getCalculationStatus(0, 0, 1);
  assert.equal(state.status, 'cached');
  assert.equal(state.reason, 'unsupported-function');
  assert.deepEqual(state.functions, ['NO_SUCH_FN']);
  assert.equal(wb.getCalculationStatus(0, 0, 2).status, 'cached');
  assert.equal(wb.getCalculationStatus(0, 0, 0), null);
});

test('관련 입력 변경은 옛 미지원 결과와 그 합계를 #NAME?로 차단하며 무관한 캐시는 유지한다', () => {
  const wb = fixture();
  wb.getValue(0, 0, 2); wb.getValue(1, 0, 0);
  edit(wb);
  assert.equal(wb.getValue(0, 0, 1), ERR.NAME);
  assert.equal(wb.getValue(0, 0, 2), ERR.NAME);
  assert.equal(wb.getValue(1, 0, 0), ERR.NAME);
  assert.equal(wb.getValue(0, 0, 3), 99);
  const state = wb.getCalculationStatus(0, 0, 1);
  assert.equal(state.status, 'stale'); assert.equal(state.savedValue, 20);
  assert.equal(wb.getCalculationStatus(0, 0, 2).reason, 'unresolved-result');
  const list = wb.calculationIssues({ limit: 1 });
  assert.equal(list.items.length, 1); assert.equal(list.total, 4); assert.equal(list.truncated, true);
});

test('해석불가 수식은 다른 시트의 입력 변경에도 보수적으로 불신하고 의존 합계로 전파한다', () => {
  const wb = new Workbook({ sheets: [{ name: '원본', fileValues: true, cells: {
    '0,0': { raw: '=[1]Table[[#UNREADABLE', cached: 50 }, '0,1': { raw: '=SUM(A1,1)', cached: 51 },
  } }, { name: '다른시트', cells: {} }] });
  assert.equal(wb.getValue(0, 0, 0), 50);
  assert.equal(wb.getCalculationStatus(0, 0, 0).reason, 'unreadable-formula');
  wb.transact(() => wb.setInput(1, 0, 0, '변경'));
  assert.equal(wb.getValue(0, 0, 0), ERR.NAME);
  assert.equal(wb.getValue(0, 0, 1), ERR.NAME);
  wb.transact(() => wb.setInput(0, 0, 0, '10'));
  assert.equal(wb.getValue(0, 0, 1), 11);
});

test('입력 Undo/Redo는 미지원 캐시의 신뢰 상태를 되돌리고 서식 변경으로 신뢰를 승격하지 않는다', () => {
  const wb = fixture();
  edit(wb); assert.equal(wb.getValue(0, 0, 1), ERR.NAME);
  wb.undo(); assert.equal(wb.getValue(0, 0, 1), 20); assert.equal(wb.getValue(1, 0, 0), 45);
  wb.redo(); assert.equal(wb.getValue(0, 0, 1), ERR.NAME);
  wb.transact(() => wb.setStyle(0, 0, 1, { bold: true }));
  assert.equal(wb.getValue(0, 0, 1), ERR.NAME);
  wb.undo(); wb.undo();
  assert.equal(wb.getValue(0, 0, 1), 20);
  wb.redo(); assert.equal(wb.getValue(0, 0, 1), ERR.NAME);
});

test('JSON·셀 청크 저장/재로드는 staleCached를 정상 숫자로 복원하지 않는다', () => {
  const wb = fixture(); edit(wb);
  const data = wb.serialize();
  assert.equal(data.sheets[0].cells['0,1'].cached, undefined);
  assert.equal(data.sheets[0].cells['0,1'].staleCached, 20);
  const chunks = [...wb.cellChunks(0, 2)].flat();
  assert.equal(chunks.find(([k]) => k === '0,1')[1].staleCached, 20);
  const restored = new Workbook(JSON.parse(JSON.stringify(data)));
  assert.equal(restored.getValue(0, 0, 1), ERR.NAME);
  assert.equal(restored.getValue(1, 0, 0), ERR.NAME);
  assert.equal(restored.getCalculationStatus(0, 0, 1).status, 'stale');
});

test('XLSX 내보내기/재열기는 미지원 원문을 보존하되 오래된 숫자 대신 표준 오류값을 저장한다', () => {
  const wb = fixture(); edit(wb);
  const restored = new Workbook(readXlsx(writeXlsx(wb)).data);
  assert.equal(restored.getRaw(0, 0, 1), '=NO_SUCH_FN(A1)');
  assert.equal(restored.getValue(0, 0, 1), ERR.NAME);
  assert.equal(restored.getValue(0, 0, 2), ERR.NAME);
  assert.equal(restored.getValue(1, 0, 0), ERR.NAME);
});

test('참조 열·행·시트 삭제는 옛 숫자를 쓰지 않고 변경 없는 전체 재계산은 저장값 상태를 유지한다', () => {
  const col = fixture(); col.transact(() => col.deleteCols(0, 0, 1));
  assert.equal(col.getValue(0, 0, 0), ERR.NAME);
  assert.match(col.getRaw(0, 0, 0), /#REF!/);
  const row = new Workbook({ sheets: [{ name: 'S', fileValues: true, cells: {
    '0,0': { raw: '2' }, '1,0': { raw: '=NO_SUCH_FN(A1)', cached: 20 },
  } }] });
  row.transact(() => row.deleteRows(0, 0, 1));
  assert.equal(row.getValue(0, 0, 0), ERR.NAME);
  const sheet = new Workbook({ sheets: [{ name: 'Source', cells: { '0,0': { raw: '2' } } },
    { name: 'S', fileValues: true, cells: { '0,0': { raw: '=NO_SUCH_FN(Source!A1)', cached: 20 } } }] });
  sheet.transact(() => sheet.deleteSheet(0));
  assert.equal(sheet.getValue(0, 0, 0), ERR.NAME);
  const full = fixture(); full.invalidate();
  assert.equal(full.getValue(0, 0, 3), 99);
  assert.equal(full.getCalculationStatus(0, 0, 3).status, 'cached');
});

test('LET/LAMBDA 지역 함수·정의된 함수와 실행하지 않은 IF 분기를 미지원 결과로 오인하지 않는다', () => {
  const wb = new Workbook({ names: [{ name: 'twice', ref: '=LAMBDA(x,x*2)' }], sheets: [{ name: 'S', cells: {
    '0,0': { raw: '=LET(f,LAMBDA(x,x+1),f(2))' }, '0,1': { raw: '=twice(3)' },
    '0,2': { raw: '=IF(FALSE,NO_SUCH_FN(1),4)' }, '0,3': { raw: '=IFERROR(NO_SUCH_FN(1),0)' },
  } }] });
  assert.equal(wb.getValue(0, 0, 0), 3); assert.equal(wb.getValue(0, 0, 1), 6);
  assert.equal(wb.getValue(0, 0, 2), 4); assert.equal(wb.getValue(0, 0, 3), 0);
  assert.equal(wb.calculationIssues().total, 0);
});

test('이름 정의 변경·신규 정상 수식으로의 교체는 상태 목록과 계산 결과를 갱신한다', () => {
  const wb = new Workbook({ names: [{ name: 'amount', ref: '=2' }], sheets: [{ name: 'S', fileValues: true, cells: {
    '0,0': { raw: '=NO_SUCH_FN(amount)', cached: 20 },
  } }] });
  wb.transact(() => wb.setNames([{ name: 'amount', ref: '=3' }]));
  assert.equal(wb.getValue(0, 0, 0), ERR.NAME);
  wb.transact(() => wb.setInput(0, 0, 0, '=amount*10'));
  assert.equal(wb.getValue(0, 0, 0), 30);
  assert.equal(wb.calculationIssues().total, 0);
});

test('대량 변경의 시트 단위 fallback도 오래된 미지원 캐시를 차단한다', () => {
  const wb = fixture();
  wb.dirtyPoints(new Array(300003).fill(0));
  assert.equal(wb.getValue(0, 0, 1), ERR.NAME);
  assert.equal(wb.getValue(1, 0, 0), ERR.NAME);
});

test('새 해석불가 수식 추가/삭제는 기존 의존 그래프에도 반영된다', () => {
  const wb = fixture(); wb.getValue(0, 0, 2);
  wb.transact(() => wb.setCellData(0, 2, 2, { raw: '=A1@@', cached: 10 }));
  wb.transact(() => wb.setInput(0, 0, 0, '4'));
  assert.equal(wb.getValue(0, 2, 2), ERR.NAME);
  wb.transact(() => wb.setInput(0, 2, 2, ''));
  assert.equal(wb.getCalculationStatus(0, 2, 2), null);
});

test('사용자 정의 LAMBDA의 등록·교체·참조 입력 변경은 함수 이름 의존성을 추적한다', () => {
  const wb = new Workbook({ sheets: [{ name: 'S', fileValues: true, cells: {
    '0,0': { raw: '2' }, '0,1': { raw: '=mycalc(1)', cached: 20 },
  } }] });
  wb.transact(() => wb.setNames([{ name: 'mycalc', ref: '=LAMBDA(x,x+S!A1)' }]));
  assert.equal(wb.getValue(0, 0, 1), 3);
  wb.transact(() => wb.setInput(0, 0, 0, '4'));
  assert.equal(wb.getValue(0, 0, 1), 5);
  wb.transact(() => wb.setNames([{ name: 'mycalc', ref: '=LAMBDA(x,NO_SUCH_FN(x))' }]));
  assert.equal(wb.getValue(0, 0, 1), ERR.NAME);
  assert.equal(wb.getCalculationStatus(0, 0, 1).reason, 'unsupported-function');
});

test('joinPrev 기록과 계산값 객체를 가진 저장 캐시도 Undo 신뢰 상태를 안전하게 유지한다', () => {
  const wb = fixture();
  wb.transact(() => wb.setInput(0, 3, 0, '보고서'));
  wb.transact(() => wb.setInput(0, 0, 0, '3'), { joinPrev: true });
  assert.equal(wb.getValue(0, 0, 1), ERR.NAME);
  wb.undo(); assert.equal(wb.getValue(0, 0, 1), 20);
  wb.redo(); assert.equal(wb.getValue(0, 0, 1), ERR.NAME);
  const errorBook = new Workbook({ sheets: [{ name: 'S', fileValues: true, cells: {
    '0,0': { raw: '=NO_SUCH_FN(1)', cached: { error: '#DIV/0!' } },
  } }] });
  const state = errorBook.getCalculationStatus(0, 0, 0);
  state.savedValue.error = '#N/A';
  assert.equal(errorBook.getValue(0, 0, 0), ERR.DIV0);
});

test('수동 계산은 대기 상태를 알리고 F9 이후에는 미지원 저장값과 의존 합계를 차단한다', () => {
  const wb = fixture();
  wb.manualCalc = true;
  edit(wb);
  assert.equal(wb.getValue(0, 0, 1), 20);
  assert.equal(wb.getCalculationStatus(0, 0, 1).pending, true);
  assert.equal(wb.calculationIssues().items[0].pending, true);
  wb.calculateNow();
  assert.equal(wb.getValue(0, 0, 1), ERR.NAME);
  assert.equal(wb.getValue(1, 0, 0), ERR.NAME);
  assert.equal(wb.getCalculationStatus(0, 0, 1).pending, undefined);
});

test('보고서의 정상 함수값과 서식 변경은 캐시 신뢰 문제 목록을 만들지 않는다', () => {
  const wb = new Workbook({ sheets: [{ name: 'S', fileValues: true, cells: {
    '0,0': { raw: '2' }, '0,1': { raw: '=A1*3', cached: 6 },
  } }] });
  wb.transact(() => wb.setStyle(0, 0, 1, { color: '#ff0000' }));
  assert.equal(wb.getValue(0, 0, 1), 6);
  assert.equal(wb.calculationIssues().total, 0);
  wb.transact(() => wb.setInput(0, 0, 0, '4'));
  assert.equal(wb.getValue(0, 0, 1), 12);
  assert.equal(wb.getCalculationStatus(0, 0, 1).status, 'calculated');
  assert.equal(wb.calculationIssues().total, 0);
});

test('아직 계산하지 않은 정의된 LAMBDA 내부의 미지원 함수도 저장값 검토 목록에 포함한다', () => {
  const wb = new Workbook({ names: [{ name: 'mycalc', ref: '=LAMBDA(x,NO_SUCH_FN(x))' }], sheets: [{ name: 'S', fileValues: true, cells: {
    '0,0': { raw: '=mycalc(1)', cached: 20 },
  } }] });
  const list = wb.calculationIssues();
  assert.equal(list.total, 1);
  assert.deepEqual(list.items[0].functions, ['NO_SUCH_FN']);
  assert.deepEqual(list.items[0].viaNames, ['MYCALC']);
  assert.equal(list.items[0].status, 'cached');
});
