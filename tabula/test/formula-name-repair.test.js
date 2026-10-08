import { test } from 'node:test';
import assert from 'node:assert/strict';
import { Workbook } from '../src/workbook.js';
import { FUNCS, unknownFunctions } from '../src/formula.js';
import { readXlsx, writeXlsx } from '../src/xlsx.js';
import { unzip } from '../src/zip.js';
import { suggestFunctionName, functionNameRepairPlan, applyFunctionNameRepairs } from '../src/formula-name-repair.js';

const typo = '=IFERROR(IFS(A1="red",SUMIFSㅋ(C1:C3,B1:B3,"red"),A1="blue",SUMIFS(C1:C3,B1:B3,"blue"),TRUE,SUM(C1:C3)),0)';
function fixture() {
  return new Workbook({ sheets: [{ name: '입력', fileValues: true, cells: {
    '0,0': { raw: 'red' }, '0,1': { raw: 'red' }, '1,1': { raw: 'blue' }, '2,1': { raw: 'red' },
    '0,2': { raw: '10' }, '1,2': { raw: '20' }, '2,2': { raw: '30' },
    '0,3': { raw: typo, cached: 0, style: { numFmt: 'text', bold: true, fill: '#ffeecc' }, fx: true, comment: '메모', link: '#입력!A1' },
    '1,3': { raw: '=SUM(D1,5)', cached: 5 },
    '0,4': { raw: '=UNKNOWN_OTHER(1)', cached: 99 },
  } }, { name: '요약', fileValues: true, cells: { '0,0': { raw: '=입력!D1*2', cached: 0 } } }] });
}
const values = book => [book.getValue(0, 0, 3), book.getValue(0, 1, 3), book.getValue(1, 0, 0)];

test('알려진 함수 뒤의 한글 한 글자만 정확한 함수 이름 후보로 제안한다', () => {
  for (const name of ['SUMIFSㅋ', 'sumifsㅋ', 'SUMIFS가', 'STDEV.Sㄱ', 'XLOOKUPᄀ']) assert.equal(suggestFunctionName(name), name.slice(0, -1).toUpperCase());
  for (const name of ['SUMIFS', 'SUMIFSㅋㅋ', 'SUMIFSX', 'SUMIFSΣ', 'SUMIFSㅋ ', ' SUMIFSㅋ', 'MISSINGㅋ', 'constructorㅋ', '_xlfn.SUMIFSㅋ', '', null]) assert.equal(suggestFunctionName(name), null, String(name));
  assert.equal(FUNCS['SUMIFSㅋ'], undefined);
  assert.deepEqual(unknownFunctions('SUMIFSㅋ(A1,A1,1)'), ['SUMIFSㅋ']);
});

test('검토 계획은 문서 값·버전·계산 캐시·Undo와 공유 AST를 변경하지 않는다', () => {
  const book = fixture(), before = JSON.stringify(book.serialize()), version = book.version, ast = book.getCell(0, 0, 3).ast;
  const plan = functionNameRepairPlan(book);
  assert.equal(JSON.stringify(book.serialize()), before);
  assert.equal(book.version, version);
  assert.equal(book.undoStack.length, 0);
  assert.equal(book.getCell(0, 0, 3).ast, ast);
  assert.deepEqual(values(book), [0, 5, 0]);
  assert.equal(plan.cellCount, 1); assert.equal(plan.templateCount, 0); assert.equal(plan.total, 1);
  assert.deepEqual(plan.functions, [{ from: 'SUMIFSㅋ', to: 'SUMIFS', cells: 1, templates: 0 }]);
  assert.equal(plan.cells[0].before, typo);
  assert.equal(plan.cells[0].after, typo.replace('SUMIFSㅋ', 'SUMIFS'));
  assert.deepEqual(plan.cells[0].functions, [{ from: 'SUMIFSㅋ', to: 'SUMIFS' }]);
  assert.ok(Object.isFrozen(plan)); assert.ok(Object.isFrozen(plan.cells[0]));
});

test('일반 XLSX 열기와 저장은 잘못된 이름을 그대로 보존하며 별칭을 등록하지 않는다', () => {
  const book = fixture(), first = readXlsx(writeXlsx(book)), restored = new Workbook(first.data);
  assert.equal(restored.getRaw(0, 0, 3), typo);
  const warning = first.warnings.find(w => w.includes('확인이 필요한 함수 이름'));
  assert.match(warning, /수식 2개/); assert.match(warning, /SUMIFSㅋ/); assert.match(warning, /UNKNOWN_OTHER/);
  assert.match(warning, /IFERROR와 IF·IFS의 조건 처리는 그대로 계산/);
  const second = new Workbook(readXlsx(writeXlsx(restored)).data);
  assert.equal(second.getRaw(0, 0, 3), typo);
  assert.deepEqual(values(second), [0, 5, 0]);
  assert.equal(FUNCS['SUMIFSㅋ'], undefined);
});

test('명시적인 수정은 선택된 오류 분기와 의존 결과를 재계산하고 한 번의 Undo/Redo로 복원한다', () => {
  const book = fixture(), style = { ...book.getCell(0, 0, 3).style };
  assert.deepEqual(values(book), [0, 5, 0]);
  assert.deepEqual(applyFunctionNameRepairs(book, functionNameRepairPlan(book)), { cellCount: 1, templateCount: 0, total: 1 });
  assert.deepEqual(values(book), [40, 45, 80]);
  const cell = book.getCell(0, 0, 3);
  assert.deepEqual(cell.style, style); assert.equal(cell.comment, '메모'); assert.equal(cell.link, '#입력!A1'); assert.equal(cell.fx, true);
  assert.equal(cell.cached, undefined);
  assert.equal(book.getRaw(0, 0, 4), '=UNKNOWN_OTHER(1)'); assert.equal(book.getValue(0, 0, 4), 99);
  assert.equal(book.undoStack.length, 1);
  book.undo(); assert.equal(book.getRaw(0, 0, 3), typo); assert.deepEqual(values(book), [0, 5, 0]);
  book.redo(); assert.equal(book.getRaw(0, 0, 3), typo.replace('SUMIFSㅋ', 'SUMIFS')); assert.deepEqual(values(book), [40, 45, 80]);
  assert.equal(FUNCS['SUMIFSㅋ'], undefined);
});

test('수정 뒤 입력 변경으로 정상 분기·전체 분기·합계가 계속 재계산되고 XLSX 재열기에 남는다', () => {
  const book = fixture(); applyFunctionNameRepairs(book, functionNameRepairPlan(book));
  book.transact(() => book.setInput(0, 0, 0, 'blue')); assert.deepEqual(values(book), [20, 25, 40]);
  book.transact(() => book.setInput(0, 0, 0, 'mixed')); assert.deepEqual(values(book), [60, 65, 120]);
  book.transact(() => book.setInput(0, 0, 2, '100')); assert.deepEqual(values(book), [150, 155, 300]);
  book.undo(); assert.deepEqual(values(book), [60, 65, 120]); book.redo(); assert.deepEqual(values(book), [150, 155, 300]);
  const parsed = readXlsx(writeXlsx(book)), restored = new Workbook(parsed.data);
  assert.ok(!restored.getRaw(0, 0, 3).includes('SUMIFSㅋ'));
  assert.deepEqual(values(restored), [150, 155, 300]);
  assert.equal(functionNameRepairPlan(restored).total, 0);
  restored.transact(() => restored.setInput(0, 0, 0, 'red')); assert.deepEqual(values(restored), [130, 135, 260]);
});

test('배열 함수 이름 수정은 옛 스칼라 캐시를 버리고 분산 결과와 의존 합계를 계산한다', () => {
  const raw = '=IFERROR(FILTERㅋ(A1:A3,A1:A3>1),0)';
  const book = new Workbook({ sheets: [{ name: 'S', fileValues: true, cells: {
    '0,0': { raw: '1' }, '1,0': { raw: '2' }, '2,0': { raw: '3' },
    '0,3': { raw, cached: 0 }, '0,5': { raw: '=SUM(D1:D2)', cached: 0 },
  } }] });
  applyFunctionNameRepairs(book, functionNameRepairPlan(book));
  assert.deepEqual([book.getValue(0, 0, 3), book.getValue(0, 1, 3), book.getValue(0, 0, 5)], [2, 3, 5]);
  assert.equal(book.getCell(0, 0, 3).cached, undefined);
  book.undo(); assert.equal(book.getRaw(0, 0, 3), raw);
  assert.deepEqual([book.getValue(0, 0, 3), book.getValue(0, 1, 3), book.getValue(0, 0, 5)], [0, null, 0]);
  book.redo(); assert.deepEqual([book.getValue(0, 0, 3), book.getValue(0, 1, 3), book.getValue(0, 0, 5)], [2, 3, 5]);
});

test('함수 토큰만 수정하며 문자열·시트명·구조적 열 이름과 다른 누락 함수는 보존한다', () => {
  const raw = '=IFERROR(SUMIFSㅋ(A1:A2,A1:A2,1),0)&"SUMIFSㅋ(A1)"&\'SUMIFSㅋ\'!A1&SUMIFSㅋ표[SUMIFSㅋ]';
  const book = new Workbook({ sheets: [{ name: 'S', cells: { '0,0': { raw } } }, { name: 'SUMIFSㅋ', cells: {} }] });
  const plan = functionNameRepairPlan(book);
  assert.equal(plan.cells[0].after, raw.replace('SUMIFSㅋ(A1:A2', 'SUMIFS(A1:A2'));
  const multi = new Workbook({ sheets: [{ name: 'S', cells: { '0,0': { raw: '=SUMIFSㅋ(A1,A1,1)+SUMIFSㅋ(A2,A2,1)+MISSING(1)' } } }] });
  const target = functionNameRepairPlan(multi);
  assert.equal(target.cellCount, 1); assert.equal(target.functions[0].cells, 1);
  assert.equal(target.cells[0].after, '=SUMIFS(A1,A1,1)+SUMIFS(A2,A2,1)+MISSING(1)');
});

test('전역·시트 범위 정의 함수와 LET/LAMBDA 지역 함수를 오타로 수정하지 않는다', () => {
  const book = new Workbook({ names: [{ name: 'SUMIFSㅋ', ref: '=LAMBDA(x,x)' }, { name: 'ROUNDㅋ', sheet: 'Local', ref: '=LAMBDA(x,x)' }], sheets: [
    { name: 'Local', cells: { '0,0': { raw: '=SUMIFSㅋ(1)+ROUNDㅋ(2)' }, '1,0': { raw: '=LET(SUMㅋ,LAMBDA(x,x),SUMㅋ(3))' }, '2,0': { raw: '=LAMBDA(SUMㅋ,SUMㅋ(3))(LAMBDA(x,x))' } } },
    { name: 'Other', cells: { '0,0': { raw: '=SUMIFSㅋ(1)+ROUNDㅋ(2)' } } },
  ] });
  const plan = functionNameRepairPlan(book);
  assert.equal(plan.cellCount, 1); assert.equal(plan.cells[0].si, 1);
  assert.equal(plan.cells[0].after, '=SUMIFSㅋ(1)+ROUND(2)');
  const unreadable = new Workbook({ sheets: [{ name: 'S', cells: { '0,0': { raw: '=LET(SUMIFSㅋ,LAMBDA(x,x),SUMIFSㅋ(3)' }, '1,0': { raw: '=SUMIFSㅋ(A1' } } }] });
  assert.equal(functionNameRepairPlan(unreadable).total, 0, '구문을 읽지 못한 지역 이름은 추정하지 않음');
});

test('같은 이름의 바깥 오타 호출만 수정하고 LET/LAMBDA 지역 함수 호출의 위치와 결과를 보존한다', () => {
  for (const [raw, after] of [
    ['=IFERROR(SUMㅋ(1),0)+LET(SUMㅋ,LAMBDA(x,x*10),SUMㅋ(2))', '=IFERROR(SUM(1),0)+LET(SUMㅋ,LAMBDA(x,x*10),SUMㅋ(2))'],
    ['=IFERROR(SUMㅋ(1),0)+LAMBDA(SUMㅋ,SUMㅋ(2))(LAMBDA(x,x*10))', '=IFERROR(SUM(1),0)+LAMBDA(SUMㅋ,SUMㅋ(2))(LAMBDA(x,x*10))'],
    ['=LET(a,IFERROR(SUMㅋ(1),0),SUMㅋ,LAMBDA(x,x*10),a+SUMㅋ(2))', '=LET(a,IFERROR(SUM(1),0),SUMㅋ,LAMBDA(x,x*10),a+SUMㅋ(2))'],
    ['=SUM(LET(SUMㅋ,LAMBDA(x,x*10),SUMㅋ(2)),IFERROR(SUMㅋ(1),0))', '=SUM(LET(SUMㅋ,LAMBDA(x,x*10),SUMㅋ(2)),IFERROR(SUM(1),0))'],
  ]) {
    const book = new Workbook({ sheets: [{ name: 'S', cells: { '0,0': { raw } } }] });
    assert.equal(book.getValue(0, 0, 0), 20, raw);
    const plan = functionNameRepairPlan(book);
    assert.equal(plan.cellCount, 1); assert.equal(plan.cells[0].after, after);
    applyFunctionNameRepairs(book, plan);
    assert.equal(book.getValue(0, 0, 0), 21, after);
    book.undo(); assert.equal(book.getRaw(0, 0, 0), raw); assert.equal(book.getValue(0, 0, 0), 20);
    book.redo(); assert.equal(book.getRaw(0, 0, 0), after); assert.equal(book.getValue(0, 0, 0), 21);
    const restored = new Workbook(readXlsx(writeXlsx(book)).data);
    assert.equal(restored.getRaw(0, 0, 0), after); assert.equal(restored.getValue(0, 0, 0), 21);
    assert.equal(functionNameRepairPlan(restored).total, 0, '지역 이름만 남으면 더 이상 수정하지 않음');
  }
});

test('버전·문서·원문·이름 충돌이 바뀐 계획은 아무 대상도 적용하지 않는다', () => {
  const book = fixture(), plan = functionNameRepairPlan(book);
  book.transact(() => book.setInput(0, 0, 0, 'blue'));
  assert.throws(() => applyFunctionNameRepairs(book, plan), /통합 문서가 바뀌었습니다/);
  assert.equal(book.getRaw(0, 0, 3), typo);
  assert.throws(() => applyFunctionNameRepairs(fixture(), functionNameRepairPlan(book)), /통합 문서가 바뀌었습니다/);
  assert.throws(() => applyFunctionNameRepairs(book, { ...functionNameRepairPlan(book) }), /통합 문서가 바뀌었습니다/);
  const rawChanged = fixture(), rawPlan = functionNameRepairPlan(rawChanged);
  rawChanged.getCell(0, 0, 3).raw += '+1';
  assert.throws(() => applyFunctionNameRepairs(rawChanged, rawPlan), /수정 대상 수식이 바뀌었습니다/);
  const collision = fixture(), namePlan = functionNameRepairPlan(collision);
  collision.names.push({ name: 'SUMIFSㅋ', ref: '=LAMBDA(x,x)' });
  assert.throws(() => applyFunctionNameRepairs(collision, namePlan), /정의된 이름이 바뀌었습니다/);
  assert.equal(collision.getRaw(0, 0, 3), typo); assert.equal(collision.undoStack.length, 0);
});

test('보호된 시트나 외부 시트가 포함되면 전체 계획을 거절하여 부분 수정하지 않는다', () => {
  for (const props of [{ protect: { on: true } }, { external: 1 }]) {
    const book = fixture(); book.sheets.push({ ...book.sheets[0], name: '제한', ...props });
    const plan = functionNameRepairPlan(book), before = book.getRaw(0, 0, 3);
    assert.equal(plan.blocked.length, 1);
    assert.throws(() => applyFunctionNameRepairs(book, plan), /보호되거나 외부/);
    assert.equal(book.getRaw(0, 0, 3), before); assert.equal(book.undoStack.length, 0);
  }
  const changed = fixture(), plan = functionNameRepairPlan(changed);
  changed.sheets[0].protect = { on: true };
  assert.throws(() => applyFunctionNameRepairs(changed, plan), /보호되거나 외부/);
  assert.equal(changed.getRaw(0, 0, 3), typo);
  const final = fixture(), finalPlan = functionNameRepairPlan(final);
  final.props = { markedFinal: true };
  assert.throws(() => applyFunctionNameRepairs(final, finalPlan), /최종본/);
  assert.equal(final.getRaw(0, 0, 3), typo); assert.equal(final.undoStack.length, 0);
});

function tableFixture() {
  const raw = '=IFERROR(SUMIFSㅋ(T[금액],T[분류],T[[#This Row],[분류]]),0)';
  return { raw, book: new Workbook({ sheets: [{ name: '표', fileValues: true, tables: [{ id: 't1', name: 'T', r1: 0, c1: 0, r2: 2, c2: 2, header: true, totals: false, style: 'None', calculatedColumnFormulas: { 2: raw }, custom: { untouched: true } }], cells: {
    '0,0': { raw: '분류' }, '0,1': { raw: '금액' }, '0,2': { raw: '합계' },
    '1,0': { raw: 'red' }, '1,1': { raw: '10' }, '1,2': { raw, cached: 0 },
    '2,0': { raw: 'red' }, '2,1': { raw: '30' }, '2,2': { raw, cached: 0 },
    '3,3': { raw: '=UNKNOWN_OTHER(1)', cached: 99 },
  } }] }) };
}

test('표 계산 열 템플릿과 실제 수식을 함께 수정하고 Undo/Redo 및 저장본 XML에 오래된 이름을 남기지 않는다', () => {
  const { book, raw } = tableFixture(), ast = book.getCell(0, 1, 2).ast, snapshot = JSON.stringify(ast);
  assert.equal(book.getCell(0, 2, 2).ast, ast);
  const plan = functionNameRepairPlan(book);
  assert.equal(plan.cellCount, 2); assert.equal(plan.templateCount, 1);
  assert.deepEqual(plan.functions, [{ from: 'SUMIFSㅋ', to: 'SUMIFS', cells: 2, templates: 1 }]);
  applyFunctionNameRepairs(book, plan);
  assert.equal(JSON.stringify(ast), snapshot, '기존 공유 AST 불변');
  assert.equal(book.sheets[0].tables[0].calculatedColumnFormulas[2], raw.replace('SUMIFSㅋ', 'SUMIFS'));
  assert.deepEqual([book.getValue(0, 1, 2), book.getValue(0, 2, 2)], [40, 40]);
  assert.equal(book.getValue(0, 3, 3), 99, '무관한 미지원 저장값 유지');
  assert.deepEqual(book.sheets[0].tables[0].custom, { untouched: true });
  const files = unzip(writeXlsx(book)), tableXml = new TextDecoder().decode(files['xl/tables/table1.xml']);
  assert.match(tableXml, /<calculatedColumnFormula>IFERROR\(SUMIFS\(/);
  assert.ok(!tableXml.includes('SUMIFSㅋ'));
  book.undo(); assert.equal(book.getRaw(0, 1, 2), raw); assert.equal(book.sheets[0].tables[0].calculatedColumnFormulas[2], raw);
  assert.deepEqual([book.getValue(0, 1, 2), book.getValue(0, 2, 2)], [0, 0]); assert.equal(book.getValue(0, 3, 3), 99);
  book.redo(); assert.deepEqual([book.getValue(0, 1, 2), book.getValue(0, 2, 2)], [40, 40]);
  assert.equal(book.sheets[0].tables[0].calculatedColumnFormulas[2], raw.replace('SUMIFSㅋ', 'SUMIFS'));
});

test('표 템플릿의 원문 변경·이름 충돌·보호도 사전 검사하고 다른 템플릿은 보존한다', () => {
  const { book, raw } = tableFixture();
  book.sheets[0].tables[0].calculatedColumnFormulas[1] = 'SUMIFSㅋ(A1,A1,1)';
  const plan = functionNameRepairPlan(book);
  book.sheets[0].tables[0].calculatedColumnFormulas[2] += '+1';
  assert.throws(() => applyFunctionNameRepairs(book, plan), /수정 대상 수식이 바뀌었습니다/);
  assert.equal(book.getRaw(0, 1, 2), raw); assert.equal(book.undoStack.length, 0);
  const named = tableFixture().book; named.names.push({ name: 'SUMIFSㅋ', sheet: '표', ref: '=LAMBDA(x,x)' });
  assert.equal(functionNameRepairPlan(named).total, 0);
  const onlyTemplate = tableFixture().book;
  onlyTemplate.transact(() => { onlyTemplate.setInput(0, 1, 2, '1'); onlyTemplate.setInput(0, 2, 2, '2'); });
  onlyTemplate.sheets[0].tables[0].calculatedColumnFormulas[1] = '=SUM(T[금액])';
  const target = functionNameRepairPlan(onlyTemplate);
  assert.equal(target.cellCount, 0); assert.equal(target.templateCount, 1);
  applyFunctionNameRepairs(onlyTemplate, target);
  assert.equal(onlyTemplate.sheets[0].tables[0].calculatedColumnFormulas[1], '=SUM(T[금액])');
  assert.deepEqual([onlyTemplate.getValue(0, 1, 2), onlyTemplate.getValue(0, 2, 2)], [1, 2]);
  onlyTemplate.undo(); assert.equal(onlyTemplate.sheets[0].tables[0].calculatedColumnFormulas[2], raw);
});

test('파일 형식 템플릿의 접두사와 등호 유무는 함수 토큰 수정 뒤에도 유지한다', () => {
  const book = tableFixture().book;
  book.sheets[0].tables[0].calculatedColumnFormulas[2] = '_xlfn.XLOOKUPㅋ(A1,B1:B2,C1:C2)';
  const plan = functionNameRepairPlan(book), entry = plan.tables[0];
  assert.equal(entry.after, '_xlfn.XLOOKUP(A1,B1:B2,C1:C2)');
});

test('가져오기 경고는 채우기 수식 모양을 재사용해도 실제 함수 이름과 셀 개수를 유지한다', () => {
  const book = new Workbook({ sheets: [{ name: 'S', fileValues: true, cells: {
    '0,0': { raw: '=IFERROR(SUMIFSㅋ(B1:B3,B1:B3,1),0)', cached: 0 },
    '1,0': { raw: '=IFERROR(SUMIFSㅋ(B2:B4,B2:B4,1),0)', cached: 0 },
    '2,0': { raw: '=MISSING_SECOND(B3)', cached: 42 },
    '3,0': { raw: '=MISSING_SECOND(B4)', cached: { error: '#NAME?' } },
  } }] });
  const result = readXlsx(writeXlsx(book)), warning = result.warnings.find(w => w.includes('확인이 필요한 함수 이름'));
  assert.match(warning, /수식 3개/);
  assert.match(warning, /\(SUMIFSㅋ, MISSING_SECOND\)/);
  const restored = new Workbook(result.data);
  for (let r = 0; r < 4; r++) assert.equal(restored.getRaw(0, r, 0), book.getRaw(0, r, 0));
  assert.equal(restored.getCalculationStatus(0, 3, 0).status, 'source-error');
});

test('가져오기 경고의 함수 이름 목록은 제한되며 문자열·정의 함수·원본 이름 오류는 목록에 섞이지 않는다', () => {
  const missing = Array.from({ length: 10 }, (_, i) => `UNKNOWN_${String.fromCharCode(65 + i)}`), cells = {};
  missing.forEach((fn, r) => { cells[`${r},0`] = { raw: `=${fn}(B${r + 1})`, cached: 1 }; });
  cells['10,0'] = { raw: '=IF(TRUE,"STRING_ONLY_FN(1)",0)', cached: 'STRING_ONLY_FN(1)' };
  cells['11,0'] = { raw: '=defined_fn(2)', cached: 4 };
  cells['12,0'] = { raw: '=SOURCE_NAME_ERROR(1)', cached: { error: '#NAME?' } };
  const book = new Workbook({ names: [{ name: 'defined_fn', ref: '=LAMBDA(x,x*2)' }], sheets: [{ name: 'S', fileValues: true, cells }] });
  const result = readXlsx(writeXlsx(book)), warning = result.warnings.find(w => w.includes('확인이 필요한 함수 이름'));
  assert.match(warning, /수식 10개/);
  for (const fn of missing.slice(0, 8)) assert.ok(warning.includes(fn));
  for (const fn of [...missing.slice(8), 'STRING_ONLY_FN', 'DEFINED_FN', 'SOURCE_NAME_ERROR']) assert.ok(!warning.includes(fn), fn);
  assert.match(warning, / 등\)/);
});
