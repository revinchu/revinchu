import { test } from 'node:test';
import assert from 'node:assert/strict';
import { Workbook } from '../src/workbook.js';
import { resolveWorkbookLink, rewriteWorkbookLink } from '../src/hyperlink.js';
import { MAX_ROWS, MAX_COLS } from '../src/formula.js';

function book() {
  return new Workbook({ sheets: [{ name: '홈', cells: {} }, { name: "O'Brien 보고서", cells: {} }, { name: 'Report', cells: {} }, { name: '숨김', state: 'hidden', cells: {} }, { name: '매우 숨김', state: 'veryHidden', cells: {} }] });
}
const rg = (r1, c1, r2 = r1, c2 = c1) => ({ r1, c1, r2, c2 });
const result = (sheet, range) => ({ sheet, range });

test('내부 주소는 인용 공백·작은따옴표·대소문자·절대 주소·역순 범위를 해석한다', () => {
  const wb = book();
  assert.deepEqual(resolveWorkbookLink(wb, 0, "#'O''Brien 보고서'!$B$2:$D$9"), result(1, rg(1, 1, 8, 3)));
  assert.deepEqual(resolveWorkbookLink(wb, 0, '#rEpOrT!d9:B2'), result(2, rg(1, 1, 8, 3)));
  assert.deepEqual(resolveWorkbookLink(wb, 2, '#A1'), result(2, rg(0, 0)));
  assert.deepEqual(resolveWorkbookLink(wb, 0, 'Report!XFD1048576'), result(2, rg(1048575, 16383)));
});

test('전체 행·열 참조는 유효 범위만 허용하며 사용 범위로 잘라내지 않는다', () => {
  const wb = book();
  assert.deepEqual(resolveWorkbookLink(wb, 0, '#$D:$B'), result(0, rg(0, 1, MAX_ROWS - 1, 3)));
  assert.deepEqual(resolveWorkbookLink(wb, 0, '#$4:$2'), result(0, rg(1, 0, 3, MAX_COLS - 1)));
  for (const text of ['#XFE1', '#A0', '#A20000001', '#XFE:XFE', '#0:4', '#1:20000001']) assert.ok(resolveWorkbookLink(wb, 0, text).error, text);
});

test('이름 범위는 현재 시트 우선·통합 문서 이름·명시 시트의 이름을 구분한다', () => {
  const wb = book();
  wb.names = [
    { name: '맨위로', ref: '=Report!$C$3', sheet: null },
    { name: '맨위로', ref: '=$B$2', sheet: '홈' },
    { name: '맨위로', ref: '=$D$4', sheet: 'Report' },
    { name: '별칭', ref: '=맨위로', sheet: null },
    { name: '전역', ref: "='O''Brien 보고서'!A2:C4", sheet: null },
  ];
  assert.deepEqual(resolveWorkbookLink(wb, 0, '#맨위로'), result(0, rg(1, 1)));
  assert.deepEqual(resolveWorkbookLink(wb, 1, '#맨위로'), result(2, rg(2, 2)));
  assert.deepEqual(resolveWorkbookLink(wb, 0, '#REPORT!맨위로'), result(2, rg(3, 3)));
  assert.deepEqual(resolveWorkbookLink(wb, 0, '#별칭'), result(0, rg(1, 1)));
  assert.deepEqual(resolveWorkbookLink(wb, 0, '#전역'), result(1, rg(1, 0, 3, 2)));
});

test('정적 이름 재귀는 수식을 평가하지 않고 문서·이름·실행 취소 기록을 바꾸지 않는다', () => {
  const wb = book();
  wb.names = [{ name: '첫이름', ref: '=둘째이름' }, { name: '둘째이름', ref: '=Report!A1' }, { name: '동적', ref: '=INDIRECT("Report!A1")' }, { name: '통신', ref: '=IMPORTRANGE("개인자료","A1")' }, { name: '상수', ref: '=5' }, { name: '여러영역', ref: '=A1,C3' }];
  wb.nameValue = () => { throw Error('이름 수식 평가 금지'); };
  wb.getValue = () => { throw Error('셀 계산 금지'); };
  const before = JSON.stringify(wb.serialize()), version = wb.version;
  assert.deepEqual(resolveWorkbookLink(wb, 0, '#첫이름'), result(2, rg(0, 0)));
  for (const name of ['동적', '통신', '상수', '여러영역', '없는이름']) assert.ok(resolveWorkbookLink(wb, 0, '#' + name).error, name);
  assert.equal(JSON.stringify(wb.serialize()), before);
  assert.equal(wb.version, version);
  assert.equal(wb.undoStack.length, 0);
  assert.ok(wb.names.every(n => !('_ast' in n)));
});

test('순환·과도한 이름 재귀는 끝나며 한국어 오류를 반환한다', () => {
  const wb = book();
  wb.names = [{ name: '첫', ref: '=둘' }, { name: '둘', ref: '=첫' }];
  assert.match(resolveWorkbookLink(wb, 0, '#첫').error, /반복/);
  wb.names = Array.from({ length: 70 }, (_, i) => ({ name: '이름_' + i, ref: i === 69 ? '=A1' : '=이름_' + (i + 1) }));
  assert.match(resolveWorkbookLink(wb, 0, '#이름_0').error, /단계/);
});

test('URL 인코딩은 리터럴 주소 실패 후 한 번만 해석한다', () => {
  const wb = book();
  assert.deepEqual(resolveWorkbookLink(wb, 0, '#' + encodeURIComponent("'O''Brien 보고서'!$A$9")), result(1, rg(8, 0)));
  wb.sheets.push({ name: '%ED%99%88' });
  assert.deepEqual(resolveWorkbookLink(wb, 0, "#'%ED%99%88'!A1"), result(5, rg(0, 0)));
  wb.names = [{ name: '맨위로', ref: '=A1' }];
  assert.deepEqual(resolveWorkbookLink(wb, 0, '#' + encodeURIComponent('맨위로')), result(0, rg(0, 0)));
  assert.deepEqual(resolveWorkbookLink(wb, 0, '#%25ED%2599%2588!A1'), result(5, rg(0, 0)), '한 번만 해석하여 인코딩처럼 보이는 실제 시트 이름으로 이동');
  assert.ok(resolveWorkbookLink(wb, 0, '#%ZZ!A1').error);
});

test('제어문자·외부 통합 문서·URL·잘못된 참조와 숨김 시트를 거부한다', () => {
  const wb = book();
  for (const target of [null, '', '#', '#A1\n', '#Report!A1\u0000', '#Report!A1\u0085', '#Report!A1%0A', '#[Book.xlsx]Report!A1', "#'[1]Report'!A1", '#없는시트!A1', '#A1:B2:C3', '#Report!A1+1', '#Sheet1:Sheet3!A1', 'https://example.com/#A1', 'javascript:alert(1)', '#숨김!A1', "#'매우 숨김'!A1"]) {
    const out = resolveWorkbookLink(wb, 0, target);
    assert.ok(out.error, String(target)); assert.match(out.error, /[가-힣]/);
  }
  assert.ok(resolveWorkbookLink(wb, -1, '#A1').error);
});

function linkedBook() {
  return new Workbook({ names: [{ name: '위치', ref: "='O''Brien 보고서'!$A$5" }], sheets: [
    { name: '홈', fileValues: true, cells: {
      '0,0': { raw: '=1+1', cached: 42, link: "#'O''Brien 보고서'!$A$5" },
      '0,1': { raw: '외부', link: 'https://example.com/#Report!A5' },
      '0,2': { raw: '이름', link: '#위치' },
    }, shapes: [{ id: 'button', kind: 'rect', hyperlink: { target: "#'O''Brien 보고서'!$A$5", tooltip: '이동', action: 'ppaction://hlinksldjump' } }, { id: 'group', kind: 'group', groupItems: [{ id: 'child', kind: 'rect', hyperlink: { target: "#'O''Brien 보고서'!B5:C9" } }] }], images: [{ id: 'picture', hyperlink: { target: '#%27O%27%27Brien%20%EB%B3%B4%EA%B3%A0%EC%84%9C%27!D5' } }] },
    { name: "O'Brien 보고서", cells: { '0,0': { raw: '자체', link: '#A5' } }, shapes: [{ id: 'same', hyperlink: { target: '#A5' } }] },
    { name: '다른 시트', cells: { '0,0': { raw: '자체', link: '#A5' } } },
  ] });
}

test('시트 이름 변경은 셀·도형·그림·그룹 내부 링크를 함께 갱신하고 한 번에 실행 취소한다', () => {
  const wb = linkedBook(), before = wb.serialize();
  const oldShape = wb.sheets[0].shapes[0];
  wb.transact(() => assert.equal(wb.renameSheet(1, '새 보고서'), true));
  assert.equal(wb.getCell(0, 0, 0).link, "#'새 보고서'!$A$5");
  assert.equal(wb.sheets[0].shapes[0].hyperlink.target, "#'새 보고서'!$A$5");
  assert.equal(wb.sheets[0].shapes[0].hyperlink.tooltip, '이동');
  assert.equal(wb.sheets[0].shapes[0].hyperlink.action, 'ppaction://hlinksldjump');
  assert.equal(wb.sheets[0].shapes[1].groupItems[0].hyperlink.target, "#'새 보고서'!B5:C9");
  assert.equal(wb.sheets[0].images[0].hyperlink.target, "#'새 보고서'!D5");
  assert.equal(oldShape.hyperlink.target, "#'O''Brien 보고서'!$A$5", '원본 객체를 변이하지 않음');
  assert.equal(wb.getValue(0, 0, 0), 42, '링크 메타데이터 변경은 저장된 수식 결과를 재계산하지 않음');
  assert.equal(wb.getCell(0, 0, 1).link, before.sheets[0].cells['0,1'].link);
  assert.equal(wb.getCell(0, 0, 2).link, '#위치');
  assert.deepEqual(resolveWorkbookLink(wb, 0, '#위치'), result(1, rg(4, 0)));
  const after = wb.serialize();
  assert.equal(wb.undoStack.length, 1);
  wb.undo(); assert.deepEqual(wb.serialize(), before);
  wb.redo(); assert.deepEqual(wb.serialize(), after);
});

test('행 삽입은 호스트 시트를 구분하고 절대 참조와 이름의 대상을 함께 이동한다', () => {
  const wb = linkedBook(), before = wb.serialize();
  wb.transact(() => wb.insertRows(1, 2, 3));
  assert.equal(wb.getCell(0, 0, 0).link, "#'O''Brien 보고서'!$A$8");
  assert.equal(wb.sheets[0].shapes[1].groupItems[0].hyperlink.target, "#'O''Brien 보고서'!B8:C12");
  assert.equal(wb.getCell(1, 0, 0).link, '#A8');
  assert.equal(wb.sheets[1].shapes[0].hyperlink.target, '#A8');
  assert.equal(wb.getCell(2, 0, 0).link, '#A5');
  assert.equal(wb.getCell(0, 0, 2).link, '#위치');
  assert.deepEqual(resolveWorkbookLink(wb, 0, '#위치'), result(1, rg(7, 0)));
  const after = wb.serialize();
  wb.undo(); assert.deepEqual(wb.serialize(), before);
  wb.redo(); assert.deepEqual(wb.serialize(), after);
});

test('열 삽입·범위 일부 삭제·대상 전체 삭제는 링크를 정확히 바꾸고 복원한다', () => {
  const wb = linkedBook();
  wb.transact(() => wb.insertCols(1, 0, 2));
  assert.equal(wb.sheets[0].shapes[0].hyperlink.target, "#'O''Brien 보고서'!$C$5");
  wb.undo();
  wb.transact(() => wb.deleteRows(1, 4, 2));
  assert.match(wb.sheets[0].shapes[0].hyperlink.target, /#REF!/);
  assert.ok(resolveWorkbookLink(wb, 0, wb.sheets[0].shapes[0].hyperlink.target).error);
  assert.equal(wb.sheets[0].shapes[1].groupItems[0].hyperlink.target, "#'O''Brien 보고서'!B5:C7");
  wb.undo();
  assert.deepEqual(resolveWorkbookLink(wb, 0, wb.sheets[0].shapes[0].hyperlink.target), result(1, rg(4, 0)));
});

test('링크 재작성은 외부 URL·매크로 주소·동적 문자열을 실행하거나 바꾸지 않는다', () => {
  let calls = 0;
  for (const target of ['https://example.com/#A1', 'mailto:a@example.com', 'javascript:evil()', '#INDIRECT("A1")', '#[Book]Sheet!A1']) {
    assert.equal(rewriteWorkbookLink(target, () => { calls++; return '=B2'; }, () => false), target);
  }
  assert.equal(calls, 0);
});


test('공백·괄호가 포함된 인용되지 않은 시트 링크도 이름 변경과 행 삽입을 따른다', () => {
  const wb = new Workbook({ sheets: [{ name: '광고 요약(통합)', cells: {} }, { name: '메뉴', cells: {}, shapes: [{ id: 'go', hyperlink: { target: '#광고 요약(통합)!A4' } }] }] });
  wb.transact(() => wb.insertRows(0, 0));
  assert.deepEqual(resolveWorkbookLink(wb, 1, wb.sheets[1].shapes[0].hyperlink.target), result(0, rg(4, 0)));
  wb.transact(() => wb.renameSheet(0, '새 이름'));
  assert.equal(wb.sheets[1].shapes[0].hyperlink.target, "#'새 이름'!A5");
});

test('퍼센트 문자열의 실제 시트 이름은 재작성 때도 URL 디코딩보다 우선한다', () => {
  const wb = new Workbook({ sheets: [{ name: '홈', cells: {} }, { name: '%ED%99%88', cells: {}, shapes: [{ id: 'go', hyperlink: { target: "#'%ED%99%88'!A4" } }] }] });
  wb.transact(() => wb.renameSheet(0, '이름 변경'));
  assert.equal(wb.sheets[1].shapes[0].hyperlink.target, "#'%ED%99%88'!A4");
  wb.transact(() => wb.renameSheet(1, '실제 대상'));
  assert.equal(wb.sheets[1].shapes[0].hyperlink.target, "#'실제 대상'!A4");
});

test('개체의 # 없는 내부 관계는 갱신하고 명시적 외부 관계는 원문 그대로 보존한다', () => {
  const wb = new Workbook({ sheets: [{ name: '대상 시트', cells: {} }, { name: '메뉴', cells: {}, shapes: [
    { id: 'bare', hyperlink: { target: "'대상 시트'!B4" } },
    { id: 'internal', hyperlink: { target: "'대상 시트'!B4", targetMode: 'Internal' } },
    { id: 'external', hyperlink: { target: "'대상 시트'!B4", targetMode: 'External' } },
    { id: 'url', hyperlink: { target: 'https://example.com/#B4', targetMode: 'External' } },
  ] }] });
  wb.transact(() => wb.renameSheet(0, '대상 변경'));
  wb.transact(() => wb.insertRows(0, 0));
  assert.equal(wb.sheets[1].shapes[0].hyperlink.target, "'대상 변경'!B5");
  assert.equal(wb.sheets[1].shapes[1].hyperlink.target, "'대상 변경'!B5");
  assert.equal(wb.sheets[1].shapes[2].hyperlink.target, "'대상 시트'!B4");
  assert.equal(wb.sheets[1].shapes[3].hyperlink.target, 'https://example.com/#B4');
});

test('명시 시트 범위 이름의 링크는 시트 이름 변경을 따르며 주소 삽입은 이름 정의만 갱신한다', () => {
  const wb = new Workbook({ names: [{ name: '맨위로', sheet: '대상', ref: '=$A$4' }], sheets: [{ name: '대상', cells: {} }, { name: '메뉴', cells: {}, shapes: [{ id: 'go', hyperlink: { target: '#대상!맨위로' } }] }] });
  wb.transact(() => wb.renameSheet(0, "O'Brien 대상"));
  assert.equal(wb.sheets[1].shapes[0].hyperlink.target, "#'O''Brien 대상'!맨위로");
  wb.transact(() => wb.insertRows(0, 0));
  assert.deepEqual(resolveWorkbookLink(wb, 1, wb.sheets[1].shapes[0].hyperlink.target), result(0, rg(4, 0)));
  assert.equal(wb.sheets[1].shapes[0].hyperlink.target, "#'O''Brien 대상'!맨위로");
});

test('External 관계여도 #fragment는 현재 문서 참조이며 URL#fragment는 외부 주소이다', () => {
  const wb = new Workbook({ sheets: [{ name: '대상', cells: {} }, { name: '메뉴', cells: {}, shapes: [
    { id: 'here', hyperlink: { target: '#대상!A3', targetMode: 'External' } },
    { id: 'there', hyperlink: { target: 'https://example.com/book.xlsx#대상!A3', targetMode: 'External' } },
  ] }] });
  wb.transact(() => { wb.renameSheet(0, '변경'); wb.insertRows(0, 0); });
  assert.equal(wb.sheets[1].shapes[0].hyperlink.target, '#변경!A4');
  assert.equal(wb.sheets[1].shapes[1].hyperlink.target, 'https://example.com/book.xlsx#대상!A3');
});

test('실제로 존재하는 인코딩 모양 시트가 숨김이면 다른 디코딩 시트로 우회하지 않는다', () => {
  const wb = new Workbook({ sheets: [{ name: '홈', cells: {} }, { name: '%ED%99%88', state: 'hidden', cells: {} }] });
  assert.match(resolveWorkbookLink(wb, 0, "#'%ED%99%88'!A1").error, /숨겨진/);
});
