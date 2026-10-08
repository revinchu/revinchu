// 수식 원문을 검토한 뒤 적용하는 함수 이름 수정. 가져오기나 함수 등록은 변경하지 않습니다.
import { FUNCS, parse, tokenize } from './formula.js';

const HANGUL_LETTER = /^[\u1100-\u11ff\u3131-\u318e\ua960-\ua97f\uac00-\ud7a3\ud7b0-\ud7ff]$/;
const HANGUL_CALL = /[A-Za-z][A-Za-z0-9_.]*[\u1100-\u11ff\u3131-\u318e\ua960-\ua97f\uac00-\ud7a3\ud7b0-\ud7ff]\s*\(/;
const issuedPlans = new WeakMap();

/** 알려진 함수 이름 바로 뒤에 한글 한 글자만 붙었을 때의 검토 후보. */
export function suggestFunctionName(name) {
  if (typeof name !== 'string' || !name || FUNCS[name.toUpperCase()]) return null;
  const letters = [...name];
  if (!HANGUL_LETTER.test(letters[letters.length - 1])) return null;
  const known = letters.slice(0, -1).join('').toUpperCase();
  return Object.hasOwn(FUNCS, known) ? known : null;
}

function definedIn(book, si, name) {
  const host = book.sheets[si]?.name.toLowerCase(), low = name.toLowerCase();
  return (book.names ?? []).some(n => n.name.toLowerCase() === low && (!n.sheet || n.sheet.toLowerCase() === host));
}

function repairStarts(ast, candidates, book, si) {
  const eligible = new Set(candidates.map(t => t.s)), starts = new Set();
  const walk = (node, local) => {
    if (!node || typeof node !== 'object') return;
    if (Array.isArray(node)) { for (const child of node) walk(child, local); return; }
    if (node.type === 'func') {
      if (node.name === 'LET') {
        const scope = new Set(local);
        for (let i = 0; i + 1 < node.args.length; i += 2) {
          walk(node.args[i + 1], scope);
          if (node.args[i].type === 'name') scope.add(node.args[i].v.toLowerCase());
        }
        walk(node.args[node.args.length - 1], scope);
        return;
      }
      if (node.name === 'LAMBDA') {
        const scope = new Set(local);
        for (const arg of node.args.slice(0, -1)) if (arg.type === 'name') scope.add(arg.v.toLowerCase());
        walk(node.args[node.args.length - 1], scope);
        return;
      }
      if (eligible.has(node.s) && !local.has(node.name.toLowerCase()) && !definedIn(book, si, node.name)) starts.add(node.s);
    }
    for (const [key, value] of Object.entries(node)) if (key !== 'ref' && value && typeof value === 'object') walk(value, local);
  };
  walk(ast, new Set());
  return starts;
}

function repairedFormula(raw, book, si) {
  if (typeof raw !== 'string' || !HANGUL_CALL.test(raw)) return null;
  const offset = raw.startsWith('=') ? 1 : 0, body = raw.slice(offset);
  let tokens;
  try { tokens = tokenize(body); } catch { return null; }
  const candidates = tokens.filter(t => t.t === 'func' && suggestFunctionName(t.v));
  if (!candidates.length) return null;
  let ast;
  try { ast = parse(body); } catch { return null; }
  // 같은 철자가 바깥에서는 오타이고 LET/LAMBDA 안에서는 지역 함수일 수 있습니다.
  // 이름 집합이 아닌 각 호출의 위치와 범위로 수정 대상을 정합니다.
  const starts = repairStarts(ast, candidates, book, si);
  const changes = candidates.filter(t => starts.has(t.s));
  if (!changes.length) return null;
  const functions = new Map();
  let after = '', at = 0;
  for (const token of changes) {
    const to = suggestFunctionName(token.v), original = body.slice(token.s, token.e);
    // 파일 형식으로 보관된 템플릿도 기존 접두사는 그대로 둡니다.
    const prefix = /^(?:(?:_xlfn|_xlws)\.)+/i.exec(original)?.[0] ?? '';
    after += body.slice(at, token.s) + prefix + to;
    at = token.e;
    functions.set(token.v, Object.freeze({ from: token.v, to }));
  }
  after += body.slice(at);
  return { before: raw, after: raw.slice(0, offset) + after, functions: Object.freeze([...functions.values()]) };
}

function blockedReason(sheet) {
  if (sheet?.protect?.on) return 'protected-sheet';
  if (sheet?.external) return 'external-sheet';
  return null;
}

/** 셀과 표 계산 열 템플릿을 읽기만 하는, 특정 통합 문서·버전에 묶인 수정 계획. */
export function functionNameRepairPlan(book) {
  const cells = [], tables = [], blocked = [], summaries = new Map(), hosts = new Map();
  const add = (list, entry, sheet, kind) => {
    const item = Object.freeze(entry);
    list.push(item);
    hosts.set(item, sheet);
    for (const fn of entry.functions) {
      let summary = summaries.get(fn.from);
      if (!summary) { summary = { ...fn, cells: 0, templates: 0 }; summaries.set(fn.from, summary); }
      summary[kind]++;
    }
  };
  book.sheets.forEach((sheet, si) => {
    const memo = new Map();
    const repair = raw => {
      if (memo.has(raw)) return memo.get(raw);
      const result = repairedFormula(raw, book, si);
      if (memo.size < 200000) memo.set(raw, result);
      return result;
    };
    const beforeCount = cells.length + tables.length;
    sheet.cells.forEachFormulaRC((cell, r, c) => {
      const changed = repair(cell.raw);
      if (changed) add(cells, { si, r, c, sheet: sheet.name, ...changed }, sheet, 'cells');
    });
    (sheet.tables ?? []).forEach((table, tableIndex) => {
      for (const [key, raw] of Object.entries(table.calculatedColumnFormulas ?? {})) {
        const column = Number(key);
        if (!Number.isInteger(column) || column < table.c1 || column > table.c2) continue;
        const changed = repair(raw);
        if (changed) add(tables, { si, tableIndex, tableId: table.id, tableName: table.name, column, sheet: sheet.name, ...changed }, sheet, 'templates');
      }
    });
    const reason = blockedReason(sheet);
    if (reason && beforeCount !== cells.length + tables.length) blocked.push(Object.freeze({ si, sheet: sheet.name, reason }));
  });
  const plan = Object.freeze({ version: book.version, cells: Object.freeze(cells), tables: Object.freeze(tables),
    functions: Object.freeze([...summaries.values()].map(Object.freeze)), blocked: Object.freeze(blocked),
    cellCount: cells.length, templateCount: tables.length, total: cells.length + tables.length });
  issuedPlans.set(plan, { book, hosts });
  return plan;
}

function validateEntry(book, entry, hosts, raw) {
  const sheet = book.sheets[entry.si];
  if (sheet !== hosts.get(entry) || raw !== entry.before) throw new Error('수정 대상 수식이 바뀌었습니다. 함수 이름 수정 목록을 다시 확인하세요.');
  if (blockedReason(sheet)) throw new Error('보호되거나 외부에 연결된 시트의 수식은 수정할 수 없습니다.');
  const current = repairedFormula(raw, book, entry.si);
  if (!current || current.after !== entry.after) throw new Error('함수 이름 또는 정의된 이름이 바뀌었습니다. 수정 목록을 다시 확인하세요.');
}

/** 검토한 계획의 전 대상을 사전 검사한 뒤 하나의 실행 취소 단계로 적용. */
export function applyFunctionNameRepairs(book, plan, meta) {
  const issued = issuedPlans.get(plan);
  if (!issued || issued.book !== book || plan.version !== book.version) throw new Error('통합 문서가 바뀌었습니다. 함수 이름 수정 목록을 다시 확인하세요.');
  if (book.props?.markedFinal) throw new Error('최종본으로 표시된 통합 문서는 수정할 수 없습니다. 편집을 허용한 뒤 다시 확인하세요.');
  const nextTables = new Map();
  for (const entry of plan.cells) validateEntry(book, entry, issued.hosts, book.sheets[entry.si]?.cells.getRC(entry.r, entry.c)?.raw);
  for (const entry of plan.tables) {
    const sheet = book.sheets[entry.si], table = sheet?.tables?.[entry.tableIndex];
    if (!table || table.id !== entry.tableId || table.name !== entry.tableName) throw new Error('수정 대상 표가 바뀌었습니다. 함수 이름 수정 목록을 다시 확인하세요.');
    validateEntry(book, entry, issued.hosts, table.calculatedColumnFormulas?.[entry.column]);
    let tables = nextTables.get(entry.si);
    if (!tables) { tables = [...sheet.tables]; nextTables.set(entry.si, tables); }
    const previous = tables[entry.tableIndex];
    tables[entry.tableIndex] = { ...previous, calculatedColumnFormulas: { ...previous.calculatedColumnFormulas, [entry.column]: entry.after } };
  }
  const result = { cellCount: plan.cellCount, templateCount: plan.templateCount, total: plan.total };
  if (!plan.total) return result;
  book.transact(() => {
    for (const [si, tables] of nextTables) {
      // 템플릿만 바꾸는 것은 표 범위·열 이름·현재 계산값을 바꾸지 않습니다.
      // Workbook.setTableStyle과 같은 중립 속성 기록으로 무관한 캐시를 보존합니다.
      const existed = book.tx?.entries.find(e => e.t === 'prop' && e.si === si && e.prop === 'tables');
      book.propSnap(si, 'tables');
      const entry = book.tx?.entries.find(e => e.t === 'prop' && e.si === si && e.prop === 'tables');
      if (entry && !existed) entry.calcNeutral = true;
      book.sheets[si].tables = tables;
      book.version++;
    }
    for (const entry of plan.cells) {
      const cell = book.sheets[entry.si].cells.getRC(entry.r, entry.c);
      const { cached, cachedArray, staleCached, staleCachedArray, ...data } = cell;
      book.setCellData(entry.si, entry.r, entry.c, { ...data, raw: entry.after });
    }
  }, meta);
  return result;
}
