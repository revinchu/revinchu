import { el } from './ui.js';
import { RANGE_QUERY_TYPES, RANGE_QUERY_FILTERS, rangeQueryHeaders, transformRangeQuery } from './range-query.js';
import { formatNumber } from './format.js';

export function createRangeQueryEditor(source, { name = '정리한 데이터' } = {}) {
  let disposed = false, timer = null, lastKey = '', lastResult = null;
  const columns = [], filters = [], sorts = [];
  const nameInput = el('input', { type: 'text', value: String(name).slice(0, 31), maxlength: 31, 'aria-label': '새 시트 이름' });
  const headerInput = el('input', { type: 'checkbox', checked: source.header !== false, 'aria-label': '첫 행을 머리글로 사용' });
  const blankInput = el('input', { type: 'checkbox', 'aria-label': '빈 행 제거' });
  const duplicateInput = el('input', { type: 'checkbox', 'aria-label': '중복 행 제거' });
  const columnList = el('div', { class: 'rq-column-list' });
  const filterList = el('div', { class: 'rq-rule-list' });
  const sortList = el('div', { class: 'rq-rule-list' });
  const status = el('div', { class: 'rq-status', role: 'status', 'aria-live': 'polite' });
  const errorBox = el('div', { class: 'rq-error', role: 'alert', hidden: true });
  const preview = el('div', { class: 'rq-preview', tabindex: 0, 'aria-label': '변환 결과 미리보기' });
  const steps = el('div', { class: 'rq-steps' });
  const heads = () => rangeQueryHeaders(source, headerInput.checked);
  const select = (items, label) => el('select', { 'aria-label': label }, items.map(([value, text]) => el('option', { value }, text)));
  const field = (label, control) => el('label', { class: 'rq-field' }, el('span', {}, label), control);
  const fields = () => heads().map((h, i) => [String(i), h]);
  const spec = () => ({ header: headerInput.checked, columns: columns.flatMap((c, i) => c.check.checked ? [i] : []), types: columns.map(c => c.type.value), filters: filters.map(f => ({ column: +f.column.value, op: f.op.value, value: f.value.value })), sorts: sorts.map(s => ({ column: +s.column.value, direction: s.direction.value })), removeBlankRows: blankInput.checked, removeDuplicates: duplicateInput.checked });
  const result = () => { const s = spec(), key = JSON.stringify(s); if (key !== lastKey || !lastResult) { lastResult = transformRangeQuery(source, s); lastKey = key; } return lastResult; };
  const display = (v, fmt) => {
    if (v === null || v === undefined) return '';
    if (typeof v === 'object') return String(v.code ?? '#VALUE!');
    if (typeof v === 'boolean') return v ? 'TRUE' : 'FALSE';
    return fmt?.numFmt === 'date' && typeof v === 'number' ? formatNumber(v, 'date', undefined, source.date1904) : String(v);
  };
  const refresh = () => {
    clearTimeout(timer); timer = null; if (disposed) return;
    try {
      const out = result(), { stats } = out;
      const table = el('table', { class: 'rq-table' }, el('thead', {}, el('tr', {}, out.headers.slice(0, 30).map(h => el('th', { scope: 'col', title: h }, h)))));
      const rows = el('tbody');
      for (let r = 1; r < Math.min(out.matrix.length, 51); r++) rows.append(el('tr', {}, out.matrix[r].slice(0, 30).map((v, c) => el('td', {}, display(v, out.columnFormats[c])))));
      table.append(rows); preview.replaceChildren(table);
      if (!stats.outputRows) preview.append(el('p', { class: 'muted' }, '조건에 맞는 데이터가 없습니다. 머리글만 불러옵니다.'));
      status.textContent = `원본 ${stats.inputRows.toLocaleString('ko-KR')}행 → 결과 ${stats.outputRows.toLocaleString('ko-KR')}행 · ${out.headers.length}열 · 미리보기 최대 50행·30열`;
      errorBox.hidden = !stats.errorCount;
      errorBox.textContent = stats.errorCount ? `유형 변환 오류 ${stats.errorCount.toLocaleString('ko-KR')}개: ${out.errors.slice(0, 5).map(e => `${e.row}행 ${e.header}`).join(', ')}. 원본 유지로 바꾸거나 올바른 유형을 선택하세요. 오류가 있으면 불러올 수 없습니다.` : '';
      steps.textContent = `적용 순서: 머리글 → 유형 변환 → 필터(${filters.length}) → 열 선택 → 빈 행 제거(${stats.blankRows}) → 중복 제거(${stats.duplicateRows}) → 정렬(${sorts.length})`;
    } catch (e) { preview.replaceChildren(); status.textContent = '설정을 확인하세요.'; errorBox.hidden = false; errorBox.textContent = e.message; }
  };
  const schedule = () => { if (disposed) return; clearTimeout(timer); timer = setTimeout(refresh, 120); };
  const updateNames = () => {
    const names = heads();
    columns.forEach((c, i) => { c.label.textContent = names[i]; c.check.setAttribute('aria-label', `${names[i]} 열 가져오기`); c.type.setAttribute('aria-label', `${names[i]} 유형`); });
    for (const row of filters.concat(sorts)) { const value = row.column.value; row.column.replaceChildren(); for (const [v, t] of fields()) row.column.append(el('option', { value: v }, t)); row.column.value = value; }
    refresh();
  };
  for (const [i, title] of heads().entries()) {
    const check = el('input', { type: 'checkbox', checked: true, 'aria-label': `${title} 열 가져오기` });
    const label = el('span', { class: 'rq-column-name', title }, title);
    const type = select(RANGE_QUERY_TYPES, `${title} 유형`);
    columns.push({ check, label, type });
    columnList.append(el('div', { class: 'rq-column' }, el('label', {}, check, label), field('유형', type)));
    check.addEventListener('change', schedule); type.addEventListener('change', schedule);
  }
  const addFilter = () => {
    const column = select(fields(), '필터 열'), op = select(RANGE_QUERY_FILTERS, '필터 조건');
    const value = el('input', { type: 'text', 'aria-label': '필터 값', placeholder: '비교할 값' });
    const item = { column, op, value };
    const remove = el('button', { type: 'button', class: 'btn', 'aria-label': '필터 삭제', onclick: () => { filters.splice(filters.indexOf(item), 1); row.remove(); refresh(); } }, '삭제');
    const row = el('div', { class: 'rq-rule' }, field('열', column), field('조건', op), field('값', value), remove);
    filters.push(item); filterList.append(row);
    column.addEventListener('change', schedule); value.addEventListener('input', schedule);
    op.addEventListener('change', () => { value.disabled = op.value === 'blank' || op.value === 'notBlank'; schedule(); });
    column.focus(); refresh();
  };
  const addSort = () => {
    const column = select(fields(), '정렬 열'), direction = select([['asc', '오름차순'], ['desc', '내림차순']], '정렬 방향');
    const item = { column, direction };
    const row = el('div', { class: 'rq-rule rq-sort-rule' }, field('열', column), field('방향', direction), el('button', { type: 'button', class: 'btn', 'aria-label': '정렬 삭제', onclick: () => { sorts.splice(sorts.indexOf(item), 1); row.remove(); refresh(); } }, '삭제'));
    sorts.push(item); sortList.append(row); column.addEventListener('change', schedule); direction.addEventListener('change', schedule); column.focus(); refresh();
  };
  const body = el('div', { class: 'range-query-editor' },
    el('p', { class: 'rq-source' }, `${source.label} · ${source.matrix.length.toLocaleString('ko-KR')}행 × ${source.width}열${source.clipped ? ' (전체 행·열 선택을 사용 범위로 제한)' : ''}`),
    el('p', { class: 'muted' }, '현재 값의 사본을 정리해 새 시트로 불러옵니다. 숨긴 행도 포함하며 원본은 바꾸지 않습니다. 자동 새로 고침·Power Query M은 지원하지 않습니다.'),
    el('div', { class: 'rq-options' }, el('label', {}, '새 시트 이름 ', nameInput), el('label', {}, headerInput, ' 첫 행을 머리글로 사용')),
    el('div', { class: 'rq-workspace' },
      el('section', { class: 'rq-columns', 'aria-label': '가져올 열과 유형' }, el('h3', {}, '열 선택 및 유형'),
        el('div', { class: 'rq-actions' }, el('button', { type: 'button', class: 'btn', onclick: () => { columns.forEach(c => { c.check.checked = true; }); refresh(); } }, '모두 선택'), el('button', { type: 'button', class: 'btn', onclick: () => { columns.forEach(c => { c.check.checked = false; }); refresh(); } }, '모두 해제')),
        columnList, el('p', { class: 'muted' }, '날짜: YYYY-MM-DD 또는 기존 날짜 숫자. 숫자: 쉼표·소수·% 지원. 정수는 소수 부분이 있으면 오류입니다.')),
      el('section', { class: 'rq-rules', 'aria-label': '변환 설정과 미리보기' },
        el('details', { open: true }, el('summary', {}, '필터 · 모든 조건에 일치'), filterList, el('button', { type: 'button', class: 'btn', onclick: addFilter }, '필터 추가')),
        el('details', {}, el('summary', {}, '정렬 · 위 조건부터 적용'), sortList, el('button', { type: 'button', class: 'btn', onclick: addSort }, '정렬 추가')),
        el('div', { class: 'rq-options' }, el('label', {}, blankInput, ' 빈 행 제거'), el('label', {}, duplicateInput, ' 중복 행 제거')),
        el('p', { class: 'muted' }, '빈 행·중복은 선택한 열 기준입니다. 공백 문자와 빈 셀, 숫자와 텍스트는 구분합니다.'), steps, status, errorBox, preview)));
  headerInput.addEventListener('change', updateNames); blankInput.addEventListener('change', schedule); duplicateInput.addEventListener('change', schedule);
  // 입력의 Enter는 미리보기만 갱신합니다. 최종 로드는 대화상자의 명시적 버튼으로 실행합니다.
  body.addEventListener('keydown', e => { if (e.key === 'Enter' && e.target.tagName === 'INPUT') { e.preventDefault(); e.stopPropagation(); refresh(); } });
  refresh();
  return { body, read() {
    if (disposed) throw new Error('데이터 가져오기 창이 닫혔습니다. 다시 열어 주세요.');
    clearTimeout(timer); timer = null;
    const name = nameInput.value.trim();
    if (!name || name.length > 31 || /[\[\]:*?/\\\x00-\x1f]/.test(name) || name.startsWith("'") || name.endsWith("'")) throw new Error('새 시트 이름은 1~31자이며 [ ] : * ? / \\ 또는 양끝 작은따옴표를 사용할 수 없습니다.');
    const out = result();
    if (out.stats.errorCount) { refresh(); throw new Error(`유형 변환 오류 ${out.stats.errorCount}개를 해결한 뒤 불러오세요.`); }
    return { name, ...out, spec: spec() };
  }, dispose() { disposed = true; clearTimeout(timer); timer = null; lastResult = null; } };
}
