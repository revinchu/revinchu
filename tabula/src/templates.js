// 새로 만들기 → 템플릿 (엑셀 · 구글 스프레드시트 기본 템플릿 + 퍼포먼스 마케팅 템플릿) — DOM 없음
// 각 템플릿: { id, name, cat, desc, color, build() → 통합 문서 데이터 { sheets } }

// ───────────── 도우미 ─────────────
const INK = '#1f2937';
const MUTED = '#64748b';
const LINE = '#e2e8f0';
const NUM = { numFmt: 'comma' };
const WON = { numFmt: 'custom', code: '#,##0"원"' };
const PCT = { numFmt: 'percent', decimals: 1 };
const PCT2 = { numFmt: 'percent', decimals: 2 };
const DATE = { numFmt: 'date' };
const box = (color) => ({ bt: true, bb: true, bl: true, br: true, btc: color, bbc: color, blc: color, brc: color });
const head = (fill) => ({ bold: true, color: '#ffffff', fill, align: 'center', valign: 'middle', ...box(fill) });
const cellLine = { ...box(LINE) };
const inputSt = { fill: '#fff8e1', ...box('#f3d58a') };

class S {
  constructor(name, { grid = false, tab = null } = {}) {
    this.d = { name, cells: {}, colWidths: {}, rowHeights: {}, merges: [], cond: [], charts: [], tables: [], validations: [], ...(grid ? {} : { noGrid: true }), ...(tab ? { tabColor: tab } : {}) };
  }
  v(r, c, v, style) {
    if (v === null || v === undefined || v === '') { if (style) this.st(r, c, style); return this; }
    const k = `${r},${c}`;
    const prev = this.d.cells[k]?.style;
    this.d.cells[k] = { raw: String(v), ...(style || prev ? { style: { ...(prev ?? {}), ...(style ?? {}) } } : {}) };
    return this;
  }
  st(r, c, style) {
    const k = `${r},${c}`;
    const cur = this.d.cells[k] ?? { raw: '' };
    this.d.cells[k] = { ...cur, style: { ...(cur.style ?? {}), ...style } };
    return this;
  }
  row(r, c, arr, style) { arr.forEach((v, i) => this.v(r, c + i, v, typeof style === 'function' ? style(i, v) : style)); return this; }
  rows(r, c, matrix, style) { matrix.forEach((row, i) => this.row(r + i, c, row, typeof style === 'function' ? (j, v) => style(i, j, v) : style)); return this; }
  area(r1, c1, r2, c2, style) { for (let r = r1; r <= r2; r++) for (let c = c1; c <= c2; c++) this.st(r, c, style); return this; }
  w(obj) { Object.assign(this.d.colWidths, obj); return this; }
  h(obj) { Object.assign(this.d.rowHeights, obj); return this; }
  m(r1, c1, r2, c2) { this.d.merges.push({ r1, c1, r2, c2 }); return this; }
  title(text, sub, color = INK, span = 8) {
    this.v(0, 1, text, { bold: true, size: 18, color }).h({ 0: 36 });
    if (sub) this.v(1, 1, sub, { color: MUTED, size: 10 });
    if (span) this.m(0, 1, 0, span);
    return this;
  }
  cf(rule) { this.d.cond.push(rule); return this; }
  chart(c) { this.d.charts.push({ id: `ch${this.d.charts.length + 1}${Math.random().toString(36).slice(2, 6)}`, z: 5 + this.d.charts.length, ...c }); return this; }
  table(t) { this.d.tables.push({ id: `tb${Math.random().toString(36).slice(2, 7)}`, header: true, totals: false, banded: true, filter: null, ...t }); return this; }
  dv(rule) { this.d.validations.push({ allowBlank: true, showDropdown: true, ...rule }); return this; }
  freeze(rows, cols = 0) { this.d.freeze = { rows, cols }; return this; }
  set(k, v) { this.d[k] = v; return this; }
  done() { return this.d; }
}
const col = (c) => { let s = ''; let n = c + 1; while (n) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); } return s; };
const A = (r, c) => `${col(c)}${r + 1}`;
const list = (items) => `"${items.join(',')}"`;
// 결정적 난수 (예시 데이터가 매번 같게)
function rng(seed = 7) { let x = seed; return () => { x ^= x << 13; x ^= x >>> 17; x ^= x << 5; return (x >>> 0) / 4294967296; }; }
const iso = (d) => new Date(d).toISOString().slice(0, 10);
const addDays = (s, n) => iso(new Date(new Date(`${s}T00:00:00Z`).getTime() + n * 86400000));

// ───────────── 기본 템플릿 (엑셀 · 스프레드시트와 같은 종류) ─────────────
function perpetualCalendar() {
  const s = new S('만년 달력').title('만년 달력', '연도와 월을 바꾸면 달력이 바뀝니다.', '#1e3a8a', 7);
  s.v(2, 1, '연도', { bold: true }).v(2, 2, '=YEAR(TODAY())', { ...inputSt, align: 'center' }).v(2, 3, '월', { bold: true }).v(2, 4, '=MONTH(TODAY())', { ...inputSt, align: 'center' });
  s.dv({ r1: 2, c1: 4, r2: 2, c2: 4, type: 'list', f1: list([1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12]) });
  s.v(4, 1, '=DATE($C$3,$E$3,1)', { numFmt: 'custom', code: 'yyyy"년" m"월"', bold: true, size: 16, color: '#1e3a8a' }).m(4, 1, 4, 7);
  const days = ['일', '월', '화', '수', '목', '금', '토'];
  s.row(5, 1, days, (i) => ({ ...head(i === 0 ? '#dc2626' : i === 6 ? '#2563eb' : '#1e3a8a') }));
  for (let w = 0; w < 6; w++) {
    for (let d = 0; d < 7; d++) {
      s.v(6 + w, 1 + d, `=DATE($C$3,$E$3,1)-WEEKDAY(DATE($C$3,$E$3,1))+${w * 7 + d + 1}`, { numFmt: 'custom', code: 'd', align: 'right', valign: 'top', size: 14, ...cellLine, color: d === 0 ? '#dc2626' : d === 6 ? '#2563eb' : INK });
    }
    s.h({ [6 + w]: 56 });
  }
  s.cf({ r1: 6, c1: 1, r2: 11, c2: 7, type: 'formula', formula: '=MONTH(B7)<>$E$3', style: { color: '#cbd5e1' } });
  s.cf({ r1: 6, c1: 1, r2: 11, c2: 7, type: 'formula', formula: '=B7=TODAY()', style: { fill: '#fef3c7', bold: true } });
  return { sheets: [s.w({ 0: 16, 1: 90, 2: 90, 3: 90, 4: 90, 5: 90, 6: 90, 7: 90 }).done()] };
}

function personalBudget() {
  const s = new S('월별 개인 예산').title('월별 개인 예산', '노란 칸에 금액을 입력하세요. 차이 · 합계 · 차트가 자동으로 계산됩니다.', '#047857', 6);
  s.row(3, 1, ['수입', '예상', '실제', '차이'], head('#047857'));
  const inc = [['급여', 3200000, 3200000], ['부수입', 300000, 420000], ['기타', 0, 50000]];
  inc.forEach(([n, a, b], i) => s.row(4 + i, 1, [n, a, b, `=D${5 + i}-C${5 + i}`], (j) => (j === 1 || j === 2 ? { ...WON, ...inputSt } : j === 3 ? WON : cellLine)));
  s.row(7, 1, ['수입 합계', '=SUM(C5:C7)', '=SUM(D5:D7)', '=D8-C8'], { bold: true, ...WON, fill: '#d1fae5' });
  s.row(9, 1, ['지출', '예산', '실제', '차이', '사용률'], head('#b91c1c'));
  const exp = [['주거', 900000, 900000], ['식비', 600000, 684000], ['교통', 150000, 132000], ['통신', 80000, 78000], ['보험', 200000, 200000], ['문화 · 여가', 200000, 265000], ['저축 · 투자', 800000, 800000], ['기타', 100000, 57000]];
  exp.forEach(([n, a, b], i) => s.row(10 + i, 1, [n, a, b, `=C${11 + i}-D${11 + i}`, `=IFERROR(D${11 + i}/C${11 + i},0)`], (j) => (j === 1 || j === 2 ? { ...WON, ...inputSt } : j === 3 ? WON : j === 4 ? PCT : cellLine)));
  const e2 = 10 + exp.length - 1;
  s.row(e2 + 1, 1, ['지출 합계', `=SUM(C11:C${e2 + 1})`, `=SUM(D11:D${e2 + 1})`, `=C${e2 + 2}-D${e2 + 2}`, `=IFERROR(D${e2 + 2}/C${e2 + 2},0)`], (j) => ({ bold: true, fill: '#fee2e2', ...(j === 4 ? PCT : WON) }));
  s.row(e2 + 3, 1, ['남은 돈 (실제 수입 − 실제 지출)', '', `=D8-D${e2 + 2}`], { bold: true, size: 12, ...WON });
  s.cf({ r1: 10, c1: 4, r2: e2, c2: 4, type: 'formula', formula: '=E11<0', style: { color: '#dc2626', bold: true } });
  s.cf({ r1: 10, c1: 5, r2: e2, c2: 5, type: 'bar', color: '#fb7185' });
  s.chart({ type: 'bar', title: '항목별 예산 대비 실제', range: { r1: 9, c1: 1, r2: e2, c2: 3 }, x: 640, y: 60, w: 460, h: 330, palette: 'modern' });
  return { sheets: [s.w({ 0: 16, 1: 190, 2: 110, 3: 110, 4: 110, 5: 80 }).freeze(3).done()] };
}

function companyBudget() {
  const s = new S('월간 회사 예산').title('월간 회사 예산', '부서별 예산과 실적, 차이와 달성률', '#1e40af', 8);
  const months = ['1월', '2월', '3월', '4월', '5월', '6월'];
  s.row(3, 1, ['부서', '구분', ...months, '합계'], head('#1e40af'));
  const depts = [['영업', 42000000], ['마케팅', 35000000], ['개발', 58000000], ['인사', 12000000], ['재무', 9000000]];
  const R = rng(3);
  let r = 4;
  for (const [d, base] of depts) {
    s.row(r, 1, [d, '예산', ...months.map(() => Math.round(base / 1e5) * 1e5), `=SUM(D${r + 1}:I${r + 1})`], (j) => (j >= 2 ? { ...NUM, ...(j < 8 ? inputSt : { bold: true }) } : { bold: j === 0 }));
    s.row(r + 1, 1, ['', '실적', ...months.map(() => Math.round((base * (0.85 + R() * 0.3)) / 1e5) * 1e5), `=SUM(D${r + 2}:I${r + 2})`], (j) => (j >= 2 ? { ...NUM, ...(j < 8 ? inputSt : { bold: true }) } : {}));
    s.row(r + 2, 1, ['', '달성률', ...months.map((_, k) => `=IFERROR(${col(3 + k)}${r + 2}/${col(3 + k)}${r + 1},0)`), `=IFERROR(J${r + 2}/J${r + 1},0)`], (j) => (j >= 2 ? { ...PCT, color: MUTED } : { color: MUTED }));
    s.m(r, 1, r + 2, 1).st(r, 1, { valign: 'middle', align: 'center' }).area(r + 2, 1, r + 2, 9, { bb: true, bbc: '#cbd5e1' });
    r += 3;
  }
  s.row(r, 1, ['전체', '예산', ...months.map((_, k) => `=SUMIF($C$5:$C$${r},"예산",${col(3 + k)}$5:${col(3 + k)}$${r})`), `=SUM(D${r + 1}:I${r + 1})`], { bold: true, fill: '#dbeafe', ...NUM });
  s.row(r + 1, 1, ['', '실적', ...months.map((_, k) => `=SUMIF($C$5:$C$${r},"실적",${col(3 + k)}$5:${col(3 + k)}$${r})`), `=SUM(D${r + 2}:I${r + 2})`], { bold: true, fill: '#dbeafe', ...NUM });
  s.cf({ r1: 6, c1: 3, r2: r, c2: 9, type: 'formula', formula: '=AND($C7="달성률",D7>1)', style: { color: '#dc2626', bold: true } });
  s.chart({ type: 'combo', title: '월별 전체 예산 · 실적', range: { r1: r, c1: 2, r2: r + 1, c2: 8 }, byRows: true, x: 60, y: (r + 3) * 20 + 40, w: 640, h: 280, palette: 'modern' });
  return { sheets: [s.w({ 0: 16, 1: 90, 2: 70, 9: 120 }).freeze(4, 3).done()] };
}

function balanceSheet() {
  const s = new S('대차대조표').title('대차대조표 (재무상태표)', '기준일과 금액을 입력하세요. 자산 = 부채 + 자본 검증이 자동으로 됩니다.', '#334155', 4);
  s.v(2, 1, '기준일').v(2, 2, '2026-12-31', { ...DATE, ...inputSt });
  const sec = (r, t, items, color) => {
    s.row(r, 1, [t, '금액'], head(color));
    items.forEach((it, i) => s.row(r + 1 + i, 1, [it[0], it[1]], (j) => (j === 1 ? { ...WON, ...inputSt } : { ...cellLine, indent: 1 })));
    const end = r + items.length;
    s.row(end + 1, 1, [`${t} 합계`, `=SUM(C${r + 2}:C${end + 1})`], { bold: true, ...WON, fill: '#f1f5f9', bt: true });
    return end + 2;
  };
  let r = 4;
  const aStart = r;
  r = sec(r, '유동자산', [['현금 및 현금성자산', 85000000], ['매출채권', 42000000], ['재고자산', 31000000], ['선급비용', 5000000]], '#0f766e');
  const r1 = r;
  r = sec(r + 1, '비유동자산', [['유형자산', 120000000], ['무형자산', 18000000], ['투자자산', 25000000]], '#0f766e');
  const r2 = r;
  s.row(r, 1, ['자산 총계', `=C${r1}+C${r2}`], { bold: true, size: 12, ...WON, fill: '#ccfbf1', bb: true, bbs: 'double' });
  r += 2;
  r = sec(r, '유동부채', [['매입채무', 28000000], ['단기차입금', 30000000], ['미지급비용', 7000000]], '#b45309');
  const r3 = r;
  r = sec(r + 1, '비유동부채', [['장기차입금', 60000000], ['퇴직급여충당부채', 12000000]], '#b45309');
  const r4 = r;
  s.row(r, 1, ['부채 총계', `=C${r3}+C${r4}`], { bold: true, ...WON, fill: '#fef3c7' });
  r += 2;
  r = sec(r, '자본', [['자본금', 100000000], ['이익잉여금', '=C' + (aStart + 1) + '*0']], '#1d4ed8');
  const r5 = r;
  s.v(r5 - 2, 2, `=C${r2 + 1}-C${r3 + 0}-C${r4}-100000000`, { ...WON, ...inputSt });
  s.row(r, 1, ['부채와 자본 총계', `=C${r3}+C${r4}+C${r5}`], { bold: true, size: 12, ...WON, fill: '#dbeafe', bb: true, bbs: 'double' });
  s.row(r + 2, 1, ['검증 (자산 − 부채 · 자본)', `=C${r2 + 1}-C${r + 1}`], { bold: true, ...WON });
  s.cf({ r1: r + 2, c1: 2, r2: r + 2, c2: 2, type: 'formula', formula: `=C${r + 3}<>0`, style: { fill: '#fee2e2', color: '#b91c1c' } });
  return { sheets: [s.w({ 0: 16, 1: 240, 2: 160 }).done()] };
}

function householdBudget() {
  const s = new S('가구 월별 예산').title('가구 월별 예산', '12개월 가계 지출을 기록하면 월 합계 · 항목 비중 · 추세가 계산됩니다.', '#7c3aed', 15);
  const months = Array.from({ length: 12 }, (_, i) => `${i + 1}월`);
  s.row(3, 1, ['항목', ...months, '합계', '비중'], head('#7c3aed'));
  const cats = ['주거 · 관리비', '식료품', '외식', '교통 · 차량', '교육', '의료', '보험', '통신', '생활용품', '경조사', '여가'];
  const R = rng(11);
  cats.forEach((c, i) => {
    const base = [900, 700, 250, 300, 400, 80, 250, 120, 100, 80, 150][i] * 1000;
    s.row(4 + i, 1, [c, ...months.map(() => Math.round((base * (0.8 + R() * 0.4)) / 1000) * 1000), `=SUM(C${5 + i}:N${5 + i})`, `=O${5 + i}/$O$${5 + cats.length}`], (j) => (j === 0 ? cellLine : j === 14 ? PCT : { ...NUM, ...(j <= 12 ? {} : { bold: true }) }));
  });
  const tr = 4 + cats.length;
  s.row(tr, 1, ['월 합계', ...months.map((_, k) => `=SUM(${col(2 + k)}5:${col(2 + k)}${tr})`), `=SUM(O5:O${tr})`, '=SUM(P5:P' + tr + ')'], (j) => ({ bold: true, fill: '#ede9fe', ...(j === 14 ? PCT : NUM) }));
  s.cf({ r1: 4, c1: 2, r2: tr - 1, c2: 13, type: 'scale', colors: ['#ffffff', '#c4b5fd'] });
  s.chart({ type: 'doughnut', title: '연간 지출 비중', range: { r1: 3, c1: 1, r2: tr - 1, c2: 1 }, series: [{ name: { text: '합계' }, cat: { r1: 4, c1: 1, r2: tr - 1, c2: 1 }, val: { r1: 4, c1: 14, r2: tr - 1, c2: 14 } }], x: 80, y: (tr + 2) * 20 + 40, w: 420, h: 300, legend: 'r', palette: 'pastel' });
  s.chart({ type: 'line', title: '월별 지출 추세', series: [{ name: { text: '월 합계' }, cat: { r1: 3, c1: 2, r2: 3, c2: 13 }, val: { r1: tr, c1: 2, r2: tr, c2: 13 } }], x: 520, y: (tr + 2) * 20 + 40, w: 520, h: 300, palette: 'modern', legend: 'none' });
  return { sheets: [s.w({ 0: 16, 1: 120, 14: 110, 15: 60 }).freeze(4, 2).done()] };
}

function gantt(name = '간트 프로젝트 플래너', color = '#0369a1', tasks = null, sub = '시작일 · 기간 · 진행률을 입력하면 막대가 그려집니다. 프로젝트 시작일(D3)을 바꾸면 달력이 이동합니다.') {
  const s = new S(name).title(name, sub, color, 10);
  s.v(2, 1, '프로젝트 시작', { bold: true }).v(2, 3, '2026-10-05', { ...DATE, ...inputSt }).v(2, 5, '오늘', { bold: true }).v(2, 6, '=TODAY()', DATE);
  const T = tasks ?? [
    ['기획', '요구사항 정리', '김PM', 0, 4, 1], ['기획', '일정 · 범위 확정', '김PM', 3, 3, 1], ['디자인', '와이어프레임', '이디자인', 5, 5, 0.8],
    ['디자인', 'UI 디자인', '이디자인', 9, 7, 0.4], ['개발', '프론트엔드', '박개발', 12, 12, 0.2], ['개발', '백엔드 API', '최개발', 12, 10, 0.3],
    ['검수', 'QA 테스트', '정QA', 23, 6, 0], ['출시', '배포 · 모니터링', '박개발', 29, 3, 0],
  ];
  s.row(4, 1, ['단계', '작업', '담당', '시작', '기간(일)', '종료', '진행률'], head(color));
  const days = 42;
  for (let d = 0; d < days; d++) {
    s.v(3, 8 + d, d % 7 === 0 ? `=$D$3+${d}` : '', { numFmt: 'custom', code: 'm/d', size: 8, color: MUTED });
    s.v(4, 8 + d, `=$D$3+${d}`, { numFmt: 'custom', code: 'd', size: 8, align: 'center', fill: '#f1f5f9', color: INK });
  }
  T.forEach(([ph, task, who, off, dur, pct], i) => {
    const r = 5 + i;
    s.row(r, 1, [ph, task, who, `=$D$3+${off}`, dur, `=E${r + 1}+F${r + 1}-1`, pct], (j) => (j === 3 || j === 5 ? { ...DATE, ...(j === 3 ? inputSt : {}) } : j === 6 ? { ...PCT, ...inputSt } : j === 4 ? inputSt : cellLine));
  });
  const last = 4 + T.length;
  const c0 = col(8);
  s.cf({ r1: 5, c1: 8, r2: last, c2: 7 + days, type: 'formula', formula: `=AND(${c0}$5>=$E6,${c0}$5<=$E6+($G6-1)*$H6)`, style: { fill: color } });
  s.cf({ r1: 5, c1: 8, r2: last, c2: 7 + days, type: 'formula', formula: `=AND(${c0}$5>=$E6,${c0}$5<=$G6)`, style: { fill: '#bae6fd' } });
  s.cf({ r1: 4, c1: 8, r2: last, c2: 7 + days, type: 'formula', formula: `=${c0}$5=TODAY()`, style: { bl: true, blc: '#dc2626', bls: 'medium' } });
  s.cf({ r1: 5, c1: 7, r2: last, c2: 7, type: 'bar', color: color });
  const w = { 0: 16, 1: 70, 2: 150, 3: 70, 4: 80, 5: 60, 6: 80, 7: 70 };
  for (let d = 0; d < days; d++) w[8 + d] = 22;
  return { sheets: [s.w(w).freeze(5, 3).done()] };
}

function vacationSchedule() {
  const s = new S('휴가 일정').title('직원 휴가 일정', '시작 · 종료일을 입력하면 근무일 기준 휴가 일수와 달력 표시가 자동으로 됩니다.', '#be185d', 8);
  s.v(2, 1, '달력 시작', { bold: true }).v(2, 2, '2026-10-01', { ...DATE, ...inputSt });
  s.row(4, 1, ['직원', '부서', '유형', '시작', '종료', '일수'], head('#be185d'));
  const V = [['김하늘', '영업', '연차', 2, 6], ['이서준', '마케팅', '반차', 9, 9], ['박지우', '개발', '연차', 13, 17], ['최민서', '인사', '병가', 20, 21], ['정도윤', '마케팅', '연차', 26, 30], ['한예린', '개발', '경조사', 7, 8]];
  V.forEach(([n, d, t, a, b], i) => s.row(5 + i, 1, [n, d, t, `=$C$3+${a}`, `=$C$3+${b}`, `=NETWORKDAYS(E${6 + i},F${6 + i})`], (j) => (j === 3 || j === 4 ? { ...DATE, ...inputSt } : cellLine)));
  s.dv({ r1: 5, c1: 3, r2: 40, c2: 3, type: 'list', f1: list(['연차', '반차', '병가', '경조사', '공가']) });
  for (let d = 0; d < 31; d++) s.v(4, 7 + d, `=$C$3+${d}`, { numFmt: 'custom', code: 'd', size: 8, align: 'center', fill: '#fce7f3' });
  s.cf({ r1: 5, c1: 7, r2: 40, c2: 37, type: 'formula', formula: '=AND(H$5>=$E6,H$5<=$F6,WEEKDAY(H$5,2)<6)', style: { fill: '#ec4899' } });
  s.cf({ r1: 4, c1: 7, r2: 40, c2: 37, type: 'formula', formula: '=WEEKDAY(H$5,2)>5', style: { fill: '#f1f5f9' } });
  const w = { 0: 16, 1: 80, 2: 70, 3: 60, 4: 90, 5: 90, 6: 50 };
  for (let d = 0; d < 31; d++) w[7 + d] = 22;
  return { sheets: [s.w(w).freeze(5, 2).done()] };
}

function invoice() {
  const s = new S('송장').title('거래명세서 · 송장 (INVOICE)', null, '#111827', 6);
  s.row(2, 1, ['공급자', '(주)위셀마케팅', '', '송장 번호', 'INV-2026-1001'], (j) => (j % 3 === 0 ? { bold: true, color: MUTED } : inputSt));
  s.row(3, 1, ['주소', '서울시 강남구 테헤란로 123', '', '발행일', '=TODAY()'], (j) => (j % 3 === 0 ? { bold: true, color: MUTED } : j === 4 ? DATE : inputSt));
  s.row(4, 1, ['공급받는 자', '(주)고객사', '', '결제 기한', '=F4+30'], (j) => (j % 3 === 0 ? { bold: true, color: MUTED } : j === 4 ? DATE : inputSt));
  s.row(6, 1, ['번호', '품목 · 내용', '수량', '단가', '금액'], head('#111827'));
  const items = [['검색광고 운영 대행 (10월)', 1, 1500000], ['메타 광고 소재 제작', 6, 120000], ['랜딩 페이지 A/B 테스트', 1, 800000], ['월간 리포트', 1, 300000]];
  for (let i = 0; i < 10; i++) {
    const it = items[i];
    s.row(7 + i, 1, [i + 1, it?.[0] ?? '', it?.[1] ?? '', it?.[2] ?? '', `=IF(D${8 + i}="","",D${8 + i}*E${8 + i})`], (j) => (j === 0 ? { align: 'center', ...cellLine } : j === 1 ? { ...inputSt } : j === 4 ? { ...WON, ...cellLine } : { ...NUM, ...inputSt }));
  }
  s.row(18, 4, ['공급가액', '=SUM(F8:F17)'], (j) => ({ bold: j === 0, ...(j ? WON : { align: 'right' }) }));
  s.row(19, 4, ['부가세 (10%)', '=ROUND(F19*0.1,0)'], (j) => (j ? WON : { align: 'right' }));
  s.row(20, 4, ['합계', '=F19+F20'], (j) => ({ bold: true, size: 13, ...(j ? { ...WON, fill: '#111827', color: '#ffffff' } : { align: 'right' }) }));
  s.v(22, 1, '입금 계좌: 위셀은행 123-456-789012 (예금주: (주)위셀마케팅)', { color: MUTED });
  return { sheets: [s.w({ 0: 16, 1: 60, 2: 300, 3: 70, 4: 110, 5: 140 }).done()] };
}

function homeInventory() {
  const s = new S('집 재고').title('우리 집 물품 목록 (보험 · 이사용)', '방별 물품과 가치를 기록하세요.', '#9a3412', 8);
  s.table({ name: '집물품', r1: 3, c1: 1, r2: 11, c2: 8, style: 'WixelTableModern6' });
  s.row(3, 1, ['물품', '방', '브랜드 · 모델', '구입일', '구입가', '수량', '현재 가치', '보증 만료']);
  const it = [['TV', '거실', 'LG OLED 65', '2024-03-10', 2890000, 1], ['소파', '거실', '4인 패브릭', '2023-06-01', 1450000, 1], ['냉장고', '주방', '삼성 비스포크', '2022-11-20', 2300000, 1], ['노트북', '서재', '맥북 프로 14', '2025-02-14', 2990000, 1], ['침대', '안방', '퀸 사이즈', '2021-09-05', 1200000, 1], ['에어컨', '안방', '벽걸이', '2023-05-18', 890000, 2], ['자전거', '베란다', '하이브리드', '2024-04-02', 650000, 1], ['세탁기', '다용도실', '드럼 21kg', '2022-07-30', 1350000, 1]];
  it.forEach((row, i) => { const r = 4 + i; s.row(r, 1, [...row, `=F${r + 1}*G${r + 1}*MAX(0.2,1-0.15*DATEDIF(E${r + 1},TODAY(),"Y"))`, `=EDATE(E${r + 1},24)`], (j) => (j === 3 || j === 7 ? DATE : j === 4 || j === 6 ? WON : {})); });
  s.row(13, 6, ['합계', '=SUM(H5:H12)'], (j) => ({ bold: true, ...(j ? WON : {}) }));
  s.cf({ r1: 4, c1: 8, r2: 11, c2: 8, type: 'formula', formula: '=I5<TODAY()', style: { color: '#dc2626' } });
  s.chart({ type: 'pie', title: '방별 구입가', series: [{ name: { text: '구입가' }, cat: { r1: 4, c1: 2, r2: 11, c2: 2 }, val: { r1: 4, c1: 5, r2: 11, c2: 5 } }], x: 60, y: 330, w: 420, h: 260, legend: 'r', palette: 'modern' });
  return { sheets: [s.w({ 0: 16, 1: 90, 2: 80, 3: 130, 4: 90, 5: 100, 6: 50, 7: 110, 8: 90 }).done()] };
}

function inventoryList() {
  const s = new S('재고 목록').title('재고 목록', '재고가 재주문 수준 이하가 되면 빨간색과 [재주문] 표시가 나타납니다.', '#065f46', 9);
  s.table({ name: '재고', r1: 3, c1: 1, r2: 13, c2: 9, style: 'WixelTableModern4' });
  s.row(3, 1, ['SKU', '품명', '분류', '단가', '재고 수량', '재고 가치', '재주문 수준', '재주문 수량', '상태']);
  const R = rng(5);
  const cat = ['상의', '하의', '신발', '액세서리', '가방'];
  for (let i = 0; i < 10; i++) {
    const r = 4 + i;
    s.row(r, 1, [`SKU-${1001 + i}`, `상품 ${String.fromCharCode(65 + i)}`, cat[i % 5], Math.round(10 + R() * 90) * 1000, Math.round(R() * 120), `=E${r + 1}*F${r + 1}`, 20, 50, `=IF(F${r + 1}<=H${r + 1},"재주문","정상")`], (j) => (j === 3 || j === 5 ? WON : j === 4 || j === 6 || j === 7 ? NUM : {}));
  }
  s.cf({ r1: 4, c1: 9, r2: 13, c2: 9, type: 'formula', formula: '=J5="재주문"', style: { fill: '#fee2e2', color: '#b91c1c', bold: true } });
  s.cf({ r1: 4, c1: 5, r2: 13, c2: 5, type: 'bar', color: '#34d399' });
  s.row(15, 5, ['총 재고 가치', '=SUM(G5:G14)'], (j) => ({ bold: true, ...(j ? WON : {}) }));
  return { sheets: [s.w({ 0: 16, 1: 90, 2: 100, 3: 80, 4: 90, 5: 80, 6: 110, 7: 90, 8: 90, 9: 80 }).done()] };
}

function loanCalculator() {
  const s = new S('대출 계산기').title('대출 계산기', '대출 금액 · 연이율 · 기간을 입력하세요.', '#0f766e', 4);
  const rows = [['대출 금액', 300000000, WON], ['연이율', 0.045, PCT2], ['대출 기간 (년)', 30, NUM], ['연간 상환 횟수', 12, NUM], ['대출 시작일', '2026-11-01', DATE]];
  rows.forEach(([l, v, f], i) => s.row(3 + i, 1, [l, v], (j) => (j ? { ...f, ...inputSt } : { bold: true })));
  s.row(9, 1, ['월 상환액', '=PMT(C5/C7,C6*C7,-C4)'], (j) => (j ? { ...WON, bold: true, size: 14, color: '#0f766e' } : { bold: true }));
  s.row(10, 1, ['총 상환 횟수', '=C6*C7'], (j) => (j ? NUM : { bold: true }));
  s.row(11, 1, ['총 이자', '=C10*C11-C4'], (j) => (j ? WON : { bold: true }));
  s.row(12, 1, ['총 상환액', '=C10*C11'], (j) => (j ? WON : { bold: true }));
  s.row(13, 1, ['마지막 상환일', '=EDATE(C8,C11-1)'], (j) => (j ? DATE : { bold: true }));
  s.row(15, 1, ['금리별 월 상환액', '월 상환액'], head('#0f766e'));
  [0.03, 0.035, 0.04, 0.045, 0.05, 0.055, 0.06].forEach((rt, i) => s.row(16 + i, 1, [rt, `=PMT(B${17 + i}/$C$7,$C$6*$C$7,-$C$4)`], (j) => (j ? WON : PCT2)));
  s.chart({ type: 'line', title: '금리별 월 상환액', series: [{ name: { text: '월 상환액' }, cat: { r1: 16, c1: 1, r2: 22, c2: 1 }, val: { r1: 16, c1: 2, r2: 22, c2: 2 } }], x: 420, y: 60, w: 460, h: 280, legend: 'none', palette: 'modern' });
  return { sheets: [s.w({ 0: 16, 1: 150, 2: 150 }).done()] };
}

function loanSchedule() {
  const s = new S('대출 상환 일정').title('대출 상환 일정표 (원리금 균등)', '위쪽 노란 칸을 바꾸면 상환 일정 전체가 다시 계산됩니다.', '#1d4ed8', 7);
  const rows = [['대출 금액', 50000000, WON], ['연이율', 0.052, PCT2], ['기간 (개월)', 60, NUM], ['시작일', '2026-11-01', DATE]];
  rows.forEach(([l, v, f], i) => s.row(2 + i, 1, [l, v], (j) => (j ? { ...f, ...inputSt } : { bold: true })));
  s.row(2, 4, ['월 상환액', '=PMT(C4/12,C5,-C3)'], (j) => (j ? { ...WON, bold: true } : { bold: true }));
  s.row(3, 4, ['총 이자', '=F3*C5-C3'], (j) => (j ? WON : { bold: true }));
  s.row(7, 1, ['회차', '상환일', '기초 잔액', '상환액', '원금', '이자', '기말 잔액'], head('#1d4ed8'));
  for (let i = 0; i < 60; i++) {
    const r = 8 + i;
    const R1 = r + 1;
    s.row(r, 1, [`=IF(${i + 1}<=$C$5,${i + 1},"")`, `=IF(B${R1}="","",EDATE($C$6,B${R1}-1))`, i === 0 ? '=C3' : `=IF(B${R1}="","",H${R1 - 1})`, `=IF(B${R1}="","",$F$3)`, `=IF(B${R1}="","",PPMT($C$4/12,B${R1},$C$5,-$C$3))`, `=IF(B${R1}="","",IPMT($C$4/12,B${R1},$C$5,-$C$3))`, `=IF(B${R1}="","",D${R1}-F${R1})`],
      (j) => (j === 0 ? { align: 'center' } : j === 1 ? DATE : WON));
  }
  s.cf({ r1: 8, c1: 7, r2: 67, c2: 7, type: 'bar', color: '#93c5fd' });
  s.chart({ type: 'area', grouping: 'stacked', title: '회차별 원금 · 이자', series: [{ name: { text: '원금' }, cat: { r1: 8, c1: 1, r2: 67, c2: 1 }, val: { r1: 8, c1: 5, r2: 67, c2: 5 } }, { name: { text: '이자' }, cat: { r1: 8, c1: 1, r2: 67, c2: 1 }, val: { r1: 8, c1: 6, r2: 67, c2: 6 } }], x: 760, y: 60, w: 460, h: 280, palette: 'modern' });
  return { sheets: [s.w({ 0: 16, 1: 70, 2: 110, 3: 130, 4: 120, 5: 120, 6: 120, 7: 130 }).freeze(8).done()] };
}

function studentSchedule() {
  const s = new S('학생 일정').title('주간 시간표 · 학습 계획', '과목을 입력하면 과목별 색이 자동으로 칠해집니다.', '#4338ca', 6);
  const days = ['월', '화', '수', '목', '금'];
  s.row(3, 1, ['교시', '시간', ...days], head('#4338ca'));
  const T = [['1', '09:00'], ['2', '10:00'], ['3', '11:00'], ['4', '12:00'], ['점심', '13:00'], ['5', '14:00'], ['6', '15:00'], ['방과 후', '16:30']];
  const subj = [['국어', '수학', '영어', '과학', '사회'], ['수학', '영어', '국어', '체육', '과학'], ['영어', '과학', '수학', '국어', '음악'], ['사회', '국어', '미술', '수학', '영어'], ['', '', '', '', ''], ['과학', '체육', '영어', '사회', '수학'], ['음악', '사회', '과학', '영어', '국어'], ['학원', '독서', '학원', '독서', '자유']];
  T.forEach(([p, t], i) => s.row(4 + i, 1, [p, t, ...subj[i]], (j) => ({ ...cellLine, align: 'center', valign: 'middle', ...(j < 2 ? { bold: true, fill: '#eef2ff' } : {}) })));
  s.h(Object.fromEntries(T.map((_, i) => [4 + i, 34])));
  const colors = { 국어: '#fee2e2', 수학: '#dbeafe', 영어: '#dcfce7', 과학: '#fef9c3', 사회: '#f3e8ff', 체육: '#ffedd5', 음악: '#fce7f3', 미술: '#e0f2fe' };
  for (const [k, c] of Object.entries(colors)) s.cf({ r1: 4, c1: 3, r2: 11, c2: 7, type: 'formula', formula: `=D5="${k}"`, style: { fill: c } });
  s.row(13, 1, ['과제 · 시험', '마감', '과목', '완료'], head('#4338ca'));
  [['수학 문제집 3단원', '2026-10-08', '수학', '✓'], ['영어 에세이', '2026-10-10', '영어', ''], ['과학 탐구 보고서', '2026-10-15', '과학', '']].forEach((row, i) => s.row(14 + i, 1, row, (j) => (j === 1 ? DATE : {})));
  s.dv({ r1: 14, c1: 4, r2: 30, c2: 4, type: 'list', f1: list(['✓', '']) });
  s.cf({ r1: 14, c1: 1, r2: 30, c2: 4, type: 'formula', formula: '=$E15="✓"', style: { color: '#94a3b8', strike: true } });
  return { sheets: [s.w({ 0: 16, 1: 70, 2: 70, 3: 110, 4: 110, 5: 110, 6: 110, 7: 110 }).done()] };
}

function weeklyAttendance() {
  const s = new S('주간 출석').title('주간 출석 보고서', '출석(O) · 지각(L) · 결석(X) · 조퇴(E)를 고르면 출석률이 계산됩니다.', '#15803d', 10);
  s.v(2, 1, '주 시작일', { bold: true }).v(2, 2, '2026-10-05', { ...DATE, ...inputSt });
  const days = ['월', '화', '수', '목', '금'];
  s.row(4, 1, ['이름', ...days.map((d, i) => `=TEXT($C$3+${i},"m/d")&" (${d})"`), '출석', '지각', '결석', '출석률'], head('#15803d'));
  const names = ['김하늘', '이서준', '박지우', '최민서', '정도윤', '한예린', '오지호', '윤서아'];
  const R = rng(9);
  names.forEach((n, i) => {
    const r = 5 + i;
    const marks = days.map(() => { const x = R(); return x < 0.8 ? 'O' : x < 0.9 ? 'L' : x < 0.96 ? 'X' : 'E'; });
    s.row(r, 1, [n, ...marks, `=COUNTIF(C${r + 1}:G${r + 1},"O")`, `=COUNTIF(C${r + 1}:G${r + 1},"L")`, `=COUNTIF(C${r + 1}:G${r + 1},"X")`, `=(H${r + 1}+I${r + 1}*0.5)/5`], (j) => (j >= 1 && j <= 5 ? { align: 'center', ...inputSt } : j === 9 ? PCT : { align: 'center' }));
  });
  const last = 4 + names.length;
  s.dv({ r1: 5, c1: 2, r2: last, c2: 6, type: 'list', f1: list(['O', 'L', 'X', 'E']) });
  s.cf({ r1: 5, c1: 2, r2: last, c2: 6, type: 'formula', formula: '=C6="X"', style: { fill: '#fee2e2', color: '#b91c1c', bold: true } });
  s.cf({ r1: 5, c1: 2, r2: last, c2: 6, type: 'formula', formula: '=C6="L"', style: { fill: '#fef9c3' } });
  s.cf({ r1: 5, c1: 10, r2: last, c2: 10, type: 'icons', icons: '3TrafficLights1' });
  s.row(last + 2, 1, ['요일별 출석', ...days.map((_, k) => `=COUNTIF(${col(2 + k)}6:${col(2 + k)}${last + 1},"O")/COUNTA($B$6:$B$${last + 1})`)], (j) => (j ? { ...PCT, bold: true } : { bold: true }));
  return { sheets: [s.w({ 0: 16, 1: 80, 2: 80, 3: 80, 4: 80, 5: 80, 6: 80, 7: 50, 8: 50, 9: 50, 10: 70 }).done()] };
}

// ───────────── 퍼포먼스 마케팅 템플릿 ─────────────
const MK = '#4f46e5';
const CH = ['네이버 SA', '구글 Ads', '메타', '카카오', '유튜브', '틱톡'];

/** 성과 원본(날짜 · 채널 · 캠페인 · 노출 · 클릭 · 비용 · 전환 · 매출) 예시 행 */
function perfRows(n = 180, seed = 21, start = '2026-09-01') {
  const R = rng(seed);
  const camps = { '네이버 SA': ['브랜드', '일반 키워드'], '구글 Ads': ['검색', 'PMax'], 메타: ['리타겟팅', '관심사'], 카카오: ['비즈보드'], 유튜브: ['인스트림'], 틱톡: ['스파크'] };
  const out = [];
  for (let i = 0; i < n; i++) {
    const ch = CH[Math.floor(R() * CH.length)];
    const cp = camps[ch][Math.floor(R() * camps[ch].length)];
    const imp = Math.round(2000 + R() * 40000);
    const clk = Math.round(imp * (0.005 + R() * 0.03));
    const cost = Math.round(clk * (200 + R() * 900));
    const conv = Math.round(clk * R() * 0.08);
    const rev = conv * Math.round(30000 + R() * 50000);
    out.push([addDays(start, Math.floor(R() * 30)), ch, `${ch}_${cp}`, imp, clk, cost, conv, rev]);
  }
  return out.sort((a, b) => (a[0] < b[0] ? -1 : 1));
}

function perfDashboard() {
  const raw = new S('원본', { grid: true, tab: '#94a3b8' });
  const rows = perfRows(240, 21);
  raw.row(0, 0, ['날짜', '채널', '캠페인', '노출수', '클릭수', '비용', '전환수', '매출']);
  rows.forEach((r, i) => raw.row(1 + i, 0, r, (j) => (j === 0 ? DATE : j >= 3 ? NUM : undefined)));
  raw.table({ name: '성과', r1: 0, c1: 0, r2: rows.length, c2: 7, style: 'WixelTableMinimal1' }).freeze(1).w({ 0: 90, 1: 80, 2: 150 });
  const s = new S('퍼포먼스 대시보드', { tab: MK }).title('퍼포먼스 마케팅 대시보드', '[원본] 시트 표에 데이터를 붙여넣으면 KPI · 피벗 · 차트가 자동으로 바뀝니다.', MK, 12);
  const kpis = [['광고비', '=SUM(성과[비용])', WON], ['매출', '=SUM(성과[매출])', WON], ['ROAS', '=IFERROR(SUM(성과[매출])/SUM(성과[비용]),0)', PCT], ['전환수', '=SUM(성과[전환수])', NUM], ['CPA', '=IFERROR(SUM(성과[비용])/SUM(성과[전환수]),0)', WON], ['CTR', '=IFERROR(SUM(성과[클릭수])/SUM(성과[노출수]),0)', PCT2]];
  kpis.forEach(([l, f, fmt], i) => {
    const c = 1 + i * 2;
    s.v(3, c, l, { color: MUTED, size: 9, fill: '#f8fafc', bt: true, btc: MK, bts: 'medium' }).v(4, c, f, { ...fmt, bold: true, size: 16, color: INK, fill: '#f8fafc', bb: true, bbc: LINE }).m(3, c, 3, c + 1).m(4, c, 4, c + 1);
  });
  s.h({ 3: 22, 4: 34 });
  const calcFields = [{ name: 'CTR', formula: 'DIVIDE(클릭수,노출수)' }, { name: 'CPC', formula: 'DIVIDE(비용,클릭수)' }, { name: 'CVR', formula: 'DIVIDE(전환수,클릭수)' }, { name: 'CPA', formula: 'DIVIDE(비용,전환수)' }, { name: 'ROAS', formula: 'DIVIDE(매출,비용)' }];
  const v = (field, numFmt = NUM) => ({ field, agg: 'sum', name: field, numFmt });
  const base = { table: '성과', source: '원본', calcFields, layout: 'tabular', autoRefresh: true, autofit: false, errorCaption: '', grandRows: true, grandCols: true };
  const pv1 = { ...base, name: '채널별', rows: ['채널'], cols: [], values: [v('비용', WON), v('전환수'), v('매출', WON), v('CPA', WON), v('ROAS', PCT), v('CTR', PCT2)], top: 22, left: 1, style: 'WixelPivotModern2', needsRender: true };
  const pv2 = { ...base, name: '일별', rows: ['날짜'], cols: [], values: [v('비용', WON), v('ROAS', PCT)], top: 7, left: 11, style: 'WixelPivotLight2', needsRender: true, grandRows: false };
  const pivots = [{ sheet: '퍼포먼스 대시보드', name: '채널별' }, { sheet: '퍼포먼스 대시보드', name: '일별' }];
  s.set('pivot', pv1).set('pivotsExtra', [pv2]);
  s.set('slicers', [{ id: 'dsl1', caption: '채널', source: { kind: 'pivot', field: '채널', pivots }, columns: 6, style: 'WixelSlicerSolid2', multi: true, x: 30, y: 340, w: 860, h: 76, z: 20 }]);
  s.chart({ type: 'combo', title: '일별 광고비 · ROAS', pivot: { sheet: '퍼포먼스 대시보드', name: '일별' }, x: 30, y: 660, w: 560, h: 300, palette: 'modern', seriesFmt: [{ type: 'column' }, { type: 'line', axis: 1, numFmt: '0%' }], legend: 't', rounded: true, border: LINE, titleBold: true, fieldButtons: false });
  s.chart({ type: 'doughnut', title: '채널별 광고비 비중', pivot: { sheet: '퍼포먼스 대시보드', name: '채널별' }, x: 600, y: 660, w: 290, h: 300, palette: 'modern', legend: 'r', rounded: true, border: LINE, titleBold: true, fieldButtons: false });
  return { sheets: [s.w({ 0: 16, 1: 90, 2: 112, 3: 80, 4: 120, 5: 90, 6: 80, 7: 80, 8: 90, 9: 90, 10: 40, 11: 96, 12: 112, 13: 80 }).done(), raw.done()] };
}

function marketingCalendar() {
  const s = new S('마케팅 일정', { tab: MK }).title('퍼포먼스 마케팅 캠페인 일정', '캠페인 · 채널 · 기간 · 예산 · 상태를 한눈에. 오른쪽 달력에 기간이 표시됩니다.', MK, 10);
  s.v(2, 1, '달력 시작', { bold: true }).v(2, 2, '2026-10-01', { ...DATE, ...inputSt });
  s.row(4, 1, ['캠페인', '채널', '목표', '시작', '종료', '예산', '담당', '상태'], head(MK));
  const C = [['가을 신상 런칭', '메타', '구매', 0, 20, 12000000, '김마케터', '진행 중'], ['브랜드 검색 상시', '네이버 SA', '유입', 0, 30, 5000000, '이마케터', '진행 중'], ['핼러윈 프로모션', '카카오', '구매', 24, 30, 4000000, '박마케터', '예정'], ['유튜브 인지도', '유튜브', '도달', 7, 27, 8000000, '최마케터', '진행 중'], ['리타겟팅 전환', '구글 Ads', '구매', 10, 30, 6000000, '김마케터', '예정'], ['틱톡 챌린지', '틱톡', '참여', 14, 21, 3000000, '정마케터', '기획']];
  C.forEach(([a, b, g, x, y, bud, who, st], i) => s.row(5 + i, 1, [a, b, g, `=$C$3+${x}`, `=$C$3+${y}`, bud, who, st], (j) => (j === 3 || j === 4 ? DATE : j === 5 ? WON : cellLine)));
  s.dv({ r1: 5, c1: 2, r2: 40, c2: 2, type: 'list', f1: list(CH) }).dv({ r1: 5, c1: 8, r2: 40, c2: 8, type: 'list', f1: list(['기획', '예정', '진행 중', '완료', '중단']) });
  for (let d = 0; d < 31; d++) s.v(4, 9 + d, `=$C$3+${d}`, { numFmt: 'custom', code: 'd', size: 8, align: 'center', fill: '#eef2ff' });
  s.cf({ r1: 5, c1: 9, r2: 40, c2: 39, type: 'formula', formula: '=AND(J$5>=$E6,J$5<=$F6,$I6="진행 중")', style: { fill: '#4f46e5' } });
  s.cf({ r1: 5, c1: 9, r2: 40, c2: 39, type: 'formula', formula: '=AND(J$5>=$E6,J$5<=$F6)', style: { fill: '#c7d2fe' } });
  s.cf({ r1: 5, c1: 8, r2: 40, c2: 8, type: 'formula', formula: '=I6="진행 중"', style: { color: '#15803d', bold: true } });
  s.row(2, 5, ['총 예산', '=SUM(G6:G40)'], (j) => ({ bold: true, ...(j ? WON : {}) }));
  const w = { 0: 16, 1: 140, 2: 80, 3: 50, 4: 80, 5: 80, 6: 100, 7: 70, 8: 70 };
  for (let d = 0; d < 31; d++) w[9 + d] = 20;
  return { sheets: [s.w(w).freeze(5, 2).done()] };
}

function contentCalendar() {
  const s = new S('콘텐츠 일정', { tab: '#db2777' }).title('콘텐츠 캘린더', '채널별 게시 일정 · 담당 · 상태 · 성과를 관리하세요.', '#db2777', 10);
  s.table({ name: '콘텐츠', r1: 3, c1: 1, r2: 13, c2: 10, style: 'WixelTableMinimal5' });
  s.row(3, 1, ['게시일', '요일', '채널', '주제', '포맷', '담당', '상태', '링크', '조회수', '참여율']);
  const R = rng(4);
  const ch = ['인스타그램', '블로그', '유튜브', '틱톡', '뉴스레터'];
  const fm = ['릴스', '카드뉴스', '롱폼', '쇼츠', '아티클'];
  const st = ['게시됨', '게시됨', '제작 중', '기획', '예약됨'];
  for (let i = 0; i < 10; i++) {
    const r = 4 + i;
    s.row(r, 1, [addDays('2026-10-01', i * 2), `=TEXT(B${r + 1},"aaa")`, ch[i % 5], ['신제품 소개', '사용 후기', '비하인드', '이벤트 안내', '꿀팁'][i % 5], fm[i % 5], ['김', '이', '박'][i % 3] + '에디터', st[i % 5], '', i < 5 ? Math.round(1000 + R() * 20000) : '', i < 5 ? Math.round(R() * 800) / 10000 : ''], (j) => (j === 0 ? DATE : j === 8 ? NUM : j === 9 ? PCT2 : {}));
  }
  s.dv({ r1: 4, c1: 7, r2: 200, c2: 7, type: 'list', f1: list(['기획', '제작 중', '검수', '예약됨', '게시됨']) });
  s.cf({ r1: 4, c1: 7, r2: 200, c2: 7, type: 'formula', formula: '=H5="게시됨"', style: { fill: '#dcfce7', color: '#166534' } });
  s.cf({ r1: 4, c1: 7, r2: 200, c2: 7, type: 'formula', formula: '=H5="제작 중"', style: { fill: '#fef9c3' } });
  s.cf({ r1: 4, c1: 10, r2: 200, c2: 10, type: 'bar', color: '#f472b6' });
  return { sheets: [s.w({ 0: 16, 1: 90, 2: 40, 3: 90, 4: 120, 5: 70, 6: 80, 7: 70, 8: 140, 9: 80, 10: 70 }).freeze(4).done()] };
}

function todoList() {
  const s = new S('할 일', { tab: '#16a34a' }).title('마케팅 투두 리스트', '완료 칸에 ✓를 고르면 줄이 그어지고 진행률이 올라갑니다.', '#16a34a', 7);
  s.v(2, 1, '진행률', { bold: true }).v(2, 2, '=IFERROR(COUNTIF(B6:B200,"✓")/COUNTA(C6:C200),0)', { ...PCT, bold: true, size: 14, color: '#16a34a' });
  s.cf({ r1: 2, c1: 2, r2: 2, c2: 2, type: 'bar', color: '#4ade80' });
  s.row(4, 1, ['완료', '할 일', '우선순위', '마감일', '담당', '메모'], head('#16a34a'));
  const T = [['✓', '9월 성과 리포트 발송', '높음', '2026-10-02'], ['', '메타 소재 5종 제작 요청', '높음', '2026-10-06'], ['', '네이버 SA 제외 키워드 정리', '중간', '2026-10-07'], ['', '구글 PMax 에셋 교체', '중간', '2026-10-09'], ['✓', '프로모션 랜딩 QA', '높음', '2026-10-03'], ['', '월간 예산 재배분 제안', '낮음', '2026-10-14']];
  T.forEach((t, i) => s.row(5 + i, 1, [t[0], t[1], t[2], t[3], '마케팅팀', ''], (j) => (j === 0 ? { align: 'center', ...inputSt } : j === 3 ? DATE : cellLine)));
  s.dv({ r1: 5, c1: 1, r2: 200, c2: 1, type: 'list', f1: list(['✓', '']) }).dv({ r1: 5, c1: 3, r2: 200, c2: 3, type: 'list', f1: list(['높음', '중간', '낮음']) });
  s.cf({ r1: 5, c1: 2, r2: 200, c2: 6, type: 'formula', formula: '=$B6="✓"', style: { color: '#94a3b8', strike: true } });
  s.cf({ r1: 5, c1: 4, r2: 200, c2: 4, type: 'formula', formula: '=AND($B6<>"✓",E6<TODAY(),E6<>"")', style: { color: '#dc2626', bold: true } });
  s.cf({ r1: 5, c1: 3, r2: 200, c2: 3, type: 'formula', formula: '=D6="높음"', style: { color: '#dc2626' } });
  return { sheets: [s.w({ 0: 16, 1: 50, 2: 260, 3: 70, 4: 90, 5: 80, 6: 200 }).freeze(5).done()] };
}

function promoCalendar() {
  const s = new S('프로모션 일정', { tab: '#ea580c' }).title('프로모션 일정표', '프로모션별 기간 · 할인 · 채널 · 목표 매출과 실적을 관리합니다.', '#ea580c', 11);
  s.row(3, 1, ['프로모션', '시작', '종료', '일수', '할인율', '대상', '채널', '목표 매출', '실적 매출', '달성률', '상태'], head('#ea580c'));
  const P = [['추석 선물전', '2026-09-20', '2026-10-05', 0.2, '전 상품', '전체', 80000000, 92000000], ['핼러윈 한정', '2026-10-25', '2026-10-31', 0.15, '시즌 상품', '메타 · 카카오', 20000000, ''], ['블랙프라이데이', '2026-11-20', '2026-11-30', 0.3, '전 상품', '전체', 150000000, ''], ['연말 감사제', '2026-12-15', '2026-12-31', 0.25, '베스트', '네이버 · 구글', 90000000, '']];
  P.forEach(([n, a, b, d, tg, ch, goal, act], i) => { const r = 4 + i; s.row(r, 1, [n, a, b, `=D${r + 1}-C${r + 1}+1`, d, tg, ch, goal, act, `=IFERROR(J${r + 1}/I${r + 1},"")`, `=IF(TODAY()<C${r + 1},"예정",IF(TODAY()>D${r + 1},"종료","진행 중"))`], (j) => (j === 1 || j === 2 ? DATE : j === 4 ? PCT : j === 7 || j === 8 ? WON : j === 9 ? PCT : cellLine)); });
  s.cf({ r1: 4, c1: 10, r2: 30, c2: 10, type: 'icons', icons: '3Arrows' });
  s.cf({ r1: 4, c1: 11, r2: 30, c2: 11, type: 'formula', formula: '=L5="진행 중"', style: { fill: '#ffedd5', color: '#c2410c', bold: true } });
  s.chart({ type: 'bar', title: '프로모션별 목표 · 실적', range: { r1: 3, c1: 1, r2: 7, c2: 1 }, series: [{ name: { text: '목표' }, cat: { r1: 4, c1: 1, r2: 7, c2: 1 }, val: { r1: 4, c1: 8, r2: 7, c2: 8 } }, { name: { text: '실적' }, cat: { r1: 4, c1: 1, r2: 7, c2: 1 }, val: { r1: 4, c1: 9, r2: 7, c2: 9 } }], x: 60, y: 210, w: 560, h: 280, palette: 'vivid' });
  return { sheets: [s.w({ 0: 16, 1: 130, 2: 90, 3: 90, 4: 50, 5: 60, 6: 80, 7: 100, 8: 110, 9: 110, 10: 70, 11: 70 }).done()] };
}

/** 효율 분석 공통: 행 = 항목, 지표 계산 + 등급 + 추천 */
function efficiencySheet(name, title, sub, firstCols, rows, color, extraRules = []) {
  const s = new S(name, { tab: color }).title(title, sub, color, 14);
  const hdr = [...firstCols, '노출수', '클릭수', '비용', '전환수', '매출', 'CTR', 'CPC', 'CVR', 'CPA', 'ROAS', '판정'];
  const k = firstCols.length;
  s.v(2, 1, '목표 ROAS', { bold: true }).v(2, 2, 3, { ...PCT, ...inputSt }).v(2, 3, '목표 CPA', { bold: true }).v(2, 4, 30000, { ...WON, ...inputSt });
  s.table({ name, r1: 4, c1: 1, r2: 4 + rows.length, c2: hdr.length, style: 'WixelTableModern2' });
  s.row(4, 1, hdr);
  rows.forEach((row, i) => {
    const r = 5 + i;
    const R1 = r + 1;
    const L = (o) => `${col(1 + k + o)}${R1}`;
    s.row(r, 1, [...row, `=IFERROR(${L(1)}/${L(0)},0)`, `=IFERROR(${L(2)}/${L(1)},0)`, `=IFERROR(${L(3)}/${L(1)},0)`, `=IFERROR(${L(2)}/${L(3)},0)`, `=IFERROR(${L(4)}/${L(2)},0)`,
      `=IF(${L(3)}=0,IF(${L(2)}>$E$3,"제외 검토","관찰"),IF(${L(9)}>=$C$3,"확대",IF(${L(8)}<=$E$3,"유지","효율 개선")))`],
    (j) => (j < k ? {} : j < k + 5 ? (j === k + 2 || j === k + 4 ? WON : NUM) : j === k + 5 || j === k + 7 ? PCT2 : j === k + 9 ? PCT : j === k + 10 ? { bold: true, align: 'center' } : WON));
  });
  const last = 4 + rows.length;
  const jc = 1 + k + 10;
  s.cf({ r1: 5, c1: jc, r2: last, c2: jc, type: 'formula', formula: `=${col(jc)}6="확대"`, style: { fill: '#dcfce7', color: '#166534' } });
  s.cf({ r1: 5, c1: jc, r2: last, c2: jc, type: 'formula', formula: `=${col(jc)}6="제외 검토"`, style: { fill: '#fee2e2', color: '#b91c1c' } });
  s.cf({ r1: 5, c1: jc, r2: last, c2: jc, type: 'formula', formula: `=${col(jc)}6="효율 개선"`, style: { fill: '#fef9c3', color: '#a16207' } });
  s.cf({ r1: 5, c1: 1 + k + 9, r2: last, c2: 1 + k + 9, type: 'scale', colors: ['#fb7185', '#fde68a', '#34d399'] });
  s.cf({ r1: 5, c1: 1 + k + 2, r2: last, c2: 1 + k + 2, type: 'bar', color: '#818cf8' });
  for (const rule of extraRules) s.cf(rule);
  const w = { 0: 16 };
  firstCols.forEach((_, i) => { w[1 + i] = i === k - 1 ? 180 : 90; });
  for (let i = 0; i < 11; i++) w[1 + k + i] = 80;
  w[1 + k + 10] = 80;
  return s.w(w).freeze(5, 1 + k);
}

function creativeEfficiency() {
  const R = rng(17);
  const types = ['이미지', '영상', '캐러셀', '컬렉션'];
  const rows = Array.from({ length: 16 }, (_, i) => {
    const imp = Math.round(20000 + R() * 200000);
    const clk = Math.round(imp * (0.004 + R() * 0.03));
    const cost = Math.round(clk * (300 + R() * 700));
    const conv = Math.round(clk * R() * 0.06);
    return [['메타', '구글', '카카오', '틱톡'][i % 4], types[i % 4], `소재_${String(i + 1).padStart(2, '0')}_${['혜택강조', '후기', 'UGC', '제품컷'][i % 4]}`, imp, clk, cost, conv, conv * Math.round(30000 + R() * 40000)];
  });
  const s = efficiencySheet('소재효율', '광고 소재 효율 분석', '소재별 CTR · CVR · CPA · ROAS와 판정(확대/유지/효율 개선/제외 검토)을 자동으로 계산합니다.', ['매체', '유형', '소재명'], rows, '#9333ea');
  s.chart({ type: 'bubble', title: 'CTR × CVR (크기 = 비용)', series: [{ name: { text: '소재' }, x: { r1: 5, c1: 9, r2: 20, c2: 9 }, val: { r1: 5, c1: 11, r2: 20, c2: 11 }, size: { r1: 5, c1: 6, r2: 20, c2: 6 } }], x: 60, y: 470, w: 520, h: 300, palette: 'modern', legend: 'none' });
  return { sheets: [s.done()] };
}

function searchTermEfficiency() {
  const R = rng(23);
  const terms = ['브랜드명', '브랜드명 후기', '브랜드명 할인', '아기 영양제', '유아 두뇌 영양제', '어린이 오메가3', 'DHA 추천', '키 성장 영양제', '무료 샘플', '영양제 부작용', '가격 비교', '공식몰', '쿠폰', '성분 분석', '약국 영양제', '해외 직구'];
  const rows = terms.map((t, i) => {
    const imp = Math.round(500 + R() * 30000);
    const clk = Math.round(imp * (0.01 + R() * 0.08));
    const cost = Math.round(clk * (150 + R() * 1200));
    const conv = i % 5 === 3 ? 0 : Math.round(clk * R() * 0.09);
    return [i < 3 || i === 11 ? '브랜드' : '일반', t, imp, clk, cost, conv, conv * Math.round(40000 + R() * 30000)];
  });
  const s = efficiencySheet('검색어효율', '검색어 효율 분석', '전환 없는 고비용 검색어는 [제외 검토], 목표 ROAS 이상은 [확대]로 표시됩니다.', ['구분', '검색어'], rows, '#0891b2');
  s.v(2, 6, '제외 후보 수', { bold: true }).v(2, 7, '=COUNTIF(검색어효율[판정],"제외 검토")', { bold: true, color: '#b91c1c' });
  s.v(2, 8, '제외 시 절감액', { bold: true }).v(2, 9, '=SUMIF(검색어효율[판정],"제외 검토",검색어효율[비용])', { ...WON, bold: true, color: '#b91c1c' });
  return { sheets: [s.done()] };
}

function budgetPacing() {
  const s = new S('예산 페이싱', { tab: '#0d9488' }).title('월 광고 예산 페이싱', '일별 소진액을 입력하면 이상적인 소진 곡선 대비 페이스와 월말 예상 소진액을 계산합니다.', '#0d9488', 9);
  const inp = [['월 예산', 30000000, WON], ['시작일', '2026-10-01', DATE], ['종료일', '2026-10-31', DATE]];
  inp.forEach(([l, v, f], i) => s.row(2 + i, 1, [l, v], (j) => (j ? { ...f, ...inputSt } : { bold: true })));
  s.row(2, 4, ['누적 소진', '=SUM(D9:D39)'], (j) => (j ? { ...WON, bold: true } : { bold: true }));
  s.row(3, 4, ['소진율', '=F3/C3'], (j) => (j ? { ...PCT, bold: true } : { bold: true }));
  s.row(4, 4, ['월말 예상', '=IFERROR(F3/COUNT(D9:D39)*(C5-C4+1),0)'], (j) => (j ? { ...WON, bold: true, color: '#0d9488' } : { bold: true }));
  s.row(5, 4, ['페이스', '=IFERROR(F5/C3,0)'], (j) => (j ? { ...PCT, bold: true } : { bold: true }));
  s.cf({ r1: 5, c1: 5, r2: 5, c2: 5, type: 'formula', formula: '=OR(F6>1.1,F6<0.9)', style: { fill: '#fee2e2', color: '#b91c1c' } });
  s.row(7, 1, ['날짜', '요일', '일 소진액', '누적 소진', '이상 누적', '차이'], head('#0d9488'));
  const R = rng(31);
  for (let i = 0; i < 31; i++) {
    const r = 8 + i;
    s.row(r, 1, [`=$C$4+${i}`, `=TEXT(B${r + 1},"aaa")`, i < 18 ? Math.round((800000 + R() * 500000) / 1000) * 1000 : '', `=IF(D${r + 1}="","",SUM($D$9:D${r + 1}))`, `=$C$3/($C$5-$C$4+1)*${i + 1}`, `=IF(E${r + 1}="","",E${r + 1}-F${r + 1})`], (j) => (j === 0 ? DATE : j === 1 ? { align: 'center' } : j === 2 ? { ...WON, ...inputSt } : WON));
  }
  s.cf({ r1: 8, c1: 6, r2: 38, c2: 6, type: 'formula', formula: '=AND(G9<>"",G9<0)', style: { color: '#2563eb' } });
  s.cf({ r1: 8, c1: 6, r2: 38, c2: 6, type: 'formula', formula: '=AND(G9<>"",G9>0)', style: { color: '#dc2626' } });
  s.chart({ type: 'line', title: '누적 소진 vs 이상 곡선', series: [{ name: { text: '누적 소진' }, cat: { r1: 8, c1: 1, r2: 38, c2: 1 }, val: { r1: 8, c1: 4, r2: 38, c2: 4 } }, { name: { text: '이상 누적' }, cat: { r1: 8, c1: 1, r2: 38, c2: 1 }, val: { r1: 8, c1: 5, r2: 38, c2: 5 } }], x: 720, y: 60, w: 520, h: 300, palette: 'modern', marker: 'none' });
  return { sheets: [s.w({ 0: 16, 1: 90, 2: 110, 3: 110, 4: 110, 5: 110, 6: 100 }).freeze(8).done()] };
}

function mediaMix() {
  const s = new S('미디어 믹스', { tab: '#7c3aed' }).title('채널 · 미디어 믹스 플래너', '총 예산과 채널별 배분 비율 · 예상 CPC · CVR · 객단가를 넣으면 예상 성과가 계산됩니다.', '#7c3aed', 11);
  s.v(2, 1, '총 예산', { bold: true }).v(2, 2, 50000000, { ...WON, ...inputSt });
  s.row(4, 1, ['채널', '배분 비율', '예산', '예상 CPC', '예상 클릭', '예상 CVR', '예상 전환', '객단가', '예상 매출', '예상 ROAS', '예상 CPA'], head('#7c3aed'));
  const M = [['네이버 SA', 0.3, 650, 0.035, 52000], ['구글 Ads', 0.2, 800, 0.03, 52000], ['메타', 0.25, 550, 0.022, 48000], ['카카오', 0.1, 450, 0.018, 45000], ['유튜브', 0.1, 300, 0.008, 50000], ['틱톡', 0.05, 250, 0.01, 42000]];
  M.forEach(([ch, p, cpc, cvr, aov], i) => { const r = 5 + i; const R1 = r + 1; s.row(r, 1, [ch, p, `=$C$3*C${R1}`, cpc, `=IFERROR(D${R1}/E${R1},0)`, cvr, `=F${R1}*G${R1}`, aov, `=H${R1}*I${R1}`, `=IFERROR(J${R1}/D${R1},0)`, `=IFERROR(D${R1}/H${R1},0)`], (j) => (j === 1 || j === 5 ? { ...(j === 1 ? PCT : PCT2), ...inputSt } : j === 3 || j === 7 ? { ...WON, ...inputSt } : j === 9 ? PCT : j === 0 ? { bold: true } : j === 4 || j === 6 ? NUM : WON)); });
  s.row(11, 1, ['합계', '=SUM(C6:C11)', '=SUM(D6:D11)', '=IFERROR(D12/F12,0)', '=SUM(F6:F11)', '=IFERROR(H12/F12,0)', '=SUM(H6:H11)', '=IFERROR(J12/H12,0)', '=SUM(J6:J11)', '=IFERROR(J12/D12,0)', '=IFERROR(D12/H12,0)'], (j) => ({ bold: true, fill: '#ede9fe', ...(j === 1 || j === 9 ? PCT : j === 5 ? PCT2 : j === 4 || j === 6 ? NUM : j === 0 ? {} : WON) }));
  s.cf({ r1: 11, c1: 2, r2: 11, c2: 2, type: 'formula', formula: '=ROUND(C12,4)<>1', style: { fill: '#fee2e2', color: '#b91c1c' } });
  s.cf({ r1: 5, c1: 10, r2: 10, c2: 10, type: 'scale', colors: ['#fb7185', '#fde68a', '#34d399'] });
  s.chart({ type: 'pie', title: '예산 배분', series: [{ name: { text: '예산' }, cat: { r1: 5, c1: 1, r2: 10, c2: 1 }, val: { r1: 5, c1: 3, r2: 10, c2: 3 } }], x: 60, y: 290, w: 400, h: 280, palette: 'modern', legend: 'r' });
  s.chart({ type: 'column', title: '채널별 예상 ROAS', series: [{ name: { text: 'ROAS' }, cat: { r1: 5, c1: 1, r2: 10, c2: 1 }, val: { r1: 5, c1: 10, r2: 10, c2: 10 }, }], seriesFmt: [{ numFmt: '0%', labels: true }], x: 480, y: 290, w: 460, h: 280, palette: 'modern', legend: 'none' });
  return { sheets: [s.w({ 0: 16, 1: 100, 2: 80, 3: 120, 4: 80, 5: 90, 6: 80, 7: 80, 8: 80, 9: 120, 10: 80, 11: 90 }).done()] };
}

function abTest() {
  const s = new S('A/B 테스트', { tab: '#2563eb' }).title('A/B 테스트 유의성 계산기', '방문자 수와 전환 수를 입력하면 전환율 · 상승률 · 유의 확률을 계산합니다 (양측 z-검정).', '#2563eb', 6);
  s.row(3, 1, ['', '방문자', '전환', '전환율', '표준오차'], head('#2563eb'));
  s.row(4, 1, ['A (기존)', 12000, 360, '=D5/C5', '=SQRT(E5*(1-E5)/C5)'], (j) => (j === 1 || j === 2 ? { ...NUM, ...inputSt } : j === 3 ? PCT2 : j === 4 ? { numFmt: 'custom', code: '0.0000' } : { bold: true }));
  s.row(5, 1, ['B (변형)', 11800, 413, '=D6/C6', '=SQRT(E6*(1-E6)/C6)'], (j) => (j === 1 || j === 2 ? { ...NUM, ...inputSt } : j === 3 ? PCT2 : j === 4 ? { numFmt: 'custom', code: '0.0000' } : { bold: true }));
  const res = [['상승률 (B vs A)', '=E6/E5-1', PCT2], ['z 점수', '=(E6-E5)/SQRT(F5^2+F6^2)', { numFmt: 'custom', code: '0.00' }], ['p 값', '=2*(1-NORM.S.DIST(ABS(C9),TRUE))', { numFmt: 'custom', code: '0.0000' }], ['신뢰 수준', 0.95, { ...PCT, ...inputSt }], ['결과', '=IF(C10<1-C11,IF(C8>0,"B 승 (유의함)","A 승 (유의함)"),"아직 유의하지 않음 — 표본을 더 모으세요")', { bold: true }]];
  res.forEach(([l, f, st], i) => s.row(7 + i, 1, [l, f], (j) => (j ? st : { bold: true })));
  s.cf({ r1: 11, c1: 2, r2: 11, c2: 2, type: 'formula', formula: '=LEFT(C12,1)="B"', style: { fill: '#dcfce7', color: '#166534' } });
  s.v(13, 1, '필요 표본 수 (그룹당, 검정력 80%)', { bold: true }).v(13, 2, '=ROUNDUP((1.96+0.84)^2*(E5*(1-E5)+E5*(1+C8)*(1-E5*(1+C8)))/(E5*C8)^2,0)', NUM);
  s.chart({ type: 'column', title: '전환율 비교', series: [{ name: { text: '전환율' }, cat: { r1: 4, c1: 1, r2: 5, c2: 1 }, val: { r1: 4, c1: 4, r2: 5, c2: 4 } }], seriesFmt: [{ labels: true, numFmt: '0.00%' }], x: 520, y: 60, w: 360, h: 260, legend: 'none', palette: 'modern' });
  return { sheets: [s.w({ 0: 16, 1: 200, 2: 110, 3: 90, 4: 90, 5: 90 }).done()] };
}

function utmBuilder() {
  const s = new S('UTM 빌더', { tab: '#475569' }).title('UTM 링크 빌더', '기본 URL과 UTM 값을 입력하면 추적 링크가 만들어집니다. 공백은 자동으로 %20 으로 바뀝니다.', '#475569', 8);
  s.table({ name: 'UTM', r1: 3, c1: 1, r2: 9, c2: 8, style: 'WixelTableMinimal8' });
  s.row(3, 1, ['기본 URL', 'utm_source', 'utm_medium', 'utm_campaign', 'utm_content', 'utm_term', '완성 URL', '짧은 메모']);
  const U = [['https://wixel.shop/event', 'naver', 'cpc', 'fall_launch', 'text_a', '영양제'], ['https://wixel.shop/event', 'meta', 'paid_social', 'fall_launch', 'video_15s', ''], ['https://wixel.shop/best', 'google', 'cpc', 'pmax_best', '', ''], ['https://wixel.shop/best', 'kakao', 'display', 'bizboard', 'banner_b', ''], ['https://wixel.shop/', 'newsletter', 'email', '1015_letter', 'cta_top', ''], ['https://wixel.shop/', 'instagram', 'influencer', 'collab_oct', 'reels_01', '']];
  const enc = (ref) => `SUBSTITUTE(SUBSTITUTE(TRIM(${ref})," ","%20"),"&","%26")`;
  U.forEach((u, i) => { const R1 = 5 + i; s.row(4 + i, 1, [...u, `=B${R1}&IF(ISNUMBER(FIND("?",B${R1})),"&","?")&"utm_source="&${enc(`C${R1}`)}&"&utm_medium="&${enc(`D${R1}`)}&"&utm_campaign="&${enc(`E${R1}`)}&IF(F${R1}<>"","&utm_content="&${enc(`F${R1}`)},"")&IF(G${R1}<>"","&utm_term="&${enc(`G${R1}`)},"")`, ''], (j) => (j === 6 ? { color: '#2563eb' } : {})); });
  s.dv({ r1: 4, c1: 3, r2: 200, c2: 3, type: 'list', f1: list(['cpc', 'display', 'paid_social', 'email', 'influencer', 'organic', 'referral']) });
  return { sheets: [s.w({ 0: 16, 1: 190, 2: 90, 3: 100, 4: 110, 5: 90, 6: 80, 7: 520, 8: 120 }).done()] };
}

function kpiTracker() {
  const s = new S('KPI 트래커', { tab: '#b45309' }).title('월간 KPI 목표 트래커', '목표와 실적을 입력하면 달성률 · 신호등 · 남은 목표가 계산됩니다.', '#b45309', 9);
  s.v(2, 1, '경과율 (오늘 기준)', { bold: true }).v(2, 2, '=DAY(TODAY())/DAY(EOMONTH(TODAY(),0))', { ...PCT, bold: true });
  s.row(4, 1, ['KPI', '단위', '목표', '실적', '달성률', '페이스 대비', '남은 목표', '담당'], head('#b45309'));
  const K = [['매출', '원', 300000000, 186000000], ['광고비', '원', 60000000, 38000000], ['ROAS', '%', 5, 4.9], ['신규 고객', '명', 2500, 1710], ['CPA', '원', 25000, 22400], ['앱 설치', '건', 8000, 4200], ['CTR', '%', 0.015, 0.0171], ['재구매율', '%', 0.32, 0.29]];
  K.forEach(([k, u, g, a], i) => { const R1 = 6 + i; const fmt = u === '%' ? PCT : u === '원' ? WON : NUM; s.row(5 + i, 1, [k, u, g, a, `=IFERROR(${k === 'CPA' || k === '광고비' ? `D${R1}/E${R1}` : `E${R1}/D${R1}`},0)`, `=IFERROR(F${R1}/$C$3,0)`, `=MAX(0,D${R1}-E${R1})`, '마케팅팀'], (j) => (j === 2 || j === 3 ? { ...fmt, ...inputSt } : j === 4 || j === 5 ? PCT : j === 6 ? fmt : { bold: j === 0 })); });
  s.cf({ r1: 5, c1: 5, r2: 12, c2: 5, type: 'bar', color: '#fbbf24' });
  s.cf({ r1: 5, c1: 6, r2: 12, c2: 6, type: 'icons', icons: '3TrafficLights1' });
  return { sheets: [s.w({ 0: 16, 1: 110, 2: 50, 3: 120, 4: 120, 5: 100, 6: 100, 7: 110, 8: 90 }).done()] };
}

function funnelAnalysis() {
  const s = new S('전환 퍼널', { tab: '#0284c7' }).title('전환 퍼널 분석', '단계별 인원을 입력하면 단계 전환율 · 이탈률 · 전체 전환율을 계산합니다.', '#0284c7', 7);
  s.row(3, 1, ['단계', '인원', '단계 전환율', '이탈률', '전체 대비', '개선 시 +10% 효과'], head('#0284c7'));
  const F = [['광고 노출', 1250000], ['클릭', 28400], ['상품 조회', 19800], ['장바구니', 4100], ['결제 시작', 2300], ['구매 완료', 1480]];
  F.forEach(([n, v], i) => { const R1 = 5 + i; s.row(4 + i, 1, [n, v, i ? `=C${R1}/C${R1 - 1}` : '', i ? `=1-D${R1}` : '', `=C${R1}/$C$5`, i ? `=$C$10*1.1-$C$10` : ''], (j) => (j === 1 ? { ...NUM, ...inputSt } : j >= 2 && j <= 4 ? PCT2 : j === 5 ? NUM : { bold: true })); });
  s.cf({ r1: 5, c1: 4, r2: 9, c2: 4, type: 'scale', colors: ['#ffffff', '#fb7185'] });
  s.chart({ type: 'funnel', title: '구매 퍼널', series: [{ name: { text: '인원' }, cat: { r1: 5, c1: 1, r2: 9, c2: 1 }, val: { r1: 5, c1: 2, r2: 9, c2: 2 } }], x: 620, y: 60, w: 460, h: 300, palette: 'modern' });
  return { sheets: [s.w({ 0: 16, 1: 110, 2: 110, 3: 100, 4: 80, 5: 90, 6: 130 }).done()] };
}

function cohortRetention() {
  const s = new S('코호트 리텐션', { tab: '#be123c' }).title('코호트 리텐션 분석', '가입(첫 구매) 월별 고객 수와 경과 월의 잔존 고객 수를 입력하면 잔존율 히트맵이 만들어집니다.', '#be123c', 14);
  const months = ['2026-03', '2026-04', '2026-05', '2026-06', '2026-07', '2026-08', '2026-09'];
  s.row(3, 1, ['코호트', '고객 수', ...Array.from({ length: 7 }, (_, i) => `M${i}`)], head('#be123c'));
  s.row(12, 1, ['코호트', '잔존율', ...Array.from({ length: 7 }, (_, i) => `M${i}`)], head('#be123c'));
  const R = rng(41);
  months.forEach((m, i) => {
    const n = Math.round(800 + R() * 600);
    const vals = [];
    let cur = n;
    for (let k = 0; k < months.length - i; k++) { vals.push(cur); cur = Math.round(cur * (k === 0 ? 0.42 + R() * 0.1 : 0.72 + R() * 0.12)); }
    s.row(4 + i, 1, [m, n, ...vals], (j) => (j >= 1 ? { ...NUM, ...inputSt } : { bold: true }));
    s.row(13 + i, 1, [m, ''], { bold: true });
    for (let k = 0; k < 7; k++) s.v(13 + i, 3 + k, `=IF(${col(3 + k)}${5 + i}="","",${col(3 + k)}${5 + i}/$C${5 + i})`, PCT);
  });
  s.cf({ r1: 13, c1: 3, r2: 19, c2: 9, type: 'scale', colors: ['#fff1f2', '#fb7185', '#9f1239'] });
  s.row(21, 1, ['평균 잔존율', '', ...Array.from({ length: 7 }, (_, k) => `=IFERROR(AVERAGE(${col(3 + k)}14:${col(3 + k)}20),"")`)], (j) => (j >= 2 ? { ...PCT, bold: true } : { bold: true }));
  s.chart({ type: 'line', title: '평균 리텐션 곡선', series: [{ name: { text: '잔존율' }, cat: { r1: 12, c1: 3, r2: 12, c2: 9 }, val: { r1: 21, c1: 3, r2: 21, c2: 9 } }], x: 760, y: 60, w: 440, h: 280, palette: 'modern', legend: 'none' });
  return { sheets: [s.w({ 0: 16, 1: 90, 2: 80 }).done()] };
}

function keywordPlanner() {
  const s = new S('키워드 플래너', { tab: '#16a34a' }).title('검색 키워드 플래너', '월간 검색량 · 예상 CPC · 예상 CTR/CVR로 키워드별 예상 비용 · 전환 · 우선순위를 계산합니다.', '#16a34a', 12);
  s.row(2, 1, ['평균 객단가', 45000], (j) => (j ? { ...WON, ...inputSt } : { bold: true }));
  s.table({ name: '키워드', r1: 4, c1: 1, r2: 16, c2: 12, style: 'WixelTableModern4' });
  s.row(4, 1, ['키워드', '의도', '월간 검색량', '경쟁도', '예상 CPC', '예상 CTR', '예상 CVR', '예상 클릭', '예상 비용', '예상 전환', '예상 ROAS', '우선순위']);
  const K = [['유아 영양제', '구매'], ['어린이 오메가3', '구매'], ['아기 DHA', '구매'], ['영양제 추천', '탐색'], ['두뇌 영양제', '탐색'], ['키 성장 영양제', '구매'], ['어린이 비타민', '구매'], ['영양제 먹는 시간', '정보'], ['유산균 추천', '탐색'], ['영양제 부작용', '정보'], ['브랜드명', '브랜드'], ['브랜드명 할인', '브랜드']];
  const R = rng(51);
  K.forEach(([k, it], i) => { const r = 5 + i; const R1 = r + 1; s.row(r, 1, [k, it, Math.round(500 + R() * 60000), ['낮음', '중간', '높음'][Math.floor(R() * 3)], Math.round(150 + R() * 1500), Math.round((0.01 + R() * 0.05) * 1000) / 1000, it === '정보' ? 0.004 : it === '브랜드' ? 0.08 : 0.02 + Math.round(R() * 30) / 1000, `=ROUND(D${R1}*G${R1},0)`, `=I${R1}*F${R1}`, `=I${R1}*H${R1}`, `=IFERROR(K${R1}*$C$3/J${R1},0)`, `=IF(L${R1}>=4,"A",IF(L${R1}>=2,"B","C"))`], (j) => (j === 2 || j === 7 || j === 9 ? NUM : j === 4 || j === 8 ? WON : j === 5 || j === 6 ? PCT2 : j === 10 ? PCT : j === 11 ? { bold: true, align: 'center' } : {})); });
  s.cf({ r1: 5, c1: 12, r2: 16, c2: 12, type: 'formula', formula: '=M6="A"', style: { fill: '#dcfce7', color: '#166534' } });
  s.cf({ r1: 5, c1: 3, r2: 16, c2: 3, type: 'bar', color: '#4ade80' });
  s.dv({ r1: 5, c1: 2, r2: 200, c2: 2, type: 'list', f1: list(['브랜드', '구매', '탐색', '정보']) });
  return { sheets: [s.w({ 0: 16, 1: 150, 2: 60, 3: 90, 4: 60, 5: 80, 6: 70, 7: 70, 8: 80, 9: 100, 10: 70, 11: 80, 12: 70 }).freeze(5, 2).done()] };
}

function mediaPlan() {
  const s = new S('미디어 플랜', { tab: '#1e293b' }).title('매체 집행 계획서 (미디어 플랜)', '매체 · 상품 · 기간 · 단가 · 수량을 입력하면 수수료 · VAT 포함 총액이 계산됩니다.', '#1e293b', 12);
  s.row(2, 1, ['광고주', '(주)브랜드', '', '대행 수수료율', 0.15], (j) => (j === 1 ? inputSt : j === 4 ? { ...PCT, ...inputSt } : { bold: true }));
  s.row(4, 1, ['매체', '상품', '과금 방식', '시작', '종료', '단가', '수량', '집행 금액', '수수료', 'VAT', '총액', '예상 노출/클릭'], head('#1e293b'));
  const P = [['네이버', '파워링크', 'CPC', '2026-10-01', '2026-10-31', 700, 20000, '클릭 20,000'], ['네이버', '브랜드검색', '기간', '2026-10-01', '2026-10-31', 880000, 1, '-'], ['메타', '피드 · 릴스', 'CPM', '2026-10-01', '2026-10-31', 6500, 3000, '노출 3,000,000'], ['유튜브', '인스트림', 'CPV', '2026-10-10', '2026-10-31', 40, 200000, '조회 200,000'], ['카카오', '비즈보드', 'CPM', '2026-10-15', '2026-10-31', 4000, 1500, '노출 1,500,000']];
  P.forEach((p, i) => { const R1 = 6 + i; s.row(5 + i, 1, [p[0], p[1], p[2], p[3], p[4], p[5], p[6], `=G${R1}*H${R1}`, `=I${R1}*$F$3`, `=(I${R1}+J${R1})*0.1`, `=I${R1}+J${R1}+K${R1}`, p[7]], (j) => (j === 3 || j === 4 ? DATE : j === 5 || j >= 7 && j <= 10 ? WON : j === 6 ? NUM : {})); });
  s.row(10, 1, ['합계', '', '', '', '', '', '', '=SUM(I6:I10)', '=SUM(J6:J10)', '=SUM(K6:K10)', '=SUM(L6:L10)'], (j) => ({ bold: true, fill: '#e2e8f0', ...(j >= 7 ? WON : {}) }));
  s.dv({ r1: 5, c1: 3, r2: 40, c2: 3, type: 'list', f1: list(['CPC', 'CPM', 'CPV', 'CPA', '기간', '고정']) });
  return { sheets: [s.w({ 0: 16, 1: 70, 2: 100, 3: 70, 4: 90, 5: 90, 6: 80, 7: 80, 8: 110, 9: 100, 10: 90, 11: 120, 12: 130 }).done()] };
}

function roasSimulator() {
  const s = new S('ROAS 시뮬레이터', { tab: '#059669' }).title('ROAS · 손익 시뮬레이터', '예산 · CPC · CVR · 객단가 · 원가율을 바꿔 보세요. 아래 민감도 표가 CPC × CVR 조합별 ROAS를 보여 줍니다.', '#059669', 10);
  const I = [['광고 예산', 20000000, WON], ['평균 CPC', 700, WON], ['CVR', 0.025, PCT2], ['객단가', 48000, WON], ['원가율', 0.45, PCT], ['기타 변동비율', 0.1, PCT]];
  I.forEach(([l, v, f], i) => s.row(3 + i, 1, [l, v], (j) => (j ? { ...f, ...inputSt } : { bold: true })));
  const O = [['클릭', '=C4/C5', NUM], ['전환', '=C10*C6', NUM], ['매출', '=C11*C7', WON], ['ROAS', '=C12/C4', PCT], ['CPA', '=C4/C11', WON], ['공헌 이익', '=C12*(1-C8-C9)-C4', WON], ['손익분기 ROAS', '=1/(1-C8-C9)', PCT]];
  O.forEach(([l, f, st], i) => s.row(9 + i, 1, [l, f], (j) => (j ? { ...st, bold: true } : { bold: true })));
  s.cf({ r1: 14, c1: 2, r2: 14, c2: 2, type: 'formula', formula: '=C15<0', style: { color: '#dc2626' } });
  s.v(17, 1, 'ROAS 민감도 (행 = CPC, 열 = CVR)', { bold: true, color: '#059669' });
  const cvrs = [0.01, 0.015, 0.02, 0.025, 0.03, 0.035, 0.04];
  const cpcs = [400, 500, 600, 700, 800, 900, 1000, 1200];
  s.row(18, 1, ['CPC \\ CVR', ...cvrs], (j) => (j ? { ...PCT, ...head('#059669') } : head('#059669')));
  cpcs.forEach((cpc, i) => {
    s.v(19 + i, 1, cpc, { ...WON, bold: true, fill: '#ecfdf5' });
    cvrs.forEach((_, k) => s.v(19 + i, 2 + k, `=${col(2 + k)}$19*$C$7/$B${20 + i}`, PCT));
  });
  s.cf({ r1: 19, c1: 2, r2: 26, c2: 8, type: 'scale', colors: ['#fb7185', '#fde68a', '#34d399'] });
  s.cf({ r1: 19, c1: 2, r2: 26, c2: 8, type: 'formula', formula: '=C20<$C$16', style: { color: '#b91c1c' } });
  return { sheets: [s.w({ 0: 16, 1: 130, 2: 100, 3: 80, 4: 80, 5: 80, 6: 80, 7: 80, 8: 80 }).done()] };
}

function competitorTracker() {
  const s = new S('경쟁사 모니터링', { tab: '#64748b' }).title('경쟁사 광고 · 프로모션 모니터링', '경쟁사 광고 소재 · 프로모션 · 가격 변화를 기록합니다.', '#475569', 9);
  s.table({ name: '경쟁사', r1: 3, c1: 1, r2: 9, c2: 9, style: 'WixelTableMinimal1' });
  s.row(3, 1, ['발견일', '경쟁사', '채널', '유형', '메시지 · 혜택', '가격', '할인율', '링크', '대응 방안']);
  [['2026-09-28', 'A사', '메타', '영상', '첫 구매 50% 할인', 39000, 0.5, '', '첫 구매 쿠폰 테스트'], ['2026-09-29', 'B사', '네이버', '브랜드검색', '추석 기획전', 45000, 0.2, '', '브검 문구 교체'], ['2026-09-30', 'A사', '유튜브', '인플루언서', '성분 비교 리뷰', '', '', '', '비교 콘텐츠 제작'], ['2026-10-01', 'C사', '카카오', '비즈보드', '1+1 이벤트', 52000, 0.5, '', '관찰'], ['2026-10-02', 'B사', '구글', 'PMax', '무료 배송', 45000, 0, '', '배송비 정책 검토'], ['2026-10-03', 'C사', '틱톡', '챌린지', 'UGC 챌린지', '', '', '', 'UGC 캠페인 검토']].forEach((row, i) => s.row(4 + i, 1, row, (j) => (j === 0 ? DATE : j === 5 ? WON : j === 6 ? PCT : {})));
  s.cf({ r1: 4, c1: 7, r2: 100, c2: 7, type: 'formula', formula: '=G5>=0.5', style: { color: '#dc2626', bold: true } });
  return { sheets: [s.w({ 0: 16, 1: 90, 2: 70, 3: 70, 4: 90, 5: 200, 6: 80, 7: 70, 8: 120, 9: 160 }).done()] };
}

function influencerTracker() {
  const s = new S('인플루언서', { tab: '#c026d3' }).title('인플루언서 캠페인 트래커', '비용 · 조회 · 참여 · 전환을 입력하면 CPV · CPE · CPA · ROAS가 계산됩니다.', '#c026d3', 13);
  s.table({ name: '인플루언서', r1: 3, c1: 1, r2: 9, c2: 13, style: 'WixelTableModern7' });
  s.row(3, 1, ['인플루언서', '플랫폼', '팔로워', '게시일', '비용', '조회수', '참여수', '전환', '매출', 'CPV', 'CPE', 'CPA', 'ROAS']);
  const R = rng(61);
  ['@daily_mom', '@healthy_kids', '@parenting_tip', '@babyfood_lab', '@momlog', '@kidsfit'].forEach((h, i) => {
    const r = 4 + i; const R1 = r + 1;
    const f = Math.round(20000 + R() * 300000); const cost = Math.round((f * (8 + R() * 10)) / 10000) * 10000; const view = Math.round(f * (0.3 + R() * 1.2)); const eng = Math.round(view * (0.02 + R() * 0.06)); const conv = Math.round(eng * R() * 0.08);
    s.row(r, 1, [h, ['인스타그램', '유튜브', '틱톡'][i % 3], f, addDays('2026-09-10', i * 4), cost, view, eng, conv, conv * 45000, `=IFERROR(F${R1}/G${R1},0)`, `=IFERROR(F${R1}/H${R1},0)`, `=IFERROR(F${R1}/I${R1},0)`, `=IFERROR(J${R1}/F${R1},0)`], (j) => (j === 3 ? DATE : j === 4 || j === 8 || j >= 9 && j <= 11 ? WON : j === 12 ? PCT : j === 2 || j >= 5 && j <= 7 ? NUM : {}));
  });
  s.cf({ r1: 4, c1: 13, r2: 9, c2: 13, type: 'scale', colors: ['#fb7185', '#fde68a', '#34d399'] });
  return { sheets: [s.w({ 0: 16, 1: 120, 2: 80, 3: 80, 4: 90, 5: 90, 6: 80, 7: 70, 8: 60, 9: 100, 10: 70, 11: 70, 12: 80, 13: 70 }).done()] };
}

function monthlyReport() {
  const s = new S('월간 리포트', { tab: MK }).title('월간 마케팅 성과 리포트', '이번 달 · 지난달 실적을 입력하면 증감과 요약이 자동으로 작성됩니다.', MK, 8);
  s.row(2, 1, ['보고 월', '2026-09', '', '작성자', '마케팅팀'], (j) => (j === 1 || j === 4 ? inputSt : { bold: true }));
  s.row(4, 1, ['지표', '지난달', '이번 달', '증감', '증감률', '코멘트'], head(MK));
  const K = [['광고비', 52000000, 58400000, WON], ['노출수', 18200000, 20400000, NUM], ['클릭수', 312000, 356000, NUM], ['전환수', 7800, 9150, NUM], ['매출', 248000000, 301000000, WON]];
  K.forEach(([k, a, b, f], i) => { const R1 = 6 + i; s.row(5 + i, 1, [k, a, b, `=D${R1}-C${R1}`, `=IFERROR(D${R1}/C${R1}-1,0)`, ''], (j) => (j === 1 || j === 2 ? { ...f, ...inputSt } : j === 3 ? f : j === 4 ? PCT : { bold: true })); });
  const D = [['CTR', '=C8/C7', '=D8/D7', PCT2], ['CPC', '=C6/C8', '=D6/D8', WON], ['CVR', '=C9/C8', '=D9/D8', PCT2], ['CPA', '=C6/C9', '=D6/D9', WON], ['ROAS', '=C10/C6', '=D10/D6', PCT]];
  D.forEach(([k, a, b, f], i) => { const R1 = 11 + i; s.row(10 + i, 1, [k, a, b, `=D${R1}-C${R1}`, `=IFERROR(D${R1}/C${R1}-1,0)`, ''], (j) => (j >= 1 && j <= 3 ? f : j === 4 ? PCT : { bold: true, color: MK })); });
  s.cf({ r1: 5, c1: 5, r2: 14, c2: 5, type: 'icons', icons: '3Arrows' });
  s.v(16, 1, '요약', { bold: true, size: 12 });
  s.v(17, 1, '="이번 달 광고비 "&TEXT(D6,"#,##0")&"원으로 매출 "&TEXT(D10,"#,##0")&"원, ROAS "&TEXT(D15,"0%")&"를 기록했습니다. 지난달 대비 매출은 "&TEXT(F10,"+0.0%;-0.0%")&", CPA는 "&TEXT(F14,"+0.0%;-0.0%")&" 변했습니다."', { wrap: true, valign: 'top' }).m(17, 1, 19, 6).h({ 17: 24, 18: 24, 19: 24 });
  s.chart({ type: 'column', title: '지난달 vs 이번 달 (광고비 · 매출)', range: { r1: 4, c1: 1, r2: 9, c2: 3 }, series: [{ name: { text: '지난달' }, cat: { r1: 5, c1: 1, r2: 5, c2: 1 }, val: { r1: 5, c1: 2, r2: 5, c2: 2 } }], x: 720, y: 60, w: 420, h: 260, palette: 'modern' });
  s.d.charts[0] = { ...s.d.charts[0], series: [{ name: { text: '지난달' }, cat: { r1: 5, c1: 1, r2: 9, c2: 1 }, val: { r1: 5, c1: 2, r2: 9, c2: 2 } }, { name: { text: '이번 달' }, cat: { r1: 5, c1: 1, r2: 9, c2: 1 }, val: { r1: 5, c1: 3, r2: 9, c2: 3 } }], range: undefined };
  return { sheets: [s.w({ 0: 16, 1: 100, 2: 130, 3: 130, 4: 120, 5: 80, 6: 260 }).done()] };
}

function seoRankTracker() {
  const s = new S('SEO 순위', { tab: '#0f766e' }).title('SEO 키워드 순위 트래커', '주차별 순위를 입력하면 변동 · 스파크라인 · 최고 순위가 표시됩니다 (숫자가 작을수록 좋음).', '#0f766e', 12);
  const weeks = Array.from({ length: 8 }, (_, i) => `W${i + 1}`);
  s.row(3, 1, ['키워드', 'URL', ...weeks, '변동', '최고', '추세'], head('#0f766e'));
  const R = rng(71);
  const K = ['유아 영양제', '어린이 오메가3', '아기 DHA 추천', '키 성장 영양제', '어린이 비타민 추천', '유아 유산균'];
  K.forEach((k, i) => {
    let rank = Math.round(5 + R() * 30);
    const ranks = weeks.map(() => { rank = Math.max(1, rank + Math.round((R() - 0.6) * 6)); return rank; });
    const R1 = 5 + i;
    s.row(4 + i, 1, [k, `/blog/${i + 1}`, ...ranks, `=D${R1}-K${R1}`, `=MIN(D${R1}:K${R1})`, ''], (j) => (j >= 2 && j <= 9 ? { align: 'center', ...inputSt } : j === 10 ? { numFmt: 'custom', code: '"▲"0;"▼"0;"-"' } : {}));
  });
  s.cf({ r1: 4, c1: 11, r2: 9, c2: 11, type: 'formula', formula: '=L5>0', style: { color: '#16a34a', bold: true } });
  s.cf({ r1: 4, c1: 11, r2: 9, c2: 11, type: 'formula', formula: '=L5<0', style: { color: '#dc2626', bold: true } });
  s.cf({ r1: 4, c1: 3, r2: 9, c2: 10, type: 'scale', colors: ['#34d399', '#fde68a', '#fb7185'] });
  s.set('sparklines', [{ id: 'sp1', type: 'line', color: '#0f766e', markers: true, low: true, lowColor: '#16a34a', items: K.map((_, i) => ({ r: 4 + i, c: 13, ref: `'SEO 순위'!D${5 + i}:K${5 + i}` })) }]);
  return { sheets: [s.w({ 0: 16, 1: 150, 2: 100, 11: 60, 12: 50, 13: 120 }).done()] };
}

function creativeRequests() {
  const s = new S('소재 관리', { tab: '#e11d48' }).title('광고 소재 제작 요청 · 관리', '소재별 규격 · 문구 · 상태 · 검수를 관리하고 마감 임박 소재를 강조합니다.', '#e11d48', 11);
  s.table({ name: '소재', r1: 3, c1: 1, r2: 9, c2: 11, style: 'WixelTableMinimal5' });
  s.row(3, 1, ['소재 ID', '매체', '규격', '형식', '헤드라인', '본문', 'CTA', '요청일', '마감일', '상태', '담당']);
  [['CR-001', '메타', '1080x1350', '이미지', '가을 한정 30% 할인', '지금 바로 확인하세요', '구매하기', '2026-10-01', '2026-10-06', '검수'], ['CR-002', '메타', '1080x1920', '릴스', '엄마들이 선택한 1위', '후기 3만 개', '더 알아보기', '2026-10-01', '2026-10-08', '제작 중'], ['CR-003', '구글', '1200x628', '이미지', '공식몰 단독 혜택', '무료 배송', '쇼핑하기', '2026-10-02', '2026-10-05', '완료'], ['CR-004', '카카오', '1029x258', '비즈보드', '오늘만 특가', '', '바로가기', '2026-10-02', '2026-10-04', '완료'], ['CR-005', '틱톡', '1080x1920', '영상', 'UGC 챌린지', '#브레인챌린지', '참여하기', '2026-10-03', '2026-10-10', '요청'], ['CR-006', '네이버', '-', '확장소재', '첫 구매 쿠폰', '', '', '2026-10-03', '2026-10-07', '요청']].forEach((row, i) => s.row(4 + i, 1, [...row, ['김', '이', '박'][i % 3] + '디자이너'], (j) => (j === 7 || j === 8 ? DATE : {})));
  s.dv({ r1: 4, c1: 10, r2: 200, c2: 10, type: 'list', f1: list(['요청', '제작 중', '검수', '수정', '완료', '보류']) });
  s.cf({ r1: 4, c1: 9, r2: 200, c2: 9, type: 'formula', formula: '=AND($K5<>"완료",J5-TODAY()<=2,J5<>"")', style: { fill: '#fee2e2', color: '#b91c1c', bold: true } });
  s.cf({ r1: 4, c1: 10, r2: 200, c2: 10, type: 'formula', formula: '=K5="완료"', style: { color: '#16a34a' } });
  return { sheets: [s.w({ 0: 16, 1: 70, 2: 60, 3: 90, 4: 70, 5: 160, 6: 140, 7: 80, 8: 90, 9: 90, 10: 70, 11: 80 }).done()] };
}

function dailyReport() {
  const raw = perfRows(60, 81, '2026-10-01');
  const s = new S('일일 리포트', { tab: MK }).title('일일 광고 성과 리포트', '날짜를 고르면 그날의 채널별 성과와 전일 대비 증감이 계산됩니다.', MK, 11);
  s.v(2, 1, '기준일', { bold: true }).v(2, 2, raw.at(-1)[0], { ...DATE, ...inputSt });
  s.row(4, 1, ['채널', '광고비', '클릭', '전환', '매출', 'ROAS', '전일 광고비', '전일 ROAS', 'ROAS 증감'], head(MK));
  CH.forEach((ch, i) => {
    const R1 = 6 + i;
    const sum = (col2, day) => `SUMIFS(일자료[${col2}],일자료[날짜],${day},일자료[채널],$B${R1})`;
    s.row(5 + i, 1, [ch, `=${sum('비용', '$C$3')}`, `=${sum('클릭수', '$C$3')}`, `=${sum('전환수', '$C$3')}`, `=${sum('매출', '$C$3')}`, `=IFERROR(F${R1}/C${R1},0)`, `=${sum('비용', '$C$3-1')}`, `=IFERROR(${sum('매출', '$C$3-1')}/H${R1},0)`, `=G${R1}-I${R1}`], (j) => (j === 1 || j === 4 || j === 6 ? WON : j === 5 || j === 7 || j === 8 ? PCT : j === 0 ? { bold: true } : NUM));
  });
  s.row(11, 1, ['합계', '=SUM(C6:C11)', '=SUM(D6:D11)', '=SUM(E6:E11)', '=SUM(F6:F11)', '=IFERROR(F12/C12,0)', '=SUM(H6:H11)'], (j) => ({ bold: true, fill: '#eef2ff', ...(j === 1 || j === 4 || j === 6 ? WON : j === 5 ? PCT : j === 0 ? {} : NUM) }));
  s.cf({ r1: 5, c1: 9, r2: 10, c2: 9, type: 'icons', icons: '3Arrows' });
  const d = new S('일자료', { grid: true, tab: '#94a3b8' });
  d.row(0, 0, ['날짜', '채널', '캠페인', '노출수', '클릭수', '비용', '전환수', '매출']);
  raw.forEach((r, i) => d.row(1 + i, 0, r, (j) => (j === 0 ? DATE : j >= 3 ? NUM : undefined)));
  d.table({ name: '일자료', r1: 0, c1: 0, r2: raw.length, c2: 7, style: 'WixelTableMinimal1' }).freeze(1);
  s.dv({ r1: 2, c1: 2, r2: 2, c2: 2, type: 'date', op: 'between', f1: '2026-01-01', f2: '2030-12-31' });
  return { sheets: [s.w({ 0: 16, 1: 100, 2: 110, 3: 80, 4: 70, 5: 110, 6: 70, 7: 110, 8: 80, 9: 80 }).done(), d.w({ 0: 90, 1: 80, 2: 150 }).done()] };
}

// ───────────── 네이버 검색광고 기본 리포트 (주간) ─────────────
const NRAW_IN = ['캠페인', '광고그룹', '키워드', 'date', '캠페인유형', '소재', '소재 유형', 'PC/모바일 매체', '검색/콘텐츠 매체', '노출수', '클릭수', '총비용', '평균노출순위', '구매완료 전환수', '구매완료 전환매출액(원)', '구매완료 광고수익률(%)'];
const NRAW_CALC = [
  ['년', '=YEAR([@date])'], ['월', '=MONTH([@date])'], ['년월', '=[@년]&"년 "&[@월]&"월"'],
  ['주차', '=TEXT([@date]-WEEKDAY([@date],2)+1,"yyyy/mm/dd")&" ~ "&TEXT([@date]-WEEKDAY([@date],2)+7,"yyyy/mm/dd")'],
  ['요일', '=TEXT([@date],"aaa")'], ['평일주말', '=IF(WEEKDAY([@date],2)>5,"주말","평일")'], ['가중치순위', '=[@평균노출순위]*[@노출수]'],
];
const NKRAW_IN = ['캠페인', '광고그룹', 'date', '캠페인유형', 'PC/모바일 매체', '검색어', '검색/콘텐츠 매체', '노출수', '클릭수', '총비용', '평균노출순위', '구매완료 전환수', '구매완료 전환매출액(원)', '구매완료 광고수익률(%)'];

function naverRows(seed = 91, days = 42) {
  const R = rng(seed);
  const camps = [['SA_브랜드_MO', '파워링크', '모바일'], ['SA_브랜드_PC', '파워링크', 'PC'], ['쇼핑검색_MO', '쇼핑검색', '모바일'], ['쇼핑검색_PC', '쇼핑검색', 'PC'], ['BS_브랜드검색_MO', '브랜드검색/신제품검색', '모바일']];
  const groups = ['A.브랜드', 'B.일반', 'C.경쟁', 'D.시즌'];
  const kws = ['브랜드명', '브랜드명 후기', '유아 영양제', '어린이 오메가3', '아기 DHA', '키 성장', '두뇌 영양제', '-'];
  const out = [];
  const kout = [];
  for (let d = 0; d < days; d++) {
    const date = addDays('2026-08-17', d);
    for (const [cp, type, dev] of camps) {
      for (let g = 0; g < (type === '브랜드검색/신제품검색' ? 1 : 2); g++) {
        const grp = groups[(g + cp.length) % groups.length];
        const kw = type === '파워링크' ? kws[Math.floor(R() * (kws.length - 1))] : '-';
        const imp = Math.round(30 + R() * (dev === 'PC' ? 300 : 900));
        const clk = Math.round(imp * (0.01 + R() * 0.06));
        const cost = type === '브랜드검색/신제품검색' ? 29333 : Math.round(clk * (150 + R() * 700));
        const conv = Math.round(clk * R() * 0.12);
        const rev = conv * Math.round(40000 + R() * 60000);
        out.push([cp, grp, kw, date, type, `nad-${Math.floor(R() * 1e9)}`, `${type}-단일형`, dev, '검색', imp, clk, cost, Math.round((1 + R() * 4) * 10) / 10, conv, rev, cost ? Math.round((rev / cost) * 10000) / 100 : 0]);
        const st = ['유아 영양제 추천', '어린이 오메가3 순위', '아기 DHA 언제부터', '브랜드명 가격', '영양제 쿠폰'][Math.floor(R() * 5)];
        kout.push([cp, grp, date, type, dev, st, '검색', imp, clk, cost, Math.round((1 + R() * 4) * 10) / 10, conv, rev, cost ? Math.round((rev / cost) * 10000) / 100 : 0]);
      }
    }
  }
  return { out, kout };
}

function naverReport() {
  const { out, kout } = naverRows();
  const calcFields = [
    { name: 'CPC', formula: 'DIVIDE(총비용,클릭수)' }, { name: 'CTR', formula: 'DIVIDE(클릭수,노출수)' }, { name: 'CPM', formula: 'DIVIDE(총비용,노출수)*1000' },
    { name: 'CVR', formula: "DIVIDE('구매완료 전환수',클릭수)" }, { name: 'CPA', formula: "DIVIDE(총비용,'구매완료 전환수')" }, { name: 'AOV', formula: "DIVIDE('구매완료 전환매출액(원)','구매완료 전환수')" },
    { name: 'ROAS', formula: "DIVIDE('구매완료 전환매출액(원)',총비용)" }, { name: '가중치평균순위', formula: 'DIVIDE(가중치순위,노출수)' },
  ];
  const v = (field, numFmt = NUM) => ({ field, agg: 'sum', name: field, numFmt });
  const kpiValues = [v('총비용', WON), v('노출수'), v('클릭수'), v('구매완료 전환수'), v('구매완료 전환매출액(원)', WON), v('가중치평균순위', { numFmt: 'number', decimals: 2 }), v('CPC', WON), v('CTR', PCT2), v('CPM', WON), v('CVR', PCT2), v('CPA', WON), v('AOV', WON), v('ROAS', PCT)];
  const base = { table: 'nraw', source: 'nraw', calcFields, layout: 'tabular', autoRefresh: true, autofit: false, errorCaption: '', style: 'WixelPivotModern1', needsRender: true };
  const D = '대시보드';
  const pvWeek = { ...base, name: '주차비교', rows: ['주차'], cols: [], values: kpiValues, top: 14, left: 1, sort: { 주차: { dir: 'desc' } }, top10: undefined };
  const pvCamp = { ...base, name: '캠페인', rows: ['캠페인'], cols: [], values: kpiValues, top: 34, left: 1 };
  const pvGrp = { ...base, name: '광고그룹', rows: ['광고그룹'], cols: [], values: kpiValues, top: 56, left: 1 };
  const pvDev = { ...base, name: '기기', rows: ['PC/모바일 매체'], cols: [], values: kpiValues, top: 78, left: 1 };
  const pvDay = { ...base, name: '요일', rows: ['요일'], cols: [], values: kpiValues, top: 90, left: 1, order: { 요일: ['월', '화', '수', '목', '금', '토', '일'] } };
  const pvDate = { ...base, name: '일자', rows: ['date'], cols: [], values: [v('총비용', WON), v('클릭수'), v('구매완료 전환수'), v('ROAS', PCT)], top: 105, left: 1 };
  const all = [pvWeek, pvCamp, pvGrp, pvDev, pvDay, pvDate];
  const pivots = all.map((p) => ({ sheet: D, name: p.name }));
  const s = new S(D, { tab: '#1e3a8a' });
  s.v(0, 1, '네이버 검색광고 주간 리포트', { bold: true, size: 18, color: '#1e3a8a' }).m(0, 1, 0, 8).h({ 0: 36 });
  const info = [['광고주', '광고주명을 입력하세요'], ['에이전시', '에이전시명을 입력하세요'], ['기간', '=TEXT(MIN(nraw[date]),"yyyy/mm/dd")&" ~ "&TEXT(MAX(nraw[date]),"yyyy/mm/dd")'], ['매체', '네이버']];
  info.forEach(([l, val], i) => s.row(2 + i, 1, [l, val], (j) => (j ? { ...(i === 2 ? {} : inputSt), align: 'center', ...box('#94a3b8') } : { bold: true, color: '#ffffff', fill: '#1e3a8a', align: 'center', ...box('#1e3a8a') })));
  info.forEach((_, i) => s.m(2 + i, 2, 2 + i, 4));
  s.v(7, 1, '사용 방법: ① [nraw] 시트 표에 네이버 광고 보고서(키워드 · 일별) 원본을 붙여넣기 ② [nkraw] 시트에 검색어 보고서 원본 붙여넣기 → 피벗 · 차트 · 슬라이서가 자동으로 새로 고쳐집니다.', { color: MUTED, size: 9 }).m(7, 1, 7, 14);
  s.v(12, 1, '■ 주차별 성과', { bold: true, size: 12, color: '#1e3a8a' });
  s.v(32, 1, '■ 캠페인별 성과', { bold: true, size: 12, color: '#1e3a8a' });
  s.v(54, 1, '■ 광고그룹별 성과', { bold: true, size: 12, color: '#1e3a8a' });
  s.v(76, 1, '■ 기기별 성과', { bold: true, size: 12, color: '#1e3a8a' });
  s.v(88, 1, '■ 요일별 성과', { bold: true, size: 12, color: '#1e3a8a' });
  s.v(103, 1, '■ 일자별 성과', { bold: true, size: 12, color: '#1e3a8a' });
  s.set('pivot', pvWeek).set('pivotsExtra', all.slice(1));
  const slicer = (id, field, x, y, w, h, columns = 1) => ({ id, caption: field, source: { kind: 'pivot', field, pivots }, columns, style: 'WixelSlicerSolid1', multi: true, x, y, w, h, z: 20, fontSize: 9 });
  s.set('slicers', [slicer('ns1', '캠페인유형', 560, 36, 170, 150), slicer('ns2', '년월', 740, 36, 130, 150), slicer('ns3', '주차', 880, 36, 420, 150, 2), slicer('ns4', '캠페인', 1310, 36, 300, 150, 2)]);
  const chartAt = (name, y) => ({ type: 'combo', title: `${name}별 광고비 · CVR`, pivot: { sheet: D, name }, x: 1250, y, w: 560, h: 280, palette: 'modern', rounded: true, border: LINE, titleBold: true, fieldButtons: false, seriesFmt: [{ type: 'column' }, { type: 'line', axis: 1, numFmt: '0.0%' }] });
  const trim = (p) => ({ ...p, name: `${p.name}그래프`, values: [v('총비용', WON), v('CVR', PCT2)], left: 30, needsRender: true });
  const graphPivots = [trim(pvCamp), trim(pvGrp), trim(pvDev), trim(pvDay), trim(pvDate)];
  s.set('pivotsExtra', [...all.slice(1), ...graphPivots.map((p) => ({ ...p, left: 40 }))]);
  graphPivots.forEach((p, i) => s.chart({ ...chartAt(p.name, [640, 1080, 1520, 1760, 2060][i]), x: 1240 }));
  for (const sl of s.d.slicers) sl.source.pivots = [...pivots, ...graphPivots.map((p) => ({ sheet: D, name: p.name }))];
  s.w({ 0: 16, 1: 190, ...Object.fromEntries(Array.from({ length: 13 }, (_, i) => [2 + i, i === 4 ? 120 : 96])) });
  // 키워드 · 검색어 시트
  const kwSheet = (name, field, table, tab, color) => {
    const k = new S(name, { tab: color });
    k.v(0, 1, `네이버 검색광고 ${name} 성과`, { bold: true, size: 16, color: '#1e3a8a' }).h({ 0: 32 });
    const pv = { ...base, table, source: table, name: `${name}피벗`, rows: [field], cols: [], values: kpiValues, top: 16, left: 1, sort: { [field]: { dir: 'desc', by: 0 } } };
    k.set('pivot', pv);
    const ps = [{ sheet: name, name: pv.name }];
    const sl = (id, f, x, y, w, h, columns = 1) => ({ id, caption: f, source: { kind: 'pivot', field: f, pivots: ps }, columns, style: 'WixelSlicerSoft1', multi: true, x, y, w, h, z: 20, fontSize: 9 });
    k.set('slicers', [sl(`${tab}1`, '캠페인유형', 30, 44, 170, 230), sl(`${tab}2`, '캠페인', 210, 44, 240, 230), sl(`${tab}3`, 'PC/모바일 매체', 460, 44, 150, 110), sl(`${tab}4`, '평일주말', 460, 164, 150, 110), sl(`${tab}5`, '주차', 620, 44, 320, 230, 2), sl(`${tab}6`, '요일', 950, 44, 120, 230)]);
    k.cf({ r1: 17, c1: 14, r2: 5000, c2: 14, type: 'scale', colors: ['#fb7185', '#fde68a', '#34d399'] });
    return k.w({ 0: 16, 1: 200, ...Object.fromEntries(Array.from({ length: 13 }, (_, i) => [2 + i, i === 4 ? 120 : 96])) }).done();
  };
  // 원본 시트 (표 + 계산 열)
  const rawSheet = (name, cols, rows, tab) => {
    const r = new S(name, { grid: true, tab });
    const headers = [...cols, ...(name === 'nraw' ? NRAW_CALC : NRAW_CALC).map(([h]) => h)];
    r.row(0, 0, headers, (j) => ({ bold: true, ...(j >= cols.length ? { fill: '#e0e7ff' } : {}) }));
    rows.forEach((row, i) => {
      row.forEach((val, j) => r.v(1 + i, j, val, cols[j] === 'date' ? DATE : typeof val === 'number' ? NUM : undefined));
      NRAW_CALC.forEach(([, f], j) => r.v(1 + i, cols.length + j, f.replace(/\[@/g, `${name}[@`), j === 6 ? NUM : undefined));
    });
    r.table({ name, r1: 0, c1: 0, r2: rows.length, c2: headers.length - 1, style: 'TableStyleMedium2', banded: false });
    return r.freeze(1).w({ 0: 150, 1: 90, 2: 110, [cols.indexOf('date')]: 90 }).done();
  };
  const guide = new S('사용 방법', { tab: '#16a34a' }).title('네이버 검색광고 리포트 템플릿 사용 방법', null, '#16a34a', 10);
  [
    '1. [대시보드] C3 · C4 칸에 광고주명 · 에이전시명을 입력합니다. 기간은 원본 날짜에서 자동으로 계산됩니다.',
    '2. 네이버 광고 시스템 → 보고서 → 다차원 보고서에서 [캠페인 · 광고그룹 · 키워드 · 일별 · 캠페인유형 · 소재 · PC/모바일 · 검색/콘텐츠] 기준으로 내려받습니다.',
    `   필요한 열 (순서대로): ${NRAW_IN.join(' · ')}`,
    '3. [nraw] 시트의 A2 칸을 선택하고 원본을 붙여넣으면 표가 자동으로 늘어나고 년 · 월 · 주차 · 요일 · 평일주말 · 가중치순위 열이 계산됩니다.',
    `4. 검색어 보고서는 [nkraw] 시트 A2 에 붙여넣습니다. 필요한 열: ${NKRAW_IN.join(' · ')}`,
    '5. 피벗 테이블 · 차트 · 슬라이서는 원본이 바뀌면 자동으로 새로 고쳐집니다 (피벗 테이블 옵션 → 데이터 → 자동 새로 고침).',
    '6. 슬라이서로 캠페인유형 · 년월 · 주차 · 캠페인을 고르면 대시보드 전체가 함께 걸러집니다.',
    '7. 예시 데이터는 지우고 쓰세요: [nraw] · [nkraw] 표의 데이터 행을 선택 → 행 삭제.',
    '※ CPC · CTR · CPM · CVR · CPA · AOV · ROAS 는 피벗 계산 필드로 합계 기준(가중 평균)으로 정확하게 계산됩니다.',
  ].forEach((t, i) => guide.v(2 + i, 1, t, { color: i === 8 ? '#b45309' : INK }));
  guide.w({ 0: 16, 1: 900 });
  const kwD = kwSheet('키워드', '키워드', 'nraw', 'k', '#2563eb');
  const stD = kwSheet('검색어', '검색어', 'nkraw', 's', '#0891b2');
  return { sheets: [s.done(), kwD, stD, rawSheet('nraw', NRAW_IN, out, '#64748b'), rawSheet('nkraw', NKRAW_IN, kout, '#64748b'), guide.done()] };
}

// ───────────── 목록 ─────────────
export const TEMPLATE_CATS = ['기본', '퍼포먼스 마케팅'];
export const TEMPLATES = [
  { id: 'naver-sa', cat: '퍼포먼스 마케팅', name: '네이버 검색광고 주간 리포트', desc: '원본 붙여넣기 → 피벗 · 차트 · 슬라이서 자동 (대시보드 · 키워드 · 검색어)', color: '#03c75a', featured: true, build: naverReport },
  { id: 'perf-dash', cat: '퍼포먼스 마케팅', name: '퍼포먼스 대시보드', desc: 'KPI 카드 · 채널 피벗 · 일별 콤보 차트 · 슬라이서', color: MK, featured: true, build: perfDashboard },
  { id: 'mk-cal', cat: '퍼포먼스 마케팅', name: '퍼포먼스 마케팅 일정', desc: '캠페인 · 채널 · 기간 · 예산 달력', color: MK, build: marketingCalendar },
  { id: 'content-cal', cat: '퍼포먼스 마케팅', name: '콘텐츠 일정', desc: '채널별 게시 캘린더 · 상태 · 성과', color: '#db2777', build: contentCalendar },
  { id: 'mk-gantt', cat: '퍼포먼스 마케팅', name: '마케팅 간트 차트', desc: '캠페인 준비 단계별 일정 막대', color: '#0ea5e9', build: () => gantt('마케팅 간트 차트', '#0ea5e9', [['기획', '캠페인 브리프', '김마케터', 0, 3, 1], ['기획', '타깃 · 예산 확정', '김마케터', 2, 3, 1], ['제작', '소재 기획', '이디자이너', 4, 4, 0.7], ['제작', '소재 제작', '이디자이너', 7, 6, 0.3], ['세팅', '픽셀 · 전환 설정', '박개발', 8, 3, 0.5], ['세팅', '캠페인 세팅', '김마케터', 12, 2, 0], ['운영', '런칭 · 최적화', '김마케터', 14, 14, 0], ['보고', '결과 리포트', '김마케터', 28, 3, 0]], '캠페인 준비부터 보고까지. 시작일 · 기간 · 진행률을 입력하세요.') },
  { id: 'todo', cat: '퍼포먼스 마케팅', name: '투두 리스트', desc: '완료 체크 · 우선순위 · 마감 강조 · 진행률', color: '#16a34a', build: todoList },
  { id: 'promo', cat: '퍼포먼스 마케팅', name: '프로모션 일정표', desc: '기간 · 할인 · 목표 대비 실적 · 상태', color: '#ea580c', build: promoCalendar },
  { id: 'creative', cat: '퍼포먼스 마케팅', name: '소재 효율 분석', desc: 'CTR · CVR · CPA · ROAS · 판정 · 거품형 차트', color: '#9333ea', build: creativeEfficiency },
  { id: 'searchterm', cat: '퍼포먼스 마케팅', name: '검색어 효율 분석', desc: '제외 키워드 후보 · 절감액 자동 계산', color: '#0891b2', build: searchTermEfficiency },
  { id: 'pacing', cat: '퍼포먼스 마케팅', name: '예산 페이싱', desc: '일별 소진 · 이상 곡선 · 월말 예상', color: '#0d9488', build: budgetPacing },
  { id: 'mix', cat: '퍼포먼스 마케팅', name: '미디어 믹스 플래너', desc: '예산 배분 → 예상 클릭 · 전환 · ROAS', color: '#7c3aed', build: mediaMix },
  { id: 'ab', cat: '퍼포먼스 마케팅', name: 'A/B 테스트 계산기', desc: '전환율 · 상승률 · p 값 · 필요 표본', color: '#2563eb', build: abTest },
  { id: 'utm', cat: '퍼포먼스 마케팅', name: 'UTM 링크 빌더', desc: 'UTM 추적 링크 자동 생성', color: '#475569', build: utmBuilder },
  { id: 'kpi', cat: '퍼포먼스 마케팅', name: 'KPI 목표 트래커', desc: '달성률 · 페이스 · 신호등', color: '#b45309', build: kpiTracker },
  { id: 'funnel', cat: '퍼포먼스 마케팅', name: '전환 퍼널 분석', desc: '단계 전환율 · 이탈률 · 깔때기 차트', color: '#0284c7', build: funnelAnalysis },
  { id: 'cohort', cat: '퍼포먼스 마케팅', name: '코호트 리텐션', desc: '잔존율 히트맵 · 리텐션 곡선', color: '#be123c', build: cohortRetention },
  { id: 'kwplan', cat: '퍼포먼스 마케팅', name: '키워드 플래너', desc: '검색량 · CPC → 예상 비용 · 전환 · 우선순위', color: '#16a34a', build: keywordPlanner },
  { id: 'mediaplan', cat: '퍼포먼스 마케팅', name: '미디어 플랜 (집행 계획서)', desc: '매체 · 단가 · 수수료 · VAT 총액', color: '#1e293b', build: mediaPlan },
  { id: 'roas-sim', cat: '퍼포먼스 마케팅', name: 'ROAS · 손익 시뮬레이터', desc: '손익분기 ROAS · CPC × CVR 민감도 표', color: '#059669', build: roasSimulator },
  { id: 'competitor', cat: '퍼포먼스 마케팅', name: '경쟁사 모니터링', desc: '경쟁사 소재 · 프로모션 · 가격 기록', color: '#64748b', build: competitorTracker },
  { id: 'influencer', cat: '퍼포먼스 마케팅', name: '인플루언서 캠페인', desc: 'CPV · CPE · CPA · ROAS', color: '#c026d3', build: influencerTracker },
  { id: 'monthly', cat: '퍼포먼스 마케팅', name: '월간 마케팅 리포트', desc: '전월 대비 증감 · 자동 요약 문장', color: MK, build: monthlyReport },
  { id: 'daily', cat: '퍼포먼스 마케팅', name: '일일 광고 리포트', desc: '기준일 채널별 성과 · 전일 대비', color: MK, build: dailyReport },
  { id: 'seo', cat: '퍼포먼스 마케팅', name: 'SEO 순위 트래커', desc: '주차별 순위 · 변동 · 스파크라인', color: '#0f766e', build: seoRankTracker },
  { id: 'creative-req', cat: '퍼포먼스 마케팅', name: '광고 소재 제작 관리', desc: '규격 · 문구 · 상태 · 마감 임박 강조', color: '#e11d48', build: creativeRequests },
  { id: 'calendar', cat: '기본', name: '만년 달력', desc: '연 · 월을 고르면 바뀌는 달력', color: '#1e3a8a', build: perpetualCalendar },
  { id: 'personal', cat: '기본', name: '월별 개인 예산', desc: '수입 · 지출 · 차이 · 사용률 · 차트', color: '#047857', build: personalBudget },
  { id: 'company', cat: '기본', name: '월간 회사 예산', desc: '부서별 예산 · 실적 · 달성률', color: '#1e40af', build: companyBudget },
  { id: 'balance', cat: '기본', name: '대차대조표', desc: '자산 · 부채 · 자본 · 검증', color: '#334155', build: balanceSheet },
  { id: 'household', cat: '기본', name: '가구 월별 예산', desc: '12개월 항목별 지출 · 비중 · 추세', color: '#7c3aed', build: householdBudget },
  { id: 'gantt', cat: '기본', name: '간트 프로젝트 플래너', desc: '작업 · 기간 · 진행률 막대 · 오늘 선', color: '#0369a1', build: () => gantt() },
  { id: 'vacation', cat: '기본', name: '휴가 일정', desc: '직원별 휴가 · 근무일 수 · 달력', color: '#be185d', build: vacationSchedule },
  { id: 'invoice', cat: '기본', name: '송장 (거래명세서)', desc: '품목 · 공급가액 · 부가세 · 합계', color: '#111827', build: invoice },
  { id: 'home-inv', cat: '기본', name: '집 재고', desc: '방별 물품 · 현재 가치 · 보증 만료', color: '#9a3412', build: homeInventory },
  { id: 'inventory', cat: '기본', name: '재고 목록', desc: '재고 가치 · 재주문 표시', color: '#065f46', build: inventoryList },
  { id: 'loan-calc', cat: '기본', name: '대출 계산기', desc: '월 상환액 · 총이자 · 금리별 비교', color: '#0f766e', build: loanCalculator },
  { id: 'loan-sched', cat: '기본', name: '대출 상환 일정', desc: '원리금 균등 회차별 원금 · 이자', color: '#1d4ed8', build: loanSchedule },
  { id: 'student', cat: '기본', name: '학생 일정', desc: '주간 시간표 · 과목 색 · 과제', color: '#4338ca', build: studentSchedule },
  { id: 'attendance', cat: '기본', name: '주간 출석 보고서', desc: '출석 · 지각 · 결석 · 출석률', color: '#15803d', build: weeklyAttendance },
];
