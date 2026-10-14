// 차트 편집 UI (PowerPoint 차트 디자인 · 서식): 차트 종류 변경 대화상자, 표 형식 데이터 편집, 차트 요소 추가 메뉴,
// 빠른 레이아웃 · 스타일 · 색 변경 갤러리, 오른쪽 [차트 서식] 창 (요소별)
import { S, change, emit, selObjects } from './state.js';
import { el, openDialog, alertDialog, toast, closeMenus } from './ui.js';
import { colorButton } from './colorpick.js';
import {
  CHART_KINDS, CHART_LABEL, CHART_SUBTYPES, CHART_PALETTES, CHART_LAYOUTS, CHART_STYLES, TREND_LABEL,
  chartSvg, paletteColor, seriesType, applyChartLayout, applyChartStyle, switchRowCol, gridToChart, chartCsv, VIEW3D_DEFAULT,
} from './chart.js';

export const selChart = () => { const o = selObjects()[0]; return o?.type === 'chart' ? o : null; };
const clone = (x) => JSON.parse(JSON.stringify(x));
const setChart = (o, fn) => change(() => { const next = fn(clone(o.chart)); if (next) o.chart = next; }, { key: `chart-${o.id}` });

// ───────────── 기본 데이터 (종류마다) ─────────────
const SAMPLES = {
  waterfall: { cats: ['시작', '1분기', '2분기', '3분기', '4분기', '합계'], series: [{ name: '손익', vals: [120, 45, -30, 60, -20, 0] }] },
  funnel: { cats: ['방문', '상품 조회', '장바구니', '주문', '재구매'], series: [{ name: '고객 수', vals: [10000, 6200, 2400, 1100, 420] }] },
  histogram: { cats: Array.from({ length: 24 }, (_, i) => String(i + 1)), series: [{ name: '점수', vals: [52, 58, 61, 63, 65, 66, 68, 70, 71, 72, 73, 75, 76, 77, 78, 80, 81, 83, 85, 86, 88, 91, 94, 98] }] },
  box: { cats: ['1', '2', '3', '4', '5', '6', '7', '8'], series: [{ name: '1반', vals: [62, 70, 74, 78, 81, 85, 88, 97] }, { name: '2반', vals: [55, 64, 69, 72, 75, 79, 83, 90] }, { name: '3반', vals: [70, 73, 77, 80, 84, 86, 91, 99] }] },
  treemap: { cats: ['디지털/스마트폰', '디지털/노트북', '디지털/태블릿', '가전/TV', '가전/냉장고', '생활/주방', '생활/욕실', '식품'], series: [{ name: '매출', vals: [420, 310, 120, 260, 180, 90, 60, 150] }] },
  sunburst: { cats: ['디지털/스마트폰', '디지털/노트북', '가전/TV', '가전/냉장고', '생활/주방', '생활/욕실'], series: [{ name: '매출', vals: [420, 310, 260, 180, 90, 60] }] },
};
export function sampleChart(kind, sub) {
  if (SAMPLES[kind]) return withSubtype({ kind, title: '차트 제목', legend: kind === 'box' || kind === 'treemap' ? 'b' : null, labels: kind !== 'box', ...JSON.parse(JSON.stringify(SAMPLES[kind])) }, kind, sub);
  const base = { kind, title: '차트 제목', legend: 'b', labels: false, cats: ['항목 1', '항목 2', '항목 3', '항목 4'], series: [{ name: '계열 1', vals: [4.3, 2.5, 3.5, 4.5] }, { name: '계열 2', vals: [2.4, 4.4, 1.8, 2.8] }, { name: '계열 3', vals: [2, 2, 3, 5] }] };
  return withSubtype(base, kind, sub);
}
/** 종류 · 하위 형식 적용 (데이터는 되도록 유지) */
export function withSubtype(ch, kind, sub) {
  const c = { ...clone(ch), kind };
  delete c.stacked;
  delete c.variant;
  delete c.grouping;
  delete c.d3;
  sub ??= CHART_SUBTYPES[kind]?.[0]?.[0];
  // 3차원 하위 형식: '…3d' 는 같은 하위 형식 + d3, 'full3d' 는 계열을 깊이로 (3차원 세로 막대형)
  if (sub && /3d$/.test(sub)) {
    c.d3 = true;
    c.view3d ??= {};
    if (sub === 'full3d') { c.variant3d = 'full'; sub = 'clustered'; } else sub = sub.replace(/3d$/, '') || CHART_SUBTYPES[kind][0][0];
  } else delete c.view3d;
  if (kind === 'col' || kind === 'bar' || kind === 'area' || (kind === 'line' && ['stacked', 'percent'].includes(sub))) { if (sub && sub !== 'clustered' && sub !== 'standard') c.grouping = sub; }
  if (kind === 'line' && (sub === 'markers' || sub === 'smooth')) c.variant = sub;
  if (kind === 'scatter' && sub !== 'marker') c.variant = sub;
  if (kind === 'radar' && sub !== 'standard') c.variant = sub;
  if (kind === 'histogram' && sub === 'pareto') c.variant = 'pareto';
  if (kind === 'pie' && sub === 'explode') c.variant = 'explode';
  if (kind === 'pie' || kind === 'doughnut') { if (kind === 'pie') c.series = c.series.slice(0, 1); if (!c.labels) c.labels = true; c.legend = c.legend ?? 'r'; }
  if (kind === 'scatter' || kind === 'bubble') { if (!c.cats.every((v) => Number.isFinite(Number(v)) && String(v).trim() !== '')) c.cats = c.cats.map((_, i) => String(i + 1)); }
  if (kind === 'bubble') for (const s of c.series) s.sizes ??= s.vals.map((v, i) => (i % 3) + 1);
  if (kind === 'radar' && c.cats.length < 3) c.cats = [...c.cats, ...['항목 A', '항목 B', '항목 C'].slice(0, 3 - c.cats.length)];
  for (const s of c.series) { delete s.type; if (kind !== 'combo') delete s.axis; }
  if (kind === 'combo') {
    const last = c.series.length - 1;
    c.series.forEach((s, i) => { s.type = sub === 'areaCol' ? (i === 0 ? 'area' : 'col') : i === last && last > 0 ? 'line' : 'col'; if (sub === 'colLine2' && i === last && last > 0) s.axis = 'secondary'; });
    if (sub === 'areaCol') c.grouping = 'stacked';
  }
  return c;
}

// ───────────── 차트 종류 변경 / 삽입 대화상자 ─────────────
/** o 가 없으면 새 차트 삽입 (onInsert(chart)) */
export function chartTypeDialog(o, onInsert) {
  const theme = S.pres.theme;
  const cur = o?.chart;
  let kind = cur?.kind ?? 'col';
  let sub = cur ? currentSub(cur) : 'clustered';
  let combo = cur?.kind === 'combo' ? clone(cur.series.map((s, i) => ({ name: s.name, type: seriesType(cur, s, i), axis: s.axis === 'secondary' }))) : null;
  const left = el('div', { class: 'ct-kinds' });
  const subs = el('div', { class: 'ct-subs' });
  const preview = el('div', { class: 'ct-preview' });
  const comboBox = el('div', { class: 'ct-combo' });
  const build = () => {
    const base = cur ?? sampleChart(kind);
    let c = withSubtype(base, kind, sub);
    if (kind === 'combo' && combo) c.series.forEach((s, i) => { if (combo[i]) { s.type = combo[i].type; if (combo[i].axis) s.axis = 'secondary'; else delete s.axis; } });
    return c;
  };
  const draw = () => {
    left.innerHTML = '';
    for (const [k, l] of CHART_KINDS) left.append(el('button', { class: `ct-kind${k === kind ? ' on' : ''}`, type: 'button', onclick: () => { kind = k; sub = CHART_SUBTYPES[k][0][0]; combo = null; draw(); } }, l));
    subs.innerHTML = '';
    for (const [k, l] of CHART_SUBTYPES[kind]) {
      const c = withSubtype(cur ?? sampleChart(kind), kind, k);
      subs.append(el('button', { class: `ct-sub${k === sub ? ' on' : ''}`, type: 'button', title: l, onclick: () => { sub = k; combo = null; draw(); } }, el('div', { class: 'ct-thumb', html: chartSvg({ ...c, title: '', legend: null, labels: false, dataTable: false, axes: { x: { hide: true }, y: { hide: true } } }, 72, 50, theme) })));
    }
    const c = build();
    preview.innerHTML = `<b>${CHART_SUBTYPES[kind].find(([k]) => k === sub)?.[1] ?? CHART_LABEL[kind]}</b>${chartSvg(c, 420, 250, theme, { font: theme.fonts.minor })}`;
    comboBox.innerHTML = '';
    if (kind === 'combo') {
      combo ??= c.series.map((s, i) => ({ name: s.name, type: seriesType(c, s, i), axis: s.axis === 'secondary' }));
      comboBox.append(el('div', { class: 'ct-combo-head' }, el('span', {}, '계열 이름'), el('span', {}, '차트 종류'), el('span', {}, '보조 축')));
      combo.forEach((r) => {
        const t = el('select', {}, [['col', '묶은 세로 막대형'], ['line', '꺾은선형'], ['area', '영역형']].map(([v, l]) => el('option', { value: v, selected: v === r.type }, l)));
        t.addEventListener('change', () => { r.type = t.value; draw(); });
        const ax = el('input', { type: 'checkbox', checked: r.axis });
        ax.addEventListener('change', () => { r.axis = ax.checked; draw(); });
        comboBox.append(el('div', { class: 'ct-combo-row' }, el('span', {}, r.name), t, ax));
      });
    }
  };
  draw();
  openDialog({
    title: o ? '차트 종류 변경' : '차트 삽입', width: 760,
    body: el('div', { class: 'ct-wrap' }, left, el('div', { class: 'ct-main' }, subs, preview, comboBox)),
    buttons: [{ label: '확인', primary: true, action: () => {
      const c = build();
      if (o) change(() => { o.chart = c; });
      else onInsert?.(c);
      return true;
    } }, { label: '취소' }],
  });
}
function currentSub(c) {
  if (c.kind === 'combo') return c.grouping === 'stacked' ? 'areaCol' : c.series.some((s) => s.axis === 'secondary') ? 'colLine2' : 'colLine';
  if (c.variant) return c.variant;
  const g = c.grouping ?? (c.stacked ? 'stacked' : null);
  if (g) return g;
  return CHART_SUBTYPES[c.kind]?.[0]?.[0];
}

// ───────────── 데이터 편집 (표) ─────────────
export function chartDataDialog(o) {
  const c = o ?? selChart();
  if (!c) { toast('차트를 선택하세요'); return; }
  const ch = c.chart;
  const bub = ch.kind === 'bubble';
  let rows = [['', ...ch.series.flatMap((s) => (bub ? [s.name, `${s.name} 크기`] : [s.name]))], ...ch.cats.map((cat, i) => [cat, ...ch.series.flatMap((s) => (bub ? [s.vals[i] ?? '', s.sizes?.[i] ?? ''] : [s.vals[i] ?? '']))])];
  rows = rows.map((r) => r.map((v) => String(v ?? '')));
  const table = el('table', { class: 'cd-grid' });
  const wrap = el('div', { class: 'cd-wrap' }, table);
  let focus = [1, 1];
  const cols = () => Math.max(...rows.map((r) => r.length));
  const pad = () => { const n = cols(); for (const r of rows) while (r.length < n) r.push(''); };
  const draw = () => {
    pad();
    table.innerHTML = '';
    const n = cols();
    table.append(el('tr', {}, el('th', {}, ''), ...Array.from({ length: n }, (_, k) => el('th', {}, String.fromCharCode(65 + (k % 26))))));
    rows.forEach((r, ri) => {
      table.append(el('tr', {}, el('th', {}, String(ri + 1)), ...r.map((v, ci) => {
        const inp = el('input', { type: 'text', value: v, class: `${ri === 0 || ci === 0 ? 'hd' : ''}${ri > 0 && ci > 0 ? ' numc' : ''}`, dataset: { r: String(ri), c: String(ci) } });
        inp.addEventListener('input', () => { rows[ri][ci] = inp.value; });
        inp.addEventListener('focus', () => { focus = [ri, ci]; inp.select(); });
        inp.addEventListener('keydown', (e) => {
          e.stopPropagation();
          const mv = { ArrowDown: [1, 0], Enter: [1, 0], ArrowUp: [-1, 0], ArrowRight: inp.selectionStart === inp.value.length ? [0, 1] : null, ArrowLeft: inp.selectionStart === 0 ? [0, -1] : null }[e.key];
          if (!mv) return;
          e.preventDefault();
          const nr = Math.max(0, ri + mv[0]);
          const nc = Math.max(0, ci + mv[1]);
          if (nr >= rows.length) { rows.push(new Array(cols()).fill('')); draw(); }
          table.querySelector(`input[data-r="${nr}"][data-c="${Math.min(nc, cols() - 1)}"]`)?.focus();
        });
        // 엑셀 · WIXEL 에서 복사한 표 붙여 넣기 (이 칸부터)
        inp.addEventListener('paste', (e) => {
          const t = e.clipboardData?.getData('text/plain') ?? '';
          if (!/[\t\n]/.test(t)) return;
          e.preventDefault();
          const lines = t.replace(/\r/g, '').replace(/\n$/, '').split('\n').map((l) => l.split('\t'));
          lines.forEach((l, dr) => l.forEach((v, dc) => { while (rows.length <= ri + dr) rows.push([]); const row = rows[ri + dr]; while (row.length <= ci + dc) row.push(''); row[ci + dc] = v; }));
          draw();
        });
        return el('td', {}, inp);
      })));
    });
  };
  draw();
  const btn = (label, fn) => el('button', { class: 'btn small', type: 'button', onclick: () => { fn(); draw(); } }, label);
  const tools = el('div', { class: 'cd-tools' },
    btn('행 추가', () => rows.push(new Array(cols()).fill(''))),
    btn('열 추가', () => { for (const r of rows) r.push(''); rows[0][cols() - 1] = `계열 ${cols() - 1}`; }),
    btn('행 삭제', () => { if (rows.length > 2 && focus[0] > 0) rows.splice(focus[0], 1); }),
    btn('열 삭제', () => { if (cols() > 2 && focus[1] > 0) for (const r of rows) r.splice(focus[1], 1); }),
    btn('행/열 전환', () => { const n = cols(); rows = Array.from({ length: n }, (_, k) => rows.map((r) => r[k] ?? '')); }),
    el('span', { class: 'muted' }, '엑셀 · WIXEL 에서 복사한 표를 칸에 붙여 넣을 수 있습니다.'));
  openDialog({
    title: `차트 데이터 편집 — ${CHART_LABEL[ch.kind] ?? ''}`, width: 680,
    body: el('div', {}, tools, wrap, bub ? el('p', { class: 'muted' }, '거품형: 계열마다 값 열과 크기 열이 짝으로 있습니다. 첫 열(X 값)은 숫자입니다.') : null),
    buttons: [{ label: '확인', primary: true, action: () => {
      try {
        const next = gridToChart(ch, rows);
        change(() => { c.chart = next; });
      } catch (e) { alertDialog('차트', e.message); return false; }
      return true;
    } }, { label: '취소' }],
    onOpen: () => table.querySelector('input[data-r="1"][data-c="1"]')?.focus(),
  });
}

// ───────────── 차트 요소 추가 (+ 메뉴) ─────────────
export function chartElementsItems(o) {
  const ch = o.chart;
  const pie = ch.kind === 'pie' || ch.kind === 'doughnut';
  const set = (fn) => setChart(o, (c) => { fn(c); return c; });
  const ax = (c, k) => { c.axes ??= {}; c.axes[k] ??= {}; return c.axes[k]; };
  const lab = typeof ch.labels === 'object' ? ch.labels : ch.labels ? {} : null;
  const labPos = pie ? [['ctr', '가운데'], ['inEnd', '안쪽 끝에'], ['outEnd', '바깥쪽 끝에'], ['bestFit', '자동 맞춤']] : ch.kind === 'line' || ch.kind === 'scatter' ? [['ctr', '가운데'], ['l', '왼쪽'], ['r', '오른쪽'], ['t', '위쪽'], ['b', '아래쪽']] : [['ctr', '가운데'], ['inEnd', '안쪽 끝에'], ['inBase', '축에 가깝게'], ['outEnd', '바깥쪽 끝에']];
  const has2 = ch.series.some((s) => s.axis === 'secondary');
  const items = [];
  if (!pie) {
    items.push({ label: '축(X)', submenu: [
      { label: '기본 가로', checked: !ch.axes?.x?.hide, action: () => set((c) => { ax(c, 'x').hide = !c.axes.x.hide || undefined; }) },
      { label: '기본 세로', checked: !ch.axes?.y?.hide, action: () => set((c) => { ax(c, 'y').hide = !c.axes.y.hide || undefined; }) },
      ...(has2 ? [{ label: '보조 세로', checked: !ch.axes?.y2?.hide, action: () => set((c) => { ax(c, 'y2').hide = !c.axes.y2.hide || undefined; }) }] : []),
      { sep: true }, { label: '추가 축 옵션...', action: () => openChartPane(o, 'axis-y') },
    ] });
    items.push({ label: '축 제목(A)', submenu: [
      { label: '기본 가로', checked: !!ch.axes?.x?.title, action: () => set((c) => { const a = ax(c, 'x'); a.title = a.title ? undefined : '축 제목'; }) },
      { label: '기본 세로', checked: !!ch.axes?.y?.title, action: () => set((c) => { const a = ax(c, 'y'); a.title = a.title ? undefined : '축 제목'; }) },
      ...(has2 ? [{ label: '보조 세로', checked: !!ch.axes?.y2?.title, action: () => set((c) => { const a = ax(c, 'y2'); a.title = a.title ? undefined : '축 제목'; }) }] : []),
    ] });
  }
  items.push({ label: '차트 제목(C)', submenu: [
    { label: '없음', checked: !ch.title, action: () => set((c) => { c.title = ''; }) },
    { label: '차트 위', checked: !!ch.title && !ch.titleOverlay, action: () => set((c) => { c.title = c.title || '차트 제목'; delete c.titleOverlay; }) },
    { label: '가운데 맞춤 겹쳐서', checked: !!ch.title && !!ch.titleOverlay, action: () => set((c) => { c.title = c.title || '차트 제목'; c.titleOverlay = true; }) },
    { sep: true }, { label: '추가 제목 옵션...', action: () => openChartPane(o, 'title') },
  ] });
  items.push({ label: '데이터 레이블(D)', submenu: [
    { label: '없음', checked: !lab, action: () => set((c) => { c.labels = false; for (const s of c.series) delete s.labels; }) },
    ...labPos.map(([k, l]) => ({ label: l, checked: !!lab && lab.pos === k, action: () => set((c) => { c.labels = { ...(typeof c.labels === 'object' ? c.labels : {}), pos: k }; }) })),
    { sep: true }, { label: '기타 데이터 레이블 옵션...', action: () => openChartPane(o, 'labels') },
  ] });
  if (!pie && ch.kind !== 'scatter' && ch.kind !== 'bubble' && ch.kind !== 'radar') {
    items.push({ label: '데이터 표(B)', submenu: [
      { label: '없음', checked: !ch.dataTable, action: () => set((c) => { delete c.dataTable; }) },
      { label: '범례 표지 포함', checked: ch.dataTable === true, action: () => set((c) => { c.dataTable = true; }) },
      { label: '범례 표지 없음', checked: ch.dataTable === 'nokeys', action: () => set((c) => { c.dataTable = 'nokeys'; }) },
    ] });
  }
  if (!pie && ch.kind !== 'radar') {
    items.push({ label: '오차 막대(E)', submenu: [
      { label: '없음', checked: !ch.series.some((s) => s.err), action: () => set((c) => { for (const s of c.series) delete s.err; }) },
      { label: '표준 오차', action: () => set((c) => { for (const s of c.series) s.err = { type: 'stdErr' }; }) },
      { label: '백분율 (5%)', action: () => set((c) => { for (const s of c.series) s.err = { type: 'percent', val: 5 }; }) },
      { label: '표준 편차', action: () => set((c) => { for (const s of c.series) s.err = { type: 'stdDev', val: 1 }; }) },
    ] });
    const g = ch.grid ?? { major: true };
    items.push({ label: '눈금선(G)', submenu: [
      { label: '기본 주 가로', checked: g.major !== false, action: () => set((c) => { c.grid = { ...(c.grid ?? { major: true }), major: (c.grid ?? { major: true }).major === false }; }) },
      { label: '기본 주 세로', checked: !!g.cat, action: () => set((c) => { c.grid = { ...(c.grid ?? { major: true }), cat: !(c.grid ?? {}).cat || undefined }; }) },
      { label: '기본 보조 가로', checked: !!g.minor, action: () => set((c) => { c.grid = { ...(c.grid ?? { major: true }), minor: !(c.grid ?? {}).minor || undefined }; }) },
    ] });
  }
  items.push({ label: '범례(L)', submenu: [
    { label: '없음', checked: !ch.legend, action: () => set((c) => { c.legend = null; }) },
    ...[['r', '오른쪽'], ['t', '위쪽'], ['l', '왼쪽'], ['b', '아래쪽'], ['tr', '오른쪽 위']].map(([k, l]) => ({ label: l, checked: ch.legend === k, action: () => set((c) => { c.legend = k; }) })),
  ] });
  if (!pie && ch.kind !== 'radar') {
    items.push({ label: '추세선(T)', submenu: [
      { label: '없음', checked: !ch.series.some((s) => s.trend), action: () => set((c) => { for (const s of c.series) delete s.trend; }) },
      { label: '선형', action: () => set((c) => { c.series[0].trend = { type: 'linear' }; }) },
      { label: '지수', action: () => set((c) => { c.series[0].trend = { type: 'exp' }; }) },
      { label: '선형 예측', action: () => set((c) => { c.series[0].trend = { type: 'linear', forward: 2 }; }) },
      { label: '이동 평균', action: () => set((c) => { c.series[0].trend = { type: 'movingAvg', period: 2 }; }) },
      { sep: true }, { label: '기타 추세선 옵션...', action: () => openChartPane(o, 'ser0') },
    ] });
  }
  return items;
}

// ───────────── 갤러리 (빠른 레이아웃 · 스타일 · 색) ─────────────
const thumbOf = (c, w = 96, h = 64) => chartSvg({ ...c, title: c.title ? ' ' : '' }, w, h, S.pres.theme);
export function chartLayoutGallery(o) {
  return el('div', { class: 'chart-gal' }, CHART_LAYOUTS.map((_, i) => el('button', { class: `chart-gal-item${o.chart.layout === i + 1 ? ' on' : ''}`, title: `레이아웃 ${i + 1}`, type: 'button', onclick: () => { closeMenus(); setChart(o, (c) => applyChartLayout(c, i + 1)); } }, el('div', { html: thumbOf(applyChartLayout(o.chart, i + 1)) }))));
}
export function chartStyleGallery(o) {
  return el('div', { class: 'chart-gal' }, CHART_STYLES.map((_, i) => el('button', { class: `chart-gal-item${(o.chart.style ?? 1) === i + 1 ? ' on' : ''}`, title: `스타일 ${i + 1}`, type: 'button', onclick: () => { closeMenus(); setChart(o, (c) => applyChartStyle(c, i + 1)); } }, el('div', { html: thumbOf(applyChartStyle(o.chart, i + 1), 112, 72) }))));
}
export function chartPaletteItems(o) {
  const theme = S.pres.theme;
  return [
    { header: '다양한 색' },
    ...CHART_PALETTES.filter(([k]) => k.startsWith('colorful')).map(([k, l]) => paletteItem(o, k, l, theme)),
    { header: '단색' },
    ...CHART_PALETTES.filter(([k]) => k.startsWith('mono')).map(([k, l]) => paletteItem(o, k, l, theme)),
  ];
}
function paletteItem(o, k, l, theme) {
  return { node: el('button', { class: `menu-item pal-item${(o.chart.palette ?? 'colorful1') === k ? ' on' : ''}`, title: l, onclick: () => { closeMenus(); setChart(o, (c) => { c.palette = k === 'colorful1' ? undefined : k; for (const s of c.series) delete s.color; return c; }); } }, el('span', { class: 'pal-strip' }, Array.from({ length: 6 }, (_, i) => el('i', { style: { background: paletteColor(theme, k, i) } }))), el('span', {}, l)) };
}
export function resetChartStyle(o) {
  setChart(o, (c) => {
    for (const k of ['bg', 'plotBg', 'fontColor', 'gridColor', 'barOutline', 'dashGrid', 'shadow', 'palette', 'style', 'font', 'border', 'gap', 'overlap', 'titleSize', 'titleBold']) delete c[k];
    for (const s of c.series) { delete s.color; delete s.pointColors; delete s.lineWidth; delete s.dash; delete s.marker; delete s.markerSize; }
    return c;
  });
  toast('차트 스타일을 다시 설정했습니다');
}
export function switchChartRowCol(o) {
  if (['scatter', 'bubble'].includes(o.chart.kind)) { toast('분산형 · 거품형은 행/열을 전환할 수 없습니다'); return; }
  setChart(o, (c) => switchRowCol(c));
}
export function exportChartCsv(o) {
  const blob = new Blob([chartCsv(o.chart)], { type: 'text/csv' });
  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob);
  a.download = `${o.chart.title || '차트'}.csv`;
  document.body.append(a);
  a.click();
  setTimeout(() => { URL.revokeObjectURL(a.href); a.remove(); }, 1500);
}
export function chartFontDialog(o) {
  const ch = o.chart;
  const size = el('input', { type: 'number', min: '6', max: '72', step: '0.5', value: String(ch.font?.size ?? 12) });
  const tsize = el('input', { type: 'number', min: '6', max: '96', step: '0.5', value: String(ch.titleSize ?? 14) });
  let color = ch.fontColor ?? null;
  const cb = colorButton(color, (c) => { color = c; });
  const bold = el('input', { type: 'checkbox', checked: !!ch.titleBold });
  openDialog({
    title: '글꼴', width: 380,
    body: el('div', { class: 'form-grid' }, el('label', {}, el('span', {}, '크기 (pt)'), size), el('label', {}, el('span', {}, '제목 크기 (pt)'), tsize), el('label', {}, el('span', {}, '글꼴 색'), cb), el('label', { class: 'pane-check' }, bold, el('span', {}, '제목 굵게'))),
    buttons: [{ label: '확인', primary: true, action: () => {
      setChart(o, (c) => { c.font = { ...(c.font ?? {}), size: Number(size.value) || 12 }; c.titleSize = Number(tsize.value) || undefined; c.fontColor = color || undefined; c.titleBold = bold.checked || undefined; return c; });
      return true;
    } }, { label: '취소' }],
  });
}

// ───────────── [차트 서식] 창 ─────────────
export function openChartPane(o, part = 'chart') {
  S.chartPart = part;
  S.formatPane = 'chart';
  if (o && !S.sel.has(o.id)) { S.sel = new Set([o.id]); emit('selection'); }
  emit('pane');
}
const sec = (title, ...body) => el('section', { class: 'pane-sec' }, el('h4', {}, title), ...body);
const row = (label, ...ctrl) => el('label', { class: 'pane-row' }, el('span', {}, label), ...ctrl);
function numIn(value, onSet, { step = 1, min, max, suffix = '', blank = false } = {}) {
  const i = el('input', { type: 'number', value: value == null || value === '' ? '' : String(Math.round(value * 1000) / 1000), step: String(step), placeholder: blank ? '자동' : '', ...(min != null ? { min: String(min) } : {}), ...(max != null ? { max: String(max) } : {}) });
  i.addEventListener('change', () => { if (blank && i.value.trim() === '') { onSet(undefined); return; } const v = Number(i.value); if (Number.isFinite(v)) onSet(v); });
  i.addEventListener('keydown', (e) => e.stopPropagation());
  return suffix ? el('span', { class: 'num-wrap' }, i, el('small', {}, suffix)) : i;
}
function selIn(value, options, onSet) {
  const s = el('select', {}, options.map(([v, l]) => el('option', { value: v, selected: v === value }, l)));
  s.addEventListener('change', () => onSet(s.value));
  s.addEventListener('keydown', (e) => e.stopPropagation());
  return s;
}
function chk(value, onSet, label) {
  const c = el('input', { type: 'checkbox', checked: !!value });
  c.addEventListener('change', () => onSet(c.checked));
  return el('label', { class: 'pane-check' }, c, el('span', {}, label));
}
function txtIn(value, onSet) {
  const i = el('input', { type: 'text', value: value ?? '' });
  i.addEventListener('change', () => onSet(i.value));
  i.addEventListener('keydown', (e) => e.stopPropagation());
  return i;
}
const FMTS = [['', '일반'], ['0', '숫자 (0)'], ['0.0', '숫자 (0.0)'], ['#,##0', '1,000 단위'], ['#,##0.00', '1,000 단위 (0.00)'], ['0%', '백분율'], ['0.0%', '백분율 (0.0)'], ['₩#,##0', '통화 (₩)'], ['$#,##0', '통화 ($)'], ['#,##0"원"', '원'], ['#,##0"만"', '만']];

export function chartPane(body) {
  const o = selChart();
  if (!o) { body.append(el('p', { class: 'muted' }, '차트를 선택하면 차트 요소의 서식을 바꿀 수 있습니다.')); return; }
  const ch = o.chart;
  const pie = ch.kind === 'pie' || ch.kind === 'doughnut';
  const xy = ch.kind === 'scatter' || ch.kind === 'bubble';
  const has2 = ch.series.some((s) => s.axis === 'secondary');
  const parts = [['chart', '차트 영역'], ...(pie ? [] : [['plot', '그림 영역']]), ['title', '차트 제목'], ['legend', '범례'], ...(pie ? [] : [['axis-x', xy ? '가로 (값) 축' : '가로 (항목) 축'], ['axis-y', '세로 (값) 축']]), ...(has2 ? [['axis-y2', '보조 세로 (값) 축']] : []), ['labels', '데이터 레이블'], ...ch.series.map((s, i) => [`ser${i}`, `계열 "${s.name}"`]), ...(ch.dataTable ? [['table', '데이터 표']] : [])];
  let part = S.chartPart && parts.some(([k]) => k === S.chartPart) ? S.chartPart : 'chart';
  if (/^trend(\d+)$/.test(S.chartPart ?? '')) part = `ser${S.chartPart.slice(5)}`;
  const set = (fn) => setChart(o, (c) => { fn(c); return c; });
  const ax = (c, k) => { c.axes ??= {}; c.axes[k] ??= {}; return c.axes[k]; };
  body.append(el('div', { class: 'pane-row chart-part' }, el('span', {}, '요소'), selIn(part, parts, (v) => { S.chartPart = v; emit('pane'); })));
  if (part === 'chart') {
    body.append(sec('채우기 및 선', row('배경색', colorButton(ch.bg ?? null, (c) => set((x) => { x.bg = c || undefined; }), { auto: '없음' })), row('테두리', colorButton(ch.border ?? null, (c) => set((x) => { x.border = c || undefined; }), { auto: '없음' })), chk(ch.shadow, (v) => set((x) => { x.shadow = v || undefined; }), '그림자')));
    body.append(sec('글꼴', row('크기', numIn(ch.font?.size ?? null, (v) => set((x) => { x.font = { ...(x.font ?? {}), size: v }; if (v == null) delete x.font; }), { min: 6, max: 72, step: 0.5, suffix: 'pt', blank: true })), row('색', colorButton(ch.fontColor ?? null, (c) => set((x) => { x.fontColor = c || undefined; }), { auto: '자동' }))));
    body.append(sec('색', row('팔레트', selIn(ch.palette ?? 'colorful1', CHART_PALETTES.map(([k, l]) => [k, l]), (v) => set((x) => { x.palette = v === 'colorful1' ? undefined : v; }))), chk(ch.vary, (v) => set((x) => { x.vary = v || undefined; }), '요소마다 다른 색 사용')));
  } else if (part === 'plot') {
    body.append(sec('그림 영역', row('배경색', colorButton(ch.plotBg ?? null, (c) => set((x) => { x.plotBg = c || undefined; }), { auto: '없음' }))));
    if (ch.kind === 'col' || ch.kind === 'bar' || ch.kind === 'combo') body.append(sec('막대', row('계열 겹치기', numIn(ch.overlap ?? null, (v) => set((x) => { x.overlap = v == null ? undefined : Math.max(-100, Math.min(100, v)); }), { min: -100, max: 100, suffix: '%', blank: true })), row('간격 너비', numIn(ch.gap ?? null, (v) => set((x) => { x.gap = v == null ? undefined : Math.max(0, Math.min(500, v)); }), { min: 0, max: 500, suffix: '%', blank: true }))));
    if (!pie) {
      const g = ch.grid ?? { major: true };
      body.append(sec('눈금선', chk(g.major !== false, (v) => set((x) => { x.grid = { ...(x.grid ?? {}), major: v }; }), '주 눈금선 (가로)'), chk(g.minor, (v) => set((x) => { x.grid = { ...(x.grid ?? { major: true }), minor: v || undefined }; }), '보조 눈금선'), chk(g.cat, (v) => set((x) => { x.grid = { ...(x.grid ?? { major: true }), cat: v || undefined }; }), '세로 눈금선'), chk(ch.dashGrid, (v) => set((x) => { x.dashGrid = v || undefined; }), '점선'), row('색', colorButton(ch.gridColor ?? null, (c) => set((x) => { x.gridColor = c || undefined; }), { auto: '자동' }))));
    }
  } else if (part === 'title') {
    body.append(sec('차트 제목', chk(!!ch.title, (v) => set((x) => { x.title = v ? x.title || '차트 제목' : ''; }), '제목 표시'), row('글', txtIn(ch.title, (v) => set((x) => { x.title = v; }))), row('크기', numIn(ch.titleSize ?? null, (v) => set((x) => { x.titleSize = v; }), { min: 6, max: 96, step: 0.5, suffix: 'pt', blank: true })), chk(ch.titleBold, (v) => set((x) => { x.titleBold = v || undefined; }), '굵게'), chk(ch.titleOverlay, (v) => set((x) => { x.titleOverlay = v || undefined; }), '차트와 겹치게 (그림 영역을 줄이지 않음)')));
  } else if (part === 'legend') {
    body.append(sec('범례 옵션', selIn(ch.legend ?? '', [['', '없음'], ['t', '위쪽'], ['b', '아래쪽'], ['l', '왼쪽'], ['r', '오른쪽'], ['tr', '오른쪽 위']], (v) => set((x) => { x.legend = v || null; }))));
  } else if (part.startsWith('axis-')) {
    const k = part === 'axis-x' ? 'x' : part === 'axis-y' ? 'y' : 'y2';
    const a = ch.axes?.[k] ?? {};
    const valAxis = k !== 'x' || xy;
    body.append(sec('축 옵션',
      chk(!a.hide, (v) => set((x) => { ax(x, k).hide = !v || undefined; }), '축 표시'),
      row('축 제목', txtIn(a.title, (v) => set((x) => { ax(x, k).title = v || undefined; }))),
      ...(valAxis ? [
        row('최소값', numIn(a.min, (v) => set((x) => { ax(x, k).min = v; }), { blank: true, step: 'any' })),
        row('최대값', numIn(a.max, (v) => set((x) => { ax(x, k).max = v; }), { blank: true, step: 'any' })),
        row('주 단위', numIn(a.major, (v) => set((x) => { ax(x, k).major = v && v > 0 ? v : undefined; }), { blank: true, step: 'any' })),
        row('표시 형식', selIn(a.fmt ?? '', FMTS, (v) => set((x) => { ax(x, k).fmt = v || undefined; }))),
      ] : []),
      ...(k === 'x' && !xy ? [chk(a.reverse, (v) => set((x) => { ax(x, k).reverse = v || undefined; }), '항목을 거꾸로')] : [])));
  } else if (part === 'labels') {
    const lab = typeof ch.labels === 'object' ? ch.labels : ch.labels ? {} : null;
    const any = lab && (lab.val || lab.cat || lab.ser || lab.pct);
    const L = { val: any ? !!lab.val : !pie, cat: !!lab?.cat, ser: !!lab?.ser, pct: any ? !!lab.pct : pie };
    const setL = (k, v) => set((x) => { const cur = typeof x.labels === 'object' ? x.labels : {}; x.labels = { ...L, ...cur, [k]: v }; });
    const pos = pie ? [['', '자동'], ['ctr', '가운데'], ['inEnd', '안쪽 끝에'], ['outEnd', '바깥쪽 끝에'], ['bestFit', '자동 맞춤']] : ch.kind === 'line' || xy ? [['', '자동'], ['ctr', '가운데'], ['l', '왼쪽'], ['r', '오른쪽'], ['t', '위쪽'], ['b', '아래쪽']] : [['', '자동'], ['ctr', '가운데'], ['inEnd', '안쪽 끝에'], ['inBase', '축에 가깝게'], ['outEnd', '바깥쪽 끝에']];
    body.append(sec('레이블 옵션',
      chk(!!lab, (v) => set((x) => { x.labels = v ? {} : false; }), '데이터 레이블 표시'),
      lab ? [chk(L.ser, (v) => setL('ser', v), '계열 이름'), chk(L.cat, (v) => setL('cat', v), xy ? 'X 값' : '항목 이름'), chk(L.val, (v) => setL('val', v), '값'), ...(pie || ch.grouping === 'percent' ? [chk(L.pct, (v) => setL('pct', v), '백분율')] : []), row('레이블 위치', selIn(lab.pos ?? '', pos, (v) => set((x) => { x.labels = { ...(typeof x.labels === 'object' ? x.labels : {}), pos: v || undefined }; })))] : null));
  } else if (part === 'table') {
    body.append(sec('데이터 표', chk(!!ch.dataTable, (v) => set((x) => { x.dataTable = v || undefined; }), '데이터 표 표시'), chk(ch.dataTable === true, (v) => set((x) => { x.dataTable = v ? true : 'nokeys'; }), '범례 표지 표시')));
  } else if (part.startsWith('ser')) {
    const i = Number(part.slice(3));
    const s = ch.series[i];
    if (!s) return;
    const t = seriesType(ch, s, i);
    const ss = (fn) => set((x) => fn(x.series[i]));
    body.append(sec('계열 옵션',
      row('이름', txtIn(s.name, (v) => ss((y) => { y.name = v || y.name; }))),
      ...(ch.kind === 'combo' ? [row('차트 종류', selIn(t, [['col', '묶은 세로 막대형'], ['line', '꺾은선형'], ['area', '영역형']], (v) => ss((y) => { y.type = v; })))] : []),
      ...(!pie && ch.kind !== 'radar' ? [chk(s.axis === 'secondary', (v) => set((x) => { if (v) x.series[i].axis = 'secondary'; else delete x.series[i].axis; if (v && x.kind !== 'combo') { const k0 = x.kind; x.kind = 'combo'; x.series.forEach((y, j) => { y.type ??= j === i && k0 !== 'line' ? 'line' : k0 === 'bar' ? 'col' : ['col', 'line', 'area'].includes(k0) ? k0 : 'col'; }); } }), '보조 축')] : []),
      ...(pie ? [row('첫째 조각의 각', numIn(ch.firstAngle ?? 0, (v) => set((x) => { x.firstAngle = ((v % 360) + 360) % 360 || undefined; }), { min: 0, max: 360, suffix: '°' })), row('쪼개진 원형', numIn(ch.explode ?? 0, (v) => set((x) => { x.explode = Math.max(0, Math.min(400, v)) || undefined; delete x.variant; }), { min: 0, max: 400, suffix: '%' })), ...(ch.kind === 'doughnut' ? [row('도넛 구멍 크기', numIn(ch.hole ?? 50, (v) => set((x) => { x.hole = Math.max(10, Math.min(90, v)); }), { min: 10, max: 90, suffix: '%' }))] : [])] : [])));
    body.append(sec('채우기 및 선',
      row('색', colorButton(s.color ?? null, (c) => ss((y) => { if (c) y.color = c; else delete y.color; }), { auto: '자동' })),
      ...(t === 'line' || t === 'scatter' || t === 'radar' ? [
        row('선 너비', numIn(s.lineWidth ?? 2.25, (v) => ss((y) => { y.lineWidth = Math.max(0.25, v); }), { step: 0.25, min: 0.25, suffix: 'pt' })),
        row('대시', selIn(s.dash ?? '', [['', '실선'], ['dash', '파선'], ['dot', '점선'], ['dashDot', '파선-점선'], ['longDash', '긴 파선']], (v) => ss((y) => { y.dash = v || undefined; }))),
        row('표식', selIn(s.marker ?? '', [['', '자동'], ['none', '없음'], ['circle', '원'], ['square', '사각형'], ['diamond', '다이아몬드'], ['triangle', '삼각형'], ['x', 'X']], (v) => ss((y) => { y.marker = v || undefined; }))),
        row('표식 크기', numIn(s.markerSize ?? 5, (v) => ss((y) => { y.markerSize = Math.max(2, Math.min(72, v)); }), { min: 2, max: 72 })),
        ...(t === 'line' ? [chk(s.smooth ?? ch.variant === 'smooth', (v) => ss((y) => { y.smooth = v; }), '완만한 선')] : []),
      ] : [])));
    // 점 색 (원형 조각 · 막대 하나)
    body.append(sec('요소 색', el('div', { class: 'pt-colors' }, ch.cats.map((cat, k) => row(String(cat), colorButton(s.pointColors?.[k] ?? null, (c) => ss((y) => { y.pointColors = { ...(y.pointColors ?? {}) }; if (c) y.pointColors[k] = c; else delete y.pointColors[k]; if (!Object.keys(y.pointColors).length) delete y.pointColors; }), { auto: '자동' }))))));
    if (!pie && ch.kind !== 'radar') {
      const tr = s.trend ?? null;
      body.append(sec('추세선',
        selIn(tr?.type ?? '', [['', '없음'], ...Object.entries(TREND_LABEL)], (v) => ss((y) => { if (!v) delete y.trend; else y.trend = { ...(y.trend ?? {}), type: v }; })),
        tr ? [
          ...(tr.type === 'poly' ? [row('차수', numIn(tr.order ?? 2, (v) => ss((y) => { y.trend.order = Math.max(2, Math.min(6, v | 0)); }), { min: 2, max: 6 }))] : []),
          ...(tr.type === 'movingAvg' ? [row('구간', numIn(tr.period ?? 2, (v) => ss((y) => { y.trend.period = Math.max(2, v | 0); }), { min: 2 }))] : []),
          ...(tr.type !== 'movingAvg' ? [row('앞으로 예측', numIn(tr.forward ?? 0, (v) => ss((y) => { y.trend.forward = Math.max(0, v) || undefined; }), { min: 0, suffix: '구간' })), chk(tr.eq, (v) => ss((y) => { y.trend.eq = v || undefined; }), '수식을 차트에 표시'), chk(tr.r2, (v) => ss((y) => { y.trend.r2 = v || undefined; }), 'R-제곱 값을 차트에 표시')] : []),
          row('이름', txtIn(tr.name ?? '', (v) => ss((y) => { y.trend.name = v || undefined; }))),
        ] : null));
      const er = s.err ?? null;
      body.append(sec('오차 막대',
        selIn(er?.type ?? '', [['', '없음'], ['fixed', '고정 값'], ['percent', '백분율'], ['stdDev', '표준 편차'], ['stdErr', '표준 오차']], (v) => ss((y) => { if (!v) delete y.err; else y.err = { type: v, val: v === 'percent' ? 5 : 1 }; })),
        er && er.type !== 'stdErr' ? row('값', numIn(er.val ?? 1, (v) => ss((y) => { y.err.val = Math.max(0, v); }), { step: 'any', min: 0, suffix: er.type === 'percent' ? '%' : '' })) : null));
    }
    const lab = s.labels;
    body.append(sec('이 계열의 레이블', selIn(lab == null ? '' : lab ? 'on' : 'off', [['', '차트 설정 따름'], ['on', '표시'], ['off', '숨김']], (v) => ss((y) => { if (!v) delete y.labels; else y.labels = v === 'on' ? {} : false; }))));
  }
}

/** 편집 화면에서 차트 요소를 두 번 누르면 그 요소의 서식 창 */
export function chartPartAt(target) {
  const e = target?.closest?.('[data-el]');
  return e ? e.getAttribute('data-el') : null;
}

/** 3차원 회전 (PowerPoint [차트 영역 서식 › 효과 › 3차원 회전]): X/Y 회전 · 깊이 · 원근 + 미리 보기 */
export function chart3dDialog(o) {
  if (!o?.chart) return;
  const c0 = o.chart;
  const v = { ...VIEW3D_DEFAULT, ...(c0.kind === 'pie' ? { rotX: 30 } : {}), ...(c0.view3d ?? {}) };
  const num = (val, min, max, step = 5) => el('input', { type: 'number', min, max, step, value: val, style: { width: '70px' } });
  const on = el('input', { type: 'checkbox', checked: !!c0.d3 });
  const rx = num(v.rotX, 0, 80); const ry = num(v.rotY, 0, 360, 10); const dp = num(v.depth, 20, 500, 10); const ps = num(v.persp ?? 30, 0, 120);
  const view = el('div', { class: 'c3-prev' });
  const cur = () => ({ ...clone(c0), d3: on.checked || undefined, view3d: { rotX: Number(rx.value), rotY: Number(ry.value), depth: Number(dp.value), persp: Number(ps.value) } });
  const draw = () => { view.innerHTML = chartSvg(cur(), 360, 240, S.pres.theme); };
  for (const i of [on, rx, ry, dp, ps]) i.addEventListener('input', draw);
  openDialog({
    title: '3차원 회전', width: 420,
    body: el('div', { class: 'fx-opts' },
      el('label', { class: 'row' }, on, el('span', {}, '3차원으로 표시')),
      el('label', { class: 'row' }, el('span', {}, 'X 회전 (°)'), rx), el('label', { class: 'row' }, el('span', {}, 'Y 회전 (°)'), ry),
      el('label', { class: 'row' }, el('span', {}, '깊이 (기본값의 %)'), dp), el('label', { class: 'row' }, el('span', {}, '원근감 (°)'), ps), view),
    buttons: [{ label: '확인', primary: true, action: () => { const n = cur(); if (!n.d3) { delete n.d3; delete n.view3d; } setChart(o, () => n); } }, { label: '취소' }],
  });
  draw();
}
