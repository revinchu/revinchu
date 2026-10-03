import { paletteOf, CHART_PALETTES } from './chart.js';
import { el } from './ui.js';
import { buildChartHierarchy, hierarchyNodeColor } from './chart-hierarchy.js';
import { chartSeriesPatch, chartPointColorPatch, chartExplosionPatch } from './chart-edit.js';
import { chartAreaFormat } from './chart-area-format.js';
import { createChartAreaFormatPanel } from './chart-area-format-ui.js';

export function createChartSelectionPanel({ getChart, getPart, getData, onChange, onChoose, onDelete, onAllOptions }) {
  const body = el('div', { class: 'cfp chart-selection-pane' });
  const row = (name, input) => { input.setAttribute('aria-label', name); return el('label', { class: 'cfp-row' }, el('span', {}, name), input); };
  const section = (name, ...rows) => el('details', { class: 'cfp-sec', open: true }, el('summary', {}, name), ...rows);
  const up = patch => { if (getChart() && onChange(patch) === false) draw(); };
  const color = (value, change) => { const i = el('input', { type: 'color', value: /^#[0-9a-f]{6}$/i.test(value ?? '') ? value : '#4472c4' }); i.addEventListener('change', () => change(i.value)); return i; };
  const num = (value, change, min, max, step = 1) => { const i = el('input', { type: 'number', value: Number.isFinite(value) ? Math.round(value * 100) / 100 : value, min, max, step }); i.addEventListener('change', () => { if (!i.value || !Number.isFinite(Number(i.value))) return; const v = Math.max(min, Math.min(max, Number(i.value))); i.value = v; change(v); }); return i; };
  const text = (value, change) => { const i = el('input', { type: 'text', value: value ?? '' }); i.addEventListener('change', () => change(i.value)); return i; };
  const choose = (value, options, change) => { const i = el('select', {}, options.map(([v, label]) => el('option', { value: v, selected: String(v) === String(value) }, label))); i.addEventListener('change', () => change(i.value)); return i; };
  const check = (value, change) => { const i = el('input', { type: 'checkbox', checked: !!value }); i.addEventListener('change', () => change(i.checked)); return i; };
  const automatic = (value, change, min = -1e15, max = 1e15) => { const i = num(value ?? '', change, min, max, 'any'); i.placeholder = '자동'; i.addEventListener('change', () => { if (!i.value) change(undefined); }); return i; };
  const expanded = new Map();
  const areaTabs = new Map();
  const draw = () => {
    for (const item of body.querySelectorAll('details')) expanded.set(item.querySelector('summary')?.textContent, item.open);
    const focused = body.contains(document.activeElement) ? document.activeElement.getAttribute('aria-label') : null;
    const scrollBox = body.closest('.dialog-body'), scroll = scrollBox?.scrollTop ?? 0;
    const chart = getChart();
    if (!chart) { body.replaceChildren(el('p', { role: 'status' }, '차트가 삭제되었거나 다른 문서·시트로 이동했습니다. 차트를 다시 선택하세요.')); return; }
    const part = getPart() ?? { kind: 'chart' }, data = getData(), series = data.series.find(s => s._fi === part.s) ?? data.series[part.s];
    const fmt = chart.seriesFmt?.[part.s] ?? {}, point = Number(part.p ?? 0);
    const visiblePoint = series?._pi ? series._pi.indexOf(point) : point;
    const pointName = data.categories?.[visiblePoint] ?? point + 1;
    const hierarchy = ['treemap', 'sunburst'].includes(chart.type) ? buildChartHierarchy(data) : null;
    const node = hierarchy?.nodes.find(n => n.key === part.node);
    const noAxes = ['pie', 'doughnut', 'pieOfPie', 'barOfPie', 'sunburst', 'treemap', 'funnel', 'map', 'surface'].includes(chart.type);
    const title = part.kind === 'point' ? `데이터 요소 서식 — ${series?.name ?? ''} · ${pointName}` : part.kind === 'series' ? `데이터 계열 서식 — ${series?.name ?? ''}` : part.kind === 'node' ? `계층 항목 서식 — ${node?.path.filter(Boolean).join(' / ') ?? ''}${node?.children.length ? ' (상위 항목)' : ''}` : { 'axis-x': '가로·항목 축 서식', 'axis-y': '기본 값 축 서식', 'axis-y2': '보조 값 축 서식', title: '차트 제목 서식', legend: '범례 서식', label: '데이터 레이블 서식', dataTable: '데이터 표 서식', plot: '그림 영역 서식', chart: '차트 영역 서식' }[part.kind] ?? '차트 요소 서식';
    const options = [['chart', '차트 영역'], ['plot', '그림 영역'], ...(!noAxes ? [['axis-x', chart.type === 'bar' ? '세로(항목) 축' : '가로 축'], ['axis-y', chart.type === 'bar' ? '가로(값) 축' : '기본 세로(값) 축'], ...(data.series.some(s => s.axis === 1) || chart.type === 'pareto' || chart.volume ? [['axis-y2', '보조 값 축']] : [])] : []), ...(chart.title ? [['title', '차트 제목']] : []), ...(chart.legend !== 'none' ? [['legend', '범례']] : []), ...data.series.map((s, i) => [`series:${s._fi ?? i}`, `계열: ${s.name || i + 1}${chart.hiddenSeries?.includes(s._fi ?? i) ? ' (삭제됨)' : ''}`])];
    for (const [i, s] of data.series.entries()) options.push([`label:${s._fi ?? i}`, `데이터 레이블: ${s.name || i + 1}`]);
    if (part.kind === 'node') options.push([`node:${part.s}:${encodeURIComponent(part.node)}`, `계층 항목: ${node?.name ?? ''}`]);
    if (part.kind === 'point') options.push([`point:${part.s}:${point}`, `데이터 요소: ${pointName}`]);
    const value = part.kind === 'node' ? `node:${part.s}:${encodeURIComponent(part.node)}` : ['series', 'point', 'label'].includes(part.kind) ? `${part.kind}:${part.s}${part.kind === 'point' ? ':' + point : ''}` : part.kind;
    const picker = choose(value, options, v => { const [kind, s, p] = v.split(':'); onChoose({ kind, ...(s !== undefined ? { s: Number(s) } : {}), ...(kind === 'node' ? { node: decodeURIComponent(p) } : p !== undefined ? { p: Number(p) } : {}) }); });
    const rows = [], extra = [];
    const button = (label, action) => el('button', { type: 'button', class: 'btn small', onclick: action }, label);
    const setSeries = patch => up(chartSeriesPatch(getChart(), part.s, patch));
    const labelRows = () => [
      row('값 표시', check(fmt.labels ?? chart.labels ?? (chart.type === 'funnel'), v => setSeries({ labels: v }))),
      row('항목 이름', check(fmt.catName ?? (hierarchy ? chart.labels !== false : false), v => setSeries({ catName: v }))),
      row('계열 이름', check(fmt.serName, v => setSeries({ serName: v }))),
      ...(['pie', 'doughnut', 'pieOfPie', 'barOfPie', 'sunburst', 'treemap', 'column', 'bar', 'line', 'area', 'combo'].includes(chart.type) ? [row('백분율', check(fmt.pct, v => setSeries({ pct: v })))] : []),
      ...(!hierarchy && chart.type !== 'funnel' ? [row('레이블 위치', choose(fmt.labelPos ?? 'outEnd', [['outEnd', '바깥쪽 끝'], ['insideEnd', '안쪽 끝'], ['center', '가운데'], ['insideBase', '안쪽 기준'], ['out', '바깥쪽']], v => setSeries({ labelPos: v })))] : []),
      row('레이블 표시 형식', text(fmt.numFmt, v => setSeries({ numFmt: v || undefined }))),
      row('레이블 글꼴 색', color(fmt.labelColor ?? (hierarchy || chart.type === 'funnel' ? '#ffffff' : '#404040'), c => setSeries({ labelColor: c }))),
      row('레이블 크기(pt)', num(fmt.labelSize ?? (hierarchy ? 8.25 : 9), n => setSeries({ labelSize: n }), 6, 72)),
      row('레이블 굵게', check(fmt.labelBold ?? (chart.type === 'funnel'), v => setSeries({ labelBold: v }))),
      ...(chart.type === 'funnel' ? [el('p', { class: 'muted' }, '깔때기 데이터 레이블은 가운데에 표시됩니다.')] : []),
    ];
    if (part.kind === 'node') {
      rows.push(row('계층 항목 색', color(node ? hierarchyNodeColor(node, series, chart, paletteOf(chart)) : '#4472c4', c => setSeries({ hierarchyColors: { ...getChart().seriesFmt?.[part.s]?.hierarchyColors, [part.node]: c } }))));
      rows.push(button('계층 항목 색 자동으로', () => { const colors = { ...getChart().seriesFmt?.[part.s]?.hierarchyColors }; delete colors[part.node]; setSeries({ hierarchyColors: Object.keys(colors).length ? colors : undefined }); draw(); }));
      rows.push(row('계층 항목', choose(part.node, hierarchy.nodes.map(n => [n.key, `${'　'.repeat(Math.min(n.depth, 8))}${n.name || '(빈 항목)'}`]), key => onChoose({ kind: 'node', s: part.s, node: key }))));
      rows.push(el('p', { class: 'muted' }, '상위 항목의 색은 하위 항목에도 적용됩니다. 별도로 지정한 하위 색은 유지됩니다.'));
      extra.push(section('데이터 레이블', ...labelRows()));
    } else if (part.kind.startsWith('axis-')) {
      const key = part.kind.slice(5), axis = chart.axes?.[key] ?? {};
      const setAxis = patch => { const axes = getChart().axes ?? {}; up({ axes: { ...axes, [key]: { ...axes[key], ...patch } } }); };
      rows.push(row('축 표시', check(!axis.hide, v => setAxis({ hide: !v }))));
      rows.push(row('축 제목', text(axis.title, v => setAxis({ title: v || undefined }))));
      if (key !== 'x' || ['scatter', 'bubble'].includes(chart.type)) {
        rows.push(row('최소값', automatic(axis.min, v => setAxis({ min: v }))));
        rows.push(row('최대값', automatic(axis.max, v => setAxis({ max: v }))));
        rows.push(row('주 단위', automatic(axis.major, v => setAxis({ major: v && v > 0 ? v : undefined }), 0)));
        rows.push(row('축 표시 형식', text(axis.numFmt, v => setAxis({ numFmt: v || undefined }))));
        rows.push(button('축 범위 자동으로', () => { setAxis({ min: undefined, max: undefined, major: undefined }); draw(); }));
      }
      rows.push(row('역순으로 표시', check(axis.reverse, v => setAxis({ reverse: v }))));
      rows.push(row('축 글꼴 크기(pt)', num(chart.axisSize ?? 9, v => up({ axisSize: v }), 6, 24)));
      const grid = key === 'x' ? 'gridX' : 'gridY';
      rows.push(row('주 눈금선', check(key === 'x' ? !!chart.gridX : chart.gridY !== false, v => up({ [grid]: v }))));
      rows.push(row('눈금선 색', color(chart.gridColor ?? '#d9d9d9', v => up({ gridColor: v }))));
      rows.push(button('축 삭제', onDelete));
    } else if (part.kind === 'series' || part.kind === 'point') {
      const type = series?.type ?? chart.type, pie = ['pie', 'doughnut'].includes(chart.type), line = ['line', 'scatter', 'radar'].includes(type);
      const label = part.kind === 'point' ? line ? '선택한 표식 색' : '선택한 요소 색' : line ? '계열 선 색' : '계열 채우기 색';
      rows.push(row(label, color(part.kind === 'point' ? fmt.pointColors?.[point] ?? (pie ? fmt.color ?? series?.colors?.[point] ?? paletteOf(chart)[point % paletteOf(chart).length] : series?.color) : series?.color, c => part.kind === 'point' ? up(chartPointColorPatch(getChart(), part.s, point, c)) : setSeries({ color: c, grad: undefined }))));
      if (part.kind === 'point') {
        rows.push(el('button', { class: 'btn small', onclick: () => { up(chartPointColorPatch(getChart(), part.s, point, undefined)); draw(); } }, '요소 색 자동으로'));
        rows.push(row('데이터 요소', choose(point, data.categories.map((name, p) => [series?._pi?.[p] ?? p, `${(series?._pi?.[p] ?? p) + 1}. ${String(name).slice(0, 70)}`]), p => onChoose({ kind: 'point', s: part.s, p: Number(p) }))));
      } else if (type !== 'area' && data.categories.length) rows.push(el('button', { class: 'btn small', onclick: () => onChoose({ kind: 'point', s: part.s, p: series?._pi?.[0] ?? 0 }) }, '개별 데이터 요소 선택'));
      if (pie) {
        const current = part.kind === 'point' ? fmt.pointExplosion?.[point] ?? fmt.explode ?? chart.explode ?? 0 : fmt.explode ?? chart.explode ?? 0;
        rows.push(row(part.kind === 'point' ? '선택한 조각 분리(%)' : '계열 조각 분리(%)', num(current, n => up(chartExplosionPatch(getChart(), part, n)), 0, 400)));
        if (part.kind === 'point') rows.push(el('button', { class: 'btn small', onclick: () => { const points = { ...(getChart().seriesFmt?.[part.s]?.pointExplosion ?? {}) }; delete points[point]; setSeries({ pointExplosion: Object.keys(points).length ? points : undefined }); draw(); } }, '조각 분리 기본값으로'));
        rows.push(row('첫째 조각 각(°)', num(chart.firstAngle ?? 0, n => up({ firstAngle: n }), 0, 360)));
      }
      if (line && part.kind === 'series') rows.push(row('선 두께(pt)', num((fmt.lineWidth ?? 2.25) * .75, n => setSeries({ lineWidth: n / .75 }), .25, 30, .25)));
      if (part.kind === 'series') {
        const settings = [];
        if (!chart.threeD && ['column', 'line', 'area', 'combo'].includes(chart.type)) {
          settings.push(row('계열 차트 종류', choose(type, [['column', '세로 막대형'], ['line', '꺾은선형'], ['area', '영역형']], v => { setSeries({ type: v }); draw(); })));
          settings.push(row('계열 표시 축', choose(fmt.axis ?? series?.axis ?? 0, [[0, '기본 축 (왼쪽)'], [1, '보조 축 (오른쪽)']], v => { setSeries({ axis: Number(v) }); draw(); })));
          if (type !== 'line') settings.push(row('계열 배치', choose(fmt.grouping ?? series?.grouping ?? chart.grouping ?? 'clustered', [['clustered', '묶은'], ['stacked', '누적'], ['percentStacked', '100% 기준 누적']], v => setSeries({ grouping: v }))));
        }
        if (line) {
          settings.push(row('선 종류', choose(fmt.dash ?? '', [['', '실선'], ['dash', '파선'], ['dot', '점선'], ['dashDot', '일점쇄선']], v => setSeries({ dash: v || undefined }))));
          settings.push(row('표식 모양', choose(fmt.marker ?? chart.marker ?? (type === 'scatter' && /^(marker|.*Marker)$/.test(chart.scatterStyle ?? 'marker') ? 'circle' : 'none'), [['none', '없음'], ['circle', '원'], ['square', '사각형'], ['diamond', '마름모'], ['triangle', '삼각형']], v => setSeries({ marker: v }))));
          settings.push(row('표식 크기', num(fmt.markerSize ?? 6, v => setSeries({ markerSize: v }), 2, 30)));
          if (['line', 'scatter'].includes(type)) settings.push(row('부드러운 선', check(fmt.smooth ?? (type === 'scatter' && /smooth/i.test(chart.scatterStyle ?? '')), v => setSeries({ smooth: v }))));
        }
        if (['column', 'bar'].includes(type)) {
          settings.push(row('간격 너비(%)', num(chart.gap ?? 150, v => up({ gap: v }), 0, 500)));
          if (!chart.threeD) settings.push(row('계열 겹치기(%)', num(chart.overlap ?? 0, v => up({ overlap: v }), -100, 100)));
          else settings.push(row('막대 모양', choose(fmt.barShape ?? chart.barShape ?? 'box', [['box', '상자'], ['cylinder', '원통'], ['cone', '원뿔'], ['pyramid', '피라미드']], v => setSeries({ barShape: v }))));
        }
        if (chart.type === 'doughnut') settings.push(row('도넛 구멍 크기(%)', num(chart.hole ?? 50, v => up({ hole: v }), 10, 90)));
        if (chart.type === 'treemap') settings.push(row('상위 항목 레이블', choose(chart.treemapLabelLayout ?? 'banner', [['banner', '배너'], ['overlapping', '겹치기'], ['none', '없음']], v => up({ treemapLabelLayout: v }))));
        if (settings.length) extra.push(section('계열 옵션', ...settings));
        extra.push(section('데이터 레이블', ...labelRows()));
      }
      if (chart.hiddenSeries?.includes(part.s)) rows.push(el('button', { class: 'btn', onclick: () => { up({ hiddenSeries: getChart().hiddenSeries.filter(i => i !== part.s) }); draw(); } }, '계열 복원'));
      else rows.push(el('button', { class: 'btn', onclick: onDelete }, '계열 삭제'));
      rows.push(el('p', { class: 'muted' }, '한 번 클릭하면 계열, 같은 계열을 다시 클릭하면 데이터 요소를 선택합니다. Delete는 선택한 계열을 제거하며 원본 셀은 유지합니다.'));
    } else if (part.kind === 'title' || part.kind === 'legend') {
      const titlePart = part.kind === 'title', prefix = titlePart ? 'title' : 'legend', layoutKey = prefix + 'Layout', layout = chart[layoutKey];
      if (titlePart) rows.push(row('제목 텍스트', text(chart.title, value => up({ title: value }))));
      rows.push(row('글꼴 색', color(chart[prefix + 'Color'] ?? '#333333', value => up({ [prefix + 'Color']: value }))));
      rows.push(row('글꼴 크기(pt)', num(chart[prefix + 'Size'] ?? (titlePart ? 14 : 9), value => up({ [prefix + 'Size']: value }), 6, 72)));
      const bold = el('input', { type: 'checkbox', checked: !!chart[prefix + 'Bold'] }); bold.addEventListener('change', () => up({ [prefix + 'Bold']: bold.checked })); rows.push(row('굵게', bold));
      if (!titlePart) rows.push(row('범례 위치', choose(chart.legend ?? 'b', [['b', '아래쪽'], ['t', '위쪽'], ['l', '왼쪽'], ['r', '오른쪽']], value => { up({ legend: value, legendLayout: undefined }); draw(); })));
      if (layout) {
        rows.push(row('가로 위치(%)', num(layout.x * 100, x => up({ [layoutKey]: { ...getChart()[layoutKey], x: x / 100 } }), 0, 100, .1)));
        rows.push(row('세로 위치(%)', num(layout.y * 100, y => up({ [layoutKey]: { ...getChart()[layoutKey], y: y / 100 } }), 0, 100, .1)));
      }
      rows.push(el('button', { class: 'btn small', onclick: () => { up({ [layoutKey]: undefined }); draw(); } }, '자동 위치로 되돌리기'));
      rows.push(el('button', { class: 'btn', onclick: onDelete }, titlePart ? '제목 삭제' : '범례 삭제'));
      rows.push(el('p', { class: 'muted' }, '차트에서 선택한 제목·범례를 끌거나 방향키로 이동할 수 있습니다.'));
      if (['waterfall', 'histogram', 'pareto', 'boxWhisker', 'treemap', 'sunburst', 'funnel', 'map'].includes(chart.type)) rows.push(el('p', { class: 'muted' }, '이 차트 종류의 자유 위치는 위셀에서 보존됩니다. Excel에서는 기본 위치로 표시될 수 있습니다.'));
    } else if (part.kind === 'label') {
      rows.push(...labelRows());
      rows.push(el('button', { class: 'btn', onclick: onDelete }, '계열 데이터 레이블 삭제'));
    } else if (part.kind === 'plot') {
      extra.push(createChartAreaFormatPanel({kind:'plot',getFormat:()=>chartAreaFormat(getChart(),'plot'),getIdentity:()=>`${getChart()?.id}:${getPart()?.kind}`,initialTab:areaTabs.get('plot'),onTab:tab=>areaTabs.set('plot',tab),onChange:value=>up({plotAreaFormat:value??undefined,...(value?{}:{plotFill:undefined})})}).body);
    } else {
      rows.push(row('색 구성', choose(Array.isArray(chart.palette) ? 'imported' : chart.palette ?? 'office', [...(Array.isArray(chart.palette) ? [['imported', '가져온 색']] : []), ...Object.entries(CHART_PALETTES).map(([k, p]) => [k, p.label])], v => { if (v !== 'imported') up({ palette: v }); })));
      extra.push(createChartAreaFormatPanel({kind:'chart',getFormat:()=>chartAreaFormat(getChart()),getIdentity:()=>`${getChart()?.id}:${getPart()?.kind}`,initialTab:areaTabs.get('chart'),onTab:tab=>areaTabs.set('chart',tab),onChange:value=>up({chartAreaFormat:value??undefined,...(value?{}:{fill:undefined,border:undefined})})}).body);
    }
    body.replaceChildren(row('서식을 지정할 차트 요소', picker), el('p', { class: 'cfp-selection-name', role: 'status' }, title), ...(rows.length ? [section('선택한 요소', ...rows)] : []), ...extra, el('button', { class: 'btn', onclick: onAllOptions }, '차트 전체 옵션…'));
    for (const item of body.querySelectorAll('details')) if (expanded.has(item.querySelector('summary')?.textContent)) item.open = expanded.get(item.querySelector('summary')?.textContent);
    if (focused) [...body.querySelectorAll('[aria-label]')].find(n => n.getAttribute('aria-label') === focused)?.focus({ preventScroll: true });
    if (scrollBox) scrollBox.scrollTop = scroll;
  };
  for (const name of ['change', 'click']) body.addEventListener(name, event => { if (!getChart()) { event.preventDefault(); event.stopImmediatePropagation(); draw(); } }, true);
  draw(); return { body, refresh: draw };
}
