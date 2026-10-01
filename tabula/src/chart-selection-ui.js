import { paletteOf } from './chart.js';
import { el } from './ui.js';
import { chartSeriesPatch, chartPointColorPatch, chartExplosionPatch } from './chart-edit.js';

export function createChartSelectionPanel({ getChart, getPart, getData, onChange, onChoose, onDelete, onAllOptions }) {
  const body = el('div', { class: 'cfp chart-selection-pane' });
  const row = (name, input) => { input.setAttribute('aria-label', name); return el('label', { class: 'cfp-row' }, el('span', {}, name), input); };
  const section = (name, ...rows) => el('details', { class: 'cfp-sec', open: true }, el('summary', {}, name), ...rows);
  const up = patch => { if (getChart()) onChange(patch); };
  const color = (value, change) => { const i = el('input', { type: 'color', value: /^#[0-9a-f]{6}$/i.test(value ?? '') ? value : '#4472c4' }); i.addEventListener('change', () => change(i.value)); return i; };
  const num = (value, change, min, max, step = 1) => { const i = el('input', { type: 'number', value: Number.isFinite(value) ? Math.round(value * 100) / 100 : value, min, max, step }); i.addEventListener('change', () => { if (!i.value || !Number.isFinite(Number(i.value))) return; const v = Math.max(min, Math.min(max, Number(i.value))); i.value = v; change(v); }); return i; };
  const text = (value, change) => { const i = el('input', { type: 'text', value: value ?? '' }); i.addEventListener('change', () => change(i.value)); return i; };
  const choose = (value, options, change) => { const i = el('select', {}, options.map(([v, label]) => el('option', { value: v, selected: String(v) === String(value) }, label))); i.addEventListener('change', () => change(i.value)); return i; };
  const draw = () => {
    const chart = getChart();
    if (!chart) { body.replaceChildren(el('p', { role: 'status' }, '차트가 삭제되었거나 다른 문서·시트로 이동했습니다. 차트를 다시 선택하세요.')); return; }
    const part = getPart() ?? { kind: 'chart' }, data = getData(), series = data.series.find(s => s._fi === part.s) ?? data.series[part.s];
    const fmt = chart.seriesFmt?.[part.s] ?? {}, point = Number(part.p ?? 0);
    const title = part.kind === 'point' ? `데이터 요소 서식 — ${series?.name ?? ''} · ${data.categories?.[point] ?? point + 1}` : part.kind === 'series' ? `데이터 계열 서식 — ${series?.name ?? ''}` : { title: '차트 제목 서식', legend: '범례 서식', label: '데이터 레이블 서식', dataTable: '데이터 표 서식', plot: '그림 영역 서식', chart: '차트 영역 서식' }[part.kind] ?? '차트 요소 서식';
    const options = [['chart', '차트 영역'], ...(chart.title ? [['title', '차트 제목']] : []), ...(chart.legend !== 'none' ? [['legend', '범례']] : []), ...data.series.map((s, i) => [`series:${s._fi ?? i}`, `계열: ${s.name || i + 1}${chart.hiddenSeries?.includes(s._fi ?? i) ? ' (삭제됨)' : ''}`])];
    if (part.kind === 'label') options.push(['label', '데이터 레이블']);
    if (part.kind === 'plot') options.push(['plot', '그림 영역']);
    if (part.kind === 'point') options.push([`point:${part.s}:${point}`, `데이터 요소: ${data.categories?.[point] ?? point + 1}`]);
    const value = ['series', 'point'].includes(part.kind) ? `${part.kind}:${part.s}${part.kind === 'point' ? ':' + point : ''}` : part.kind;
    const picker = choose(value, options, v => { const [kind, s, p] = v.split(':'); onChoose({ kind, ...(s !== undefined ? { s: Number(s) } : {}), ...(p !== undefined ? { p: Number(p) } : {}) }); });
    const rows = [];
    const setSeries = patch => up(chartSeriesPatch(getChart(), part.s, patch));
    if (part.kind === 'series' || part.kind === 'point') {
      const type = series?.type ?? chart.type, pie = ['pie', 'doughnut'].includes(chart.type), line = ['line', 'scatter', 'radar'].includes(type);
      const label = part.kind === 'point' ? line ? '선택한 표식 색' : '선택한 요소 색' : line ? '계열 선 색' : '계열 채우기 색';
      rows.push(row(label, color(part.kind === 'point' ? fmt.pointColors?.[point] ?? (pie ? fmt.color ?? series?.colors?.[point] ?? paletteOf(chart)[point % paletteOf(chart).length] : series?.color) : series?.color, c => part.kind === 'point' ? up(chartPointColorPatch(getChart(), part.s, point, c)) : setSeries({ color: c, grad: undefined }))));
      if (part.kind === 'point') {
        rows.push(el('button', { class: 'btn small', onclick: () => { up(chartPointColorPatch(getChart(), part.s, point, undefined)); draw(); } }, '요소 색 자동으로'));
        rows.push(row('데이터 요소', choose(point, data.categories.map((name, p) => [p, `${p + 1}. ${String(name).slice(0, 70)}`]), p => onChoose({ kind: 'point', s: part.s, p: Number(p) }))));
      } else if (type !== 'area' && data.categories.length) rows.push(el('button', { class: 'btn small', onclick: () => onChoose({ kind: 'point', s: part.s, p: 0 }) }, '개별 데이터 요소 선택'));
      if (pie) {
        const current = part.kind === 'point' ? fmt.pointExplosion?.[point] ?? fmt.explode ?? chart.explode ?? 0 : fmt.explode ?? chart.explode ?? 0;
        rows.push(row(part.kind === 'point' ? '선택한 조각 분리(%)' : '계열 조각 분리(%)', num(current, n => up(chartExplosionPatch(getChart(), part, n)), 0, 400)));
        if (part.kind === 'point') rows.push(el('button', { class: 'btn small', onclick: () => { const points = { ...(getChart().seriesFmt?.[part.s]?.pointExplosion ?? {}) }; delete points[point]; setSeries({ pointExplosion: Object.keys(points).length ? points : undefined }); draw(); } }, '조각 분리 기본값으로'));
        rows.push(row('첫째 조각 각(°)', num(chart.firstAngle ?? 0, n => up({ firstAngle: n }), 0, 360)));
      }
      if (line && part.kind === 'series') rows.push(row('선 두께(pt)', num((fmt.lineWidth ?? 2.25) * .75, n => setSeries({ lineWidth: n / .75 }), .25, 30, .25)));
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
      rows.push(row('레이블 글꼴 색', color(fmt.labelColor ?? '#404040', c => setSeries({ labelColor: c }))));
      rows.push(row('레이블 크기(pt)', num(fmt.labelSize ?? 9, n => setSeries({ labelSize: n }), 6, 72)));
      rows.push(el('button', { class: 'btn', onclick: onDelete }, '계열 데이터 레이블 삭제'));
    } else if (part.kind === 'plot') {
      rows.push(row('그림 영역 채우기 색', color(chart.plotFill ?? '#ffffff', c => up({ plotFill: c }))));
    } else {
      rows.push(row('차트 채우기 색', color(chart.fill ?? '#ffffff', value => up({ fill: value }))));
      rows.push(row('차트 테두리 색', color(chart.border ?? '#cccccc', value => up({ border: value }))));
    }
    body.replaceChildren(row('서식을 지정할 차트 요소', picker), el('p', { class: 'cfp-selection-name', role: 'status' }, title), section('선택한 요소', ...rows), el('button', { class: 'btn', onclick: onAllOptions }, '차트 전체 옵션…'));
  };
  for (const name of ['change', 'click']) body.addEventListener(name, event => { if (!getChart()) { event.preventDefault(); event.stopImmediatePropagation(); draw(); } }, true);
  draw(); return { body, refresh: draw };
}
