// 차트: 범위 값 → 데이터 해석, SVG 그리기 (DOM 없이 문자열 생성)
// 차트 모델: { type, title, range, sheet, x, y, w, h, z,
//   series?: [{ name: { ref | text }, cat: ref, val: ref, x: ref }]   파일에서 가져온 계열 참조 (ref = { sheet, r1, c1, r2, c2 })
//   pivot?: { sheet, name }                                           피벗 차트 (피벗 결과를 따라감)
//   seriesFmt?: [{ type, axis, color, labels, marker, smooth, numFmt, explode, pointExplosion, pointColors }] 계열별 서식 (원본 순서)
//   titleLayout?, legendLayout?: { x, y, w?, h? } 차트 크기에 대한 좌상단/크기 비율 (edge)
//   labels, legend: 'b'|'t'|'r'|'l'|'none', grouping: 'clustered'|'stacked'|'percentStacked',
//   axes?: { y: { title, numFmt, min, max }, y2: {…}, x: { title } } }
import { formatGeneral, formatValue, formatCode, isDateCode } from './format.js';
import { categoryAxisLayout, categoryAxisLabelSvg } from './chart-axis-labels.js';
import { pivotSourceData, pivotChartData } from './pivot.js';
import { inferPivotCategorySeries } from './chart-source.js';
import { maxOf, minOf, pushAll } from './fxcore.js';
import { THEME, applyTint } from './stylepresets.js';
import { chartDepth, extrudedPolygon, barSolid3D, chartWalls3D, pieProjection3D, pieSolid3D } from './chart-3d.js';
import { ADVANCED_CHARTS, drawVolumeStock } from './chart-advanced.js';
import { MAP_CHARTS } from './chart-map.js';
import { hierarchyCategories, filterHierarchyData, drawHierarchyTreemap, hierarchyLegend } from './chart-hierarchy.js';
import { chartAreaFormat, chartAreaSvg } from './chart-area-format.js';
import { categoryTrendPoints } from './chart-trend.js';
import { EXTRA_CHART_PALETTES, normalizeChartPaletteColor } from './chart-palette-options.js';

export const CHART_TYPES = [
  { id: 'column', label: '세로 막대형' },
  { id: 'bar', label: '가로 막대형' },
  { id: 'line', label: '꺾은선형' },
  { id: 'area', label: '영역형' },
  { id: 'pie', label: '원형' },
  { id: 'doughnut', label: '도넛형' },
  { id: 'scatter', label: '분산형' },
  { id: 'bubble', label: '거품형' },
  { id: 'radar', label: '방사형' },
  { id: 'waterfall', label: '폭포' },
  { id: 'funnel', label: '깔때기형' },
  { id: 'histogram', label: '히스토그램' },
  { id: 'pareto', label: '파레토' },
  { id: 'treemap', label: '트리맵' },
  { id: 'map', label: '지도 (국가/지역)' },
  { id: 'sunburst', label: '선버스트' },
  { id: 'surface', label: '표면형' },
  { id: 'pieOfPie', label: '대조 원형' },
  { id: 'barOfPie', label: '대조 가로 막대형' },
  { id: 'boxWhisker', label: '상자 수염' },
  { id: 'stock', label: '주식형' },
  { id: 'combo', label: '콤보' },
];

/** 엑셀 [모든 차트] 대화상자와 같은 분류 · 하위 종류 (차트 모델에 덮어쓸 값) */
export const CHART_GALLERY = [
  ['세로 막대형', [['묶은 세로 막대형', { threeD: false, type: 'column', grouping: 'clustered' }], ['누적 세로 막대형', { threeD: false, type: 'column', grouping: 'stacked' }], ['100% 기준 누적 세로 막대형', { threeD: false, type: 'column', grouping: 'percentStacked' }]]],
  ['꺾은선형', [['꺾은선형', { threeD: false, type: 'line', grouping: 'clustered', marker: 'none' }], ['누적 꺾은선형', { threeD: false, type: 'line', grouping: 'stacked', marker: 'none' }], ['100% 기준 누적 꺾은선형', { threeD: false, type: 'line', grouping: 'percentStacked', marker: 'none' }], ['표식이 있는 꺾은선형', { threeD: false, type: 'line', grouping: 'clustered', marker: 'circle' }], ['표식이 있는 누적 꺾은선형', { threeD: false, type: 'line', grouping: 'stacked', marker: 'circle' }], ['표식이 있는 100% 기준 누적 꺾은선형', { threeD: false, type: 'line', grouping: 'percentStacked', marker: 'circle' }]]],
  ['원형', [['원형', { threeD: false, type: 'pie' }], ['쪼개진 원형', { threeD: false, type: 'pie', explode: 12 }], ['도넛형', { threeD: false, type: 'doughnut' }], ['쪼개진 도넛형', { threeD: false, type: 'doughnut', explode: 12 }], ['대조 원형', { threeD: false, type: 'pieOfPie', splitType: 'position', splitPos: 3 }], ['대조 가로 막대형', { threeD: false, type: 'barOfPie', splitType: 'position', splitPos: 3 }]]],
  ['가로 막대형', [['묶은 가로 막대형', { threeD: false, type: 'bar', grouping: 'clustered' }], ['누적 가로 막대형', { threeD: false, type: 'bar', grouping: 'stacked' }], ['100% 기준 누적 가로 막대형', { threeD: false, type: 'bar', grouping: 'percentStacked' }]]],
  ['영역형', [['영역형', { threeD: false, type: 'area' }], ['누적 영역형', { threeD: false, type: 'area', grouping: 'stacked' }], ['100% 기준 누적 영역형', { threeD: false, type: 'area', grouping: 'percentStacked' }]]],
  ['분산형', [['분산형', { threeD: false, type: 'scatter' }], ['곡선 및 표식이 있는 분산형', { threeD: false, type: 'scatter', scatterStyle: 'smoothMarker' }], ['곡선이 있는 분산형', { threeD: false, type: 'scatter', scatterStyle: 'smooth' }], ['직선 및 표식이 있는 분산형', { threeD: false, type: 'scatter', scatterStyle: 'lineMarker' }], ['직선이 있는 분산형', { threeD: false, type: 'scatter', scatterStyle: 'line' }], ['거품형', { threeD: false, type: 'bubble' }], ['3차원 효과가 있는 거품형', { threeD: true, type: 'bubble' }]]],
  ['주식형', [['고가-저가-종가', { threeD: false, type: 'stock', ohlc: false, volume: false }], ['시가-고가-저가-종가', { threeD: false, type: 'stock', ohlc: true, volume: false }], ['거래량-고가-저가-종가', { threeD: false, type: 'stock', ohlc: false, volume: true }], ['거래량-시가-고가-저가-종가', { threeD: false, type: 'stock', ohlc: true, volume: true }]]],
  ['방사형', [['방사형', { threeD: false, type: 'radar' }], ['표식이 있는 방사형', { threeD: false, type: 'radar', radarStyle: 'marker' }], ['채워진 방사형', { threeD: false, type: 'radar', radarStyle: 'filled' }]]],
  ['트리맵', [['트리맵', { threeD: false, type: 'treemap' }]]],
  ['지도', [['국가/지역 색칠지도', { threeD: false, type: 'map', legend: 'b' }]]],
  ['선버스트', [['선버스트', { threeD: false, type: 'sunburst' }]]],
  ['표면형', [['3차원 표면형', { threeD: true, type: 'surface', surfaceStyle: 'surface' }], ['3차원 표면형(골격형)', { threeD: true, type: 'surface', surfaceStyle: 'wireframe' }], ['등고선형', { threeD: false, type: 'surface', surfaceStyle: 'contour' }], ['등고선형(골격형)', { threeD: false, type: 'surface', surfaceStyle: 'wireframeContour' }]]],
  ['히스토그램', [['히스토그램', { threeD: false, type: 'histogram' }], ['파레토', { threeD: false, type: 'pareto' }]]],
  ['상자 수염', [['상자 수염', { threeD: false, type: 'boxWhisker' }]]],
  ['폭포', [['폭포', { threeD: false, type: 'waterfall' }]]],
  ['깔때기형', [['깔때기형', { threeD: false, type: 'funnel' }]]],
  ['콤보', [['묶은 세로 막대형 - 꺾은선형, 보조 축', { threeD: false, type: 'combo', grouping: 'clustered', comboAxis: 'secondary', comboLayout: 'columnLine' }], ['묶은 세로 막대형 - 꺾은선형, 기본 축', { threeD: false, type: 'combo', grouping: 'clustered', comboAxis: 'primary', comboLayout: 'columnLine' }], ['누적 세로 막대형 - 꺾은선형, 보조 축', { threeD: false, type: 'combo', grouping: 'stacked', comboAxis: 'secondary', comboLayout: 'stackedColumnLine' }], ['누적 세로 막대형 - 꺾은선형, 기본 축', { threeD: false, type: 'combo', grouping: 'stacked', comboAxis: 'primary', comboLayout: 'stackedColumnLine' }], ['누적 영역형 - 묶은 세로 막대형', { threeD: false, type: 'combo', grouping: 'clustered', comboAxis: 'primary', comboLayout: 'areaColumn' }]]],
];
for (const [name, entries] of CHART_GALLERY) {
  if (!['세로 막대형', '가로 막대형', '원형', '영역형', '꺾은선형'].includes(name)) continue;
  const base = name === '원형' ? entries.filter(([, p]) => p.type === 'pie') : name === '꺾은선형' ? entries.slice(0, 1) : [...entries];
  for (const [label, props] of base) entries.push([`3차원 ${label}`, { ...props, threeD: true }]);
}

for (const [name, entries] of CHART_GALLERY) {
  if (!['세로 막대형','가로 막대형'].includes(name)) continue;
  const column=name==='세로 막대형';
  for (const [, props] of entries) props.barShape='box';
  if(column)entries.push(['3차원 세로 막대형 (깊이 축)',{type:'column',threeD:true,grouping:'standard',barShape:'box'}]);
  const modes=[['묶은','clustered'],['누적','stacked'],['100% 기준 누적','percentStacked'],...(column?[['깊이 축','standard']]:[])];
  for(const [shape,label] of [['cylinder','원통'],['cone','원뿔'],['pyramid','피라미드']])for(const [prefix,grouping] of modes)entries.push([prefix+' '+label+' '+name,{type:column?'column':'bar',threeD:true,grouping,barShape:shape}]);
}

/** 자동 콤보 기본값. 명시 계열 형식은 호출자가 우선 적용한다. */
export function chartComboDefaults(chart,index,count) {
  if(chart.type!=='combo')return {};
  const last=index===count-1&&count>1;
  if(chart.comboLayout==='areaColumn')return last?{type:'column',axis:0,grouping:'clustered'}:{type:'area',axis:0,grouping:'stacked'};
  if(chart.comboLayout==='stackedColumnLine')return last?{type:'line',axis:chart.comboAxis==='primary'?0:1,grouping:'clustered'}:{type:'column',axis:0,grouping:'stacked'};
  return last?{type:'line',axis:chart.comboAxis==='primary'?0:1}:{type:'column'};
}


export const PALETTE = ['#4472C4', '#ED7D31', '#A5A5A5', '#FFC000', '#5B9BD5', '#70AD47', '#264478', '#9E480E', '#636363', '#997300'];
/** 색 변경 (엑셀 [차트 디자인] → [색 변경]): 통합 문서 테마의 강조 색으로 만듦 (엑셀과 같음) + WIXEL 팔레트 */
const accents = () => THEME.colors.slice(4, 10).map((c) => String(c).toUpperCase());
const hx = (c) => `#${c}`;
// 7번째 계열부터: 강조 색 60% 밝기(lumMod 60000), 13번째부터: 80% + 20%(lumMod 80000 lumOff 20000)
const cycle = (list) => [...list.map(hx), ...list.map((c) => hx(applyTint(c, -0.4))), ...list.map((c) => hx(applyTint(c, 0.2)))];
const pickAcc = (idx) => { const a = accents(); const base = idx.map((k) => a[k]); return [...base.map(hx), ...base.map((c) => hx(applyTint(c, -0.4))), ...base.map((c) => hx(applyTint(c, 0.4)))]; };
// 단색형: 한 강조 색의 어두운 색 → 밝은 색
const mono = (k) => { const c = k < 0 ? '7F7F7F' : accents()[k]; return [-0.5, -0.25, 0, 0.25, 0.5, 0.7].map((t) => hx(applyTint(c, t))); };
const monoRev = (k) => mono(k).reverse();
export const CHART_PALETTES = {
  office: { label: '다양한 색 팔레트 1', get colors() { return cycle(accents()); } },
  colorful2: { label: '다양한 색 팔레트 2', get colors() { return pickAcc([0, 2, 4]); } },
  colorful3: { label: '다양한 색 팔레트 3', get colors() { return pickAcc([1, 3, 5]); } },
  colorful4: { label: '다양한 색 팔레트 4', get colors() { return pickAcc([5, 4, 3]); } },
  mono1: { label: '단색 팔레트 1', mono: true, get colors() { return mono(0); } },
  mono2: { label: '단색 팔레트 2', mono: true, get colors() { return mono(1); } },
  mono3: { label: '단색 팔레트 3', mono: true, get colors() { return mono(2); } },
  mono4: { label: '단색 팔레트 4', mono: true, get colors() { return mono(3); } },
  mono5: { label: '단색 팔레트 5', mono: true, get colors() { return mono(4); } },
  mono6: { label: '단색 팔레트 6', mono: true, get colors() { return mono(5); } },
  mono7: { label: '단색 팔레트 7', mono: true, get colors() { return monoRev(0); } },
  mono8: { label: '단색 팔레트 8', mono: true, get colors() { return monoRev(1); } },
  mono9: { label: '단색 팔레트 9', mono: true, get colors() { return monoRev(2); } },
  mono10: { label: '단색 팔레트 10', mono: true, get colors() { return monoRev(3); } },
  mono11: { label: '단색 팔레트 11', mono: true, get colors() { return monoRev(4); } },
  mono12: { label: '단색 팔레트 12', mono: true, get colors() { return monoRev(5); } },
  mono13: { label: '단색 팔레트 13 (회색)', mono: true, get colors() { return mono(-1); } },
  modern: { group: 'report', wixel: true, label: 'WIXEL 모던', colors: ['#4F46E5', '#0EA5E9', '#10B981', '#F59E0B', '#F43F5E', '#8B5CF6', '#64748B', '#14B8A6'] },
  pastel: { group: 'pastel', wixel: true, label: 'WIXEL 파스텔', colors: ['#818CF8', '#7DD3FC', '#6EE7B7', '#FCD34D', '#FDA4AF', '#C4B5FD', '#CBD5E1', '#5EEAD4'] },
  slate: { group: 'dark', wixel: true, label: 'WIXEL 슬레이트', colors: ['#1E293B', '#475569', '#64748B', '#94A3B8', '#CBD5E1', '#0EA5E9'] },
  vivid: { group: 'vivid', wixel: true, label: 'WIXEL 비비드', colors: ['#2563EB', '#DC2626', '#16A34A', '#D97706', '#9333EA', '#0891B2', '#DB2777', '#65A30D'] },
  ...EXTRA_CHART_PALETTES,
};
export function paletteOf(ch) {
  if (Array.isArray(ch?.palette)) {
    // 기존 파일의 유효한 색·순서·개수는 유지한다. 사용자 입력 개수 제한은 UI에서만 적용한다.
    const colors = ch.palette.map(color => typeof color === 'string' && /^#[0-9a-f]{6}$/i.test(color) ? color : normalizeChartPaletteColor(color)).filter(Boolean);
    if (colors.length) return colors;
  }
  return CHART_PALETTES[ch?.palette]?.colors ?? CHART_PALETTES.office.colors;
}

const isNum = (v) => typeof v === 'number' && Number.isFinite(v);
const label = (v) => (v === null || v === undefined ? '' : typeof v === 'number' ? formatGeneral(v) : String(v?.code ?? v));

/**
 * 2차원 값 → { categories, series:[{name, values}] }
 * 엑셀처럼: 첫 행/첫 열이 숫자가 아니면 머리글로, 행 수 ≥ 열 수이면 열 단위 계열
 */
export function chartLayout(rows, type = 'column', flip = false) {
  const R = rows.length;
  const C = rows[0]?.length ?? 0;
  const headRow = R > 1 && rows[0].slice(C > 1 ? 1 : 0).every((v) => !isNum(v)) && rows[0].some((v) => v !== null && v !== '');
  const firstDataRow = headRow ? 1 : 0;
  const xy = type === 'scatter' || type === 'bubble';
  const body = rows.slice(firstDataRow);
  // 항목 열: 숫자가 없는 열, 또는 글자와 빈 칸만 있는 병합형 묶음 열(바깥 항목) — 앞쪽에 여러 개면 엑셀처럼 다단계 항목
  const isLabel = (c) => body.every((r) => !isNum(r[c]));
  // 병합형 바깥 묶음 열: 글자가 있고 빈 칸이 섞인 첫 열 (2026.04 처럼 숫자 모양 묶음 이름 포함) + 그 옆이 항목 열
  const groupCol0 = C > 2 && isLabel(1) && body.some((r) => typeof r[0] === 'string' && r[0] !== '') && body.some((r) => r[0] === null || r[0] === '');
  let lc = 0;
  if (!xy) while (lc < C - 1 && (isLabel(lc) || (lc === 0 && groupCol0))) lc++;
  const catCol = C > 1 && (xy || lc > 0);
  const firstDataCol = catCol ? (xy ? 1 : lc) : 0;
  const auto = xy || type === 'sunburst' || type === 'stock' || type === 'boxWhisker' || R - firstDataRow >= C - firstDataCol;
  const byCols = flip && !xy ? !auto : auto;
  return { R, C, headRow, catCol, firstDataRow, firstDataCol, byCols };
}

export function chartData(rows, type = 'column', flip = false) {
  if (!rows.length || !rows[0]?.length) return { categories: [], series: [] };
  if (type === 'sunburst' || type === 'treemap') {
    const grid = flip ? rows[0].map((_, c) => rows.map((r) => r[c])) : rows;
    const width = grid[0].length;
    if (width < 2) return { categories: [], series: [] };
    const header = !isNum(grid[0][width - 1]), body = header ? grid.slice(1) : grid;
    const multi = hierarchyCategories(body.map((r) => r.slice(0, width - 1).map(label)), body.length);
    return { categories: multi?.categories ?? body.map((r) => label(r[width - 2])), series: [{ name: header ? label(grid[0][width - 1]) : '값', values: body.map((r) => isNum(r[width - 1]) ? r[width - 1] : null), x: null }], ...(multi ? { catLevels: multi.levels, hierarchyPaths: multi.paths } : {}) };
  }
  if (type === 'surface') {
    const grid = flip ? rows[0].map((_, c) => rows.map((r) => r[c])) : rows;
    const headers = !isNum(grid[0][0]);
    const body = headers ? grid.slice(1) : grid, offset = headers ? 1 : 0;
    return { categories: grid[0].slice(offset).map((v, i) => headers ? label(v) : String(i + 1)), series: body.map((row, i) => ({ name: headers ? label(row[0]) : String(i + 1), values: row.slice(offset).map((v) => isNum(v) ? v : null), x: null })) };
  }
  const { R, C, headRow, catCol, firstDataRow, firstDataCol, byCols } = chartLayout(rows, type, flip);
  const series = [];
  let categories = [];
  let catLevels = null;
  if (byCols) {
    categories = rows.slice(firstDataRow).map((r, i) => (catCol ? label(r[firstDataCol - 1]) : String(i + 1)));
    if (catCol && firstDataCol > 1) {
      const multi = multiLevel(rows.slice(firstDataRow).map((r) => r.slice(0, firstDataCol).map((v) => (v === null ? '' : label(v)))), rows.length - firstDataRow);
      if (multi) catLevels = multi.levels;
    }
    for (let c = firstDataCol; c < C; c++) {
      series.push({
        name: headRow ? label(rows[0][c]) : `계열${c - firstDataCol + 1}`,
        values: rows.slice(firstDataRow).map((r) => (isNum(r[c]) ? r[c] : null)),
        x: (type === 'scatter' || type === 'bubble') && catCol && rows.slice(firstDataRow).some((r) => isNum(r[0])) ? rows.slice(firstDataRow).map((r) => (isNum(r[0]) ? r[0] : null)) : null,
      });
    }
    // 거품형: (Y, 크기) 열 쌍마다 계열 하나
    if (type === 'bubble') {
      const paired = [];
      for (let k = 0; k < series.length; k += 2) paired.push({ ...series[k], size: series[k + 1]?.values ?? series[k].values.map(() => 1) });
      series.splice(0, series.length, ...paired);
    }
  } else {
    categories = [];
    for (let c = firstDataCol; c < C; c++) categories.push(headRow ? label(rows[0][c]) : String(c - firstDataCol + 1));
    for (let r = firstDataRow; r < R; r++) {
      series.push({
        name: catCol ? label(rows[r][0]) : `계열${r - firstDataRow + 1}`,
        values: rows[r].slice(firstDataCol).map((v) => (isNum(v) ? v : null)),
        x: null,
      });
    }
  }
  // 숫자가 하나도 없는 계열(텍스트 열)은 제외
  const numeric = series.filter((sr) => sr.values.some((v) => v !== null));
  return { categories, series: numeric.length ? numeric : series, ...(catLevels ? { catLevels } : {}) };
}

/** 참조 값 → 1차원 (한 열이면 행 순서, 한 행이면 열 순서, 여러 열이면 행마다 글자를 이어 붙임) */
function flatRef(rows, asText) {
  if (!rows?.length) return [];
  if (rows.length === 1) return rows[0];
  if (rows[0].length === 1) return rows.map((r) => r[0]);
  return asText ? rows.map((r) => r.filter((v) => v !== null && v !== '').map(label).join(' / ')) : rows.map((r) => r[r.length - 1]);
}

/**
 * 여러 열(또는 행)로 된 항목 범위 → 엑셀의 다단계 항목 축
 * 안쪽 이름(마지막 열)이 항목, 바깥 열은 값이 있는 칸에서 새 묶음이 시작됨 (병합된 칸은 첫 칸에만 값)
 * 반환: { categories, levels: [[{ text, start, end }]] (안쪽 → 바깥) } | null
 */
function multiLevel(rows, n) {
  if (!rows || rows.length < 2 || rows[0].length < 2) return null;
  // 계열 값 개수와 맞는 쪽이 항목 방향
  let grid = rows;
  if (rows.length !== n && rows[0].length === n) grid = rows[0].map((_, c) => rows.map((r) => r[c]));
  const depth = grid[0].length;
  const txt = (v) => (v === null || v === undefined ? '' : String(v));
  const categories = grid.map((r) => txt(r[depth - 1]));
  const levels = [];
  for (let d = depth - 2; d >= 0; d--) {
    const spans = [];
    grid.forEach((r, i) => {
      const t = txt(r[d]);
      // 더 바깥 단계에서 새 묶음이 시작되면 이 단계도 새로 시작
      const outerBreak = r.slice(0, d).some((v) => txt(v) !== '');
      if (i === 0 || t !== '' || outerBreak) spans.push({ text: t, start: i, end: i });
      else spans[spans.length - 1].end = i;
    });
    levels.push(spans);
  }
  if (!levels.some((l) => l.some((sp) => sp.text !== ''))) return null;
  return { categories, levels };
}

/**
 * 차트 모델 → 그릴 데이터 { categories, series: [{ name, values, x, type, axis, color, labels, marker, numFmt }] }
 * api: { values(ref) → 2차원 값, pivot({ sheet, name }) → { categories, series } | null, range(ch) → 2차원 값 }
 */
export function resolveChart(ch, api) {
  // 선택 시트 게시본은 원본 참조 없이 그 시점의 표시 데이터만 사용합니다.
  if (ch.snapshotData) {
    const data = structuredClone(ch.snapshotData);
    data.series = data.series.map((s, i) => ({ ...s, ...ch.seriesFmt?.[s._fi ?? i], _fi: s._fi ?? i }));
    return filterChart(data, ch);
  }
  let base;
  if (ch.pivot) {
    base = api.pivot(ch.pivot) ?? { categories: [], series: [] };
    // 위셀: 피벗 차트에 보일 지표만 골라서 (엑셀처럼 모든 값 필드가 강제로 나오지 않게) — 고른 순서대로
    const pick = ch.pivot.values;
    if (pick?.length && base.series?.length) {
      const order = new Map(pick.map((n, i) => [n, i]));
      base = { ...base, series: base.series.filter((sr) => order.has(sr.measure)).sort((a, b) => order.get(a.measure) - order.get(b.measure)) };
    }
  }
  else if (ch.series?.length) {
    const catRef = ch.series.find((s) => s.cat)?.cat;
    const series = ch.series.map((s, i) => {
      const valueRows = s.val ? api.values(s.val) : null;
      const vals = valueRows ? flatRef(valueRows, false) : s.cache ?? [];
      const indexes = valueRows?.length > 1 ? api.rowIndexes?.(valueRows) : null;
      // 이름이 여러 칸이면 빈 칸(병합 안쪽)을 빼고 공백으로 이어 붙임 (엑셀과 같음)
      const name = s.name?.ref ? flatRef(api.values(s.name.ref), true).map(label).filter((t) => t !== '').join(' ') : s.name?.text ?? `계열${i + 1}`;
      const xs = s.x ? flatRef(api.values(s.x), false) : s.xCache ?? null;
      const size = (s.size ? flatRef(api.values(s.size), false) : s.sizeCache)?.map((v) => (isNum(v) ? v : null)) ?? null;
      return { name, ...(indexes ? { _pi: indexes } : {}), values: vals.map((v) => (isNum(v) ? v : null)), x: xs ? xs.map((v) => (isNum(v) ? v : null)) : null, ...(size ? { size } : {}) };
    });
    const n = Math.max(0, ...series.map((s) => s.values.length));
    const catRows = catRef ? (api.texts ? api.texts(catRef) : api.values(catRef)?.map((r) => r.map(label))) : null;
    const hierarchy = ch.type === 'sunburst' || ch.type === 'treemap';
    const multi = hierarchy ? hierarchyCategories(catRows, n) : multiLevel(catRows, n);
    const cachedCats = ch.series.find((s) => Array.isArray(s.catCache))?.catCache;
    const cachedLevels = ch.series.find((s) => Array.isArray(s.catLevels))?.catLevels;
    const categories = multi ? multi.categories : catRef ? (api.texts ? flatRef(catRows, true) : flatRef(api.values(catRef), true).map(label)) : cachedCats ? cachedCats.map(label) : Array.from({ length: n }, (_, i) => String(i + 1));
    base = { categories, series, ...(multi ? { catLevels: multi.levels, ...(multi.paths ? { hierarchyPaths: multi.paths } : {}) } : cachedLevels ? { catLevels: structuredClone(cachedLevels) } : {}) };
  } else {
    const rows = api.range(ch);
    const type = ch.type === 'combo' ? 'column' : ch.type;
    base = chartData(rows, type, !!ch.byRows);
    const indexes = api.rowIndexes?.(rows);
    if (indexes) {
      let points, fields;
      if (type === 'sunburst' || type === 'treemap') {
        if (!ch.byRows) { const header = !isNum(rows[0][rows[0].length - 1]); points = indexes.slice(header ? 1 : 0).map(i => i - (header ? 1 : 0)); }
      } else if (type === 'surface') {
        const offset = !isNum(rows[0][0]) ? 1 : 0;
        const ids = indexes.slice(offset).map(i => i - offset);
        if (ch.byRows) points = ids; else fields = ids;
      } else {
        const layout = chartLayout(rows, type, !!ch.byRows);
        const ids = indexes.slice(layout.firstDataRow).map(i => i - layout.firstDataRow);
        if (layout.byCols) points = ids; else fields = ids;
      }
      base.series = base.series.map((sr, i) => ({ ...sr, ...(points ? { _pi: points } : {}), ...(fields ? { _fi: fields[i] } : {}) }));
    }

  }
  const fmt = ch.seriesFmt ?? [];
  const comboDefault = (i) => chartComboDefaults(ch, i, base.series.length);
  // 계열 형식에 종류가 정해져 있으면 축도 그 형식대로 (axis 가 없으면 기본 축) — 파일의 콤보 차트에서 마지막 계열을 보조 축으로 보내지 않게
  base.series = base.series.map((s, i) => {
    const index = s._fi ?? i;
    const out = { ...s, ...(fmt[index]?.type ? { axis: 0 } : comboDefault(index)), ...(fmt[index] ?? {}), _fi: index };
    out.axis = chartAxis(out.axis); return out;
  });
  if (base.series[0]?._pi) base._ci = base.series[0]._pi;
  return filterChart(base, ch);
}

/**
 * 차트 필터 (엑셀 차트 옆 깔때기 단추): 숨긴 계열 · 항목은 그리지 않음
 * ch.hiddenSeries / ch.hiddenCats = 원래 순서의 번호. 계열 색은 숨기기 전 순서대로 유지 (엑셀과 같음)
 */
export function filterChart(base, ch) {
  const hs = ch.hiddenSeries?.length ? new Set(ch.hiddenSeries) : null;
  const hc = ch.hiddenCats?.length ? new Set(ch.hiddenCats) : null;
  if (!hs && !hc) return base;
  const pal = paletteOf(ch);
  let series = base.series.map((s, i) => ({ ...s, _fi: s._fi ?? i, color: s.color ?? pal[(s._fi ?? i) % pal.length] }));
  if (hs) series = series.filter((s) => !hs.has(s._fi));
  const out = { ...base, series };
  if (hc) {
    const keep = (base.categories ?? []).map((_, i) => !hc.has(base._ci?.[i] ?? base.series[0]?._pi?.[i] ?? i));
    const pick = (arr) => (Array.isArray(arr) ? arr.filter((_, i) => keep[i] ?? true) : arr);
    out.categories = pick(base.categories);
    if (base._ci) out._ci = pick(base._ci);
    out.series = series.map((s) => ({ ...s, _pi: pick(s._pi ?? s.values.map((_, i) => i)), values: pick(s.values), x: pick(s.x), ...(s.size ? { size: pick(s.size) } : {}) }));
    if (base.catLevels?.length || base.hierarchyPaths) Object.assign(out, filterHierarchyData(base, keep));
  }
  return out;
}

// 데이터 필터로 줄어든 화면 번호와 원본 계열/요소 서식 번호를 구분합니다.
export function chartPointIndex(series, index) { return series?._pi?.[index] ?? index; }
const pointColor = (s, i) => s?.pointColors?.[chartPointIndex(s, i)] ?? s?.colors?.[chartPointIndex(s, i)];
export function chartPointExplosion(chart, series, index) {
  const amount = series?.pointExplosion?.[chartPointIndex(series, index)] ?? series?.explode ?? chart.explode ?? 0;
  return Number.isFinite(amount) ? Math.max(0, Math.min(400, amount)) : 0;
}
const manualChartLayout = (layout) => Number.isFinite(layout?.x) && Number.isFinite(layout?.y) ? layout : null;

/** 차트 모델 → 그릴 데이터 (범위 · 계열 참조 · 피벗 차트). hostSi: 차트가 있는 시트 */
export function chartModelData(wb, hostSi, ch, options = {}) {
  // 과거 위셀의 피벗 숫자 열 범위 차트도 원본·명시 계열을 바꾸지 않고 범주 참조를 복구한다.
  const inferred = inferPivotCategorySeries(wb, hostSi, ch);
  if (inferred) ch = { ...ch, series: inferred };
  const sheetOf = (name) => { const i = name ? wb.sheetIndexByName(name) : hostSi; return i >= 0 ? i : hostSi; };
  const sampled = options.sample !== false;
  const rowIndexes = new WeakMap();
  const exportRange = (s, rg) => {
    if (sampled || !rg) return rg;
    const fullRows = rg.r2 === 1048575 || rg.r2 === 19999999, fullCols = rg.c2 === 16383;
    if (!fullRows && !fullCols) return rg;
    // 전체 열/행 참조의 빈 꼬리를 순회하지 않는다. 참조 자체는 writer가 그대로 보존한다.
    const used = wb.usedRange(s); let rows = used.rows, cols = used.cols;
    for (const sp of wb.spillsOf(s)) { rows = Math.max(rows, sp.r + sp.h); cols = Math.max(cols, sp.c + sp.w); }
    return { ...rg, r2: fullRows ? Math.min(rg.r2, Math.max(rg.r1, rows - 1)) : rg.r2, c2: fullCols ? Math.min(rg.c2, Math.max(rg.c1, cols - 1)) : rg.c2 };
  };
  const read = (s, inputRange, text = false) => {
    const rg = exportRange(s, inputRange);
    // 화면은 전체 범위의 2,000개 표본, XLSX 저장은 원래 순서의 모든 데이터.
    const total = rg.r2 - rg.r1 + 1, count = sampled ? Math.min(total, 2000) : total;
    const step = sampled && total > 2000 ? (total - 1) / 1999 : 1;
    const rows = [];
    for (let k = 0; k < count; k++) {
      const r = rg.r1 + Math.round(k * step);
      const row = [];
      for (let c = rg.c1; c <= (sampled ? Math.min(rg.c2, rg.c1 + 100) : rg.c2); c++) {
        const v = wb.getValue(s, r, c);
        if (!text || text === 'category' && c !== rg.c1) { row.push(v); continue; }
        // 새 차트의 첫 날짜 열은 항목이며, 다른 숫자 값과 분산/거품 X 값은 원래 숫자입니다.
        const st = typeof v === 'number' ? wb.styleAt(s, r, c) : null;
        if (text === 'category' && !(/date|time/.test(st?.numFmt ?? '') || st?.code && isDateCode(st.code))) { row.push(v); continue; }
        row.push(st && (st.numFmt && st.numFmt !== 'general') ? formatValue(v, st, wb.date1904).text : label(v));
      }
      rows.push(row);
    }
    if (step !== 1) rowIndexes.set(rows, Array.from({ length: count }, (_, k) => Math.round(k * step)));
    return rows;
  };
  // 정의된 이름 참조 (OFFSET 등으로 바뀌는 범위): 그릴 때마다 이름을 계산
  const nameRange = (ref) => {
    let v;
    try { v = wb.nameValue(ref.name, hostSi, ref.sheet ?? null, null); } catch { v = null; }
    if (!v || v.r1 === undefined) return null;
    return { sheet: v.sheet ?? ref.sheet ?? null, r1: v.r1, c1: v.c1, r2: v.r2, c2: v.c2 };
  };
  if (!sampled && options.onRange && ch.snapshotData && ch.range && !ch.series?.length && !ch.pivot) {
    options.onRange(read(sheetOf(ch.sheet), ch.range, ch.range.c2 > ch.range.c1 && !['scatter', 'bubble'].includes(ch.type) ? 'category' : false));
  }
  const data = resolveChart(ch, {
    rowIndexes: rows => rowIndexes.get(rows),
    range: (c) => {
      if (!c.range) return [];
      const rows = read(sheetOf(c.sheet), c.range, c.range.c2 > c.range.c1 && !['scatter', 'bubble'].includes(c.type) ? 'category' : false);
      options.onRange?.(rows);
      return rows;
    },
    values: (ref) => {
      if (ref.name) { const rg = nameRange(ref); return rg ? read(sheetOf(rg.sheet), rg) : []; }
      return read(sheetOf(ref.sheet ?? ch.sheet), ref);
    },
    texts: (ref) => {
      if (ref.name) { const rg = nameRange(ref); return rg ? read(sheetOf(rg.sheet), rg, true) : []; }
      return read(sheetOf(ref.sheet ?? ch.sheet), ref, true);
    },
    pivot: (p) => {
      const sh = wb.sheets[sheetOf(p.sheet)];
      const defs = [sh?.pivot, ...(sh?.pivotsExtra ?? [])].filter(Boolean);
      const def = defs.find((d) => d.name && d.name === p.name) ?? (p.name ? null : defs[0]);
      if (!def) return null;
      const src = pivotSourceData(wb, def);
      if (!src) return null;
      // 날짜 등 숫자 항목은 원본 열의 표시 형식으로
      const header = src.cube.header.map((h) => String(h ?? '').toLowerCase());
      const fieldStyle = (f) => {
        const i = header.indexOf(String(f).toLowerCase());
        return i >= 0 && src.ref && src.si !== undefined ? wb.styleAt(src.si, Math.min(src.ref.r1 + 1, src.ref.r2), src.ref.c1 + i) : null;
      };
      return pivotChartData(src, def, fieldStyle);
    },
  });
  return { ...data, date1904: !!wb.date1904 };
}

const chartAxis = (axis) => axis === 1 || axis === 'secondary' || axis === 'right' ? 1 : 0;

/** 축·종류별 누적 값. 눈금 계산과 도형 배치가 같은 합계/분모를 사용한다. */
export function chartStackValues(series, grouping, n) {
  const groups = new Map();
  const out = series.map((s) => ({ ...s, axis: chartAxis(s.axis), _lower: new Array(n).fill(0), _upper: [...s.values] }));
  for (const s of out) {
    const mode = s.grouping ?? grouping;
    if (!['stacked','percentStacked'].includes(mode) || !['column', 'bar', 'line', 'area'].includes(s.type)) continue;
    const key = `${s.axis}/${s.type}/${mode}`; s._stackGroup = key; s._pct = mode === 'percentStacked';
    if (!groups.has(key)) groups.set(key, []);
    groups.get(key).push(s);
  }
  for (const group of groups.values()) {
    const pct = group[0]._pct;
    const totals = pct ? Array.from({ length: n }, (_, i) => group.reduce((sum, s) => sum + (isNum(s.values[i]) ? Math.abs(s.values[i]) : 0), 0) || 1) : null;
    const pos = new Array(n).fill(0), neg = new Array(n).fill(0), signed = new Array(n).fill(0);
    for (const s of group) {
      const bar = s.type === 'column' || s.type === 'bar';
      s._stacked = true; s._pct = pct;
      s.values.forEach((raw, i) => {
        if (!isNum(raw)) return;
        const v = pct ? raw / totals[i] : raw;
        const acc = bar ? (v < 0 ? neg : pos) : signed;
        s._lower[i] = acc[i]; acc[i] += v; s._upper[i] = acc[i];
      });
    }
  }
  return out;
}

/** 보기 좋은 눈금 */
export function niceScale(min, max, ticks = 5) {
  if (min === max) {
    if (min === 0) return { min: 0, max: 1, step: 0.2 };
    const d = Math.abs(min) * 0.5;
    min -= d; max += d;
  }
  const range = max - min;
  const rough = range / ticks;
  const mag = 10 ** Math.floor(Math.log10(rough));
  const norm = rough / mag;
  const step = (norm <= 1 ? 1 : norm <= 2 ? 2 : norm <= 2.5 ? 2.5 : norm <= 5 ? 5 : 10) * mag;
  const lo = Math.floor(min / step) * step;
  const hi = Math.ceil(max / step) * step;
  return { min: lo, max: hi === lo ? lo + step : hi, step };
}

const escSvg = (s) => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');

/** 데이터 레이블의 표시 항목. 값·원본 계열명·범주명을 조합하며 원본은 바꾸지 않는다. */
export function chartDataLabel(chart, series, categories, index, value) {
  const s = series ?? {}, out = [];
  if (s.serName && s.name != null) out.push(String(s.name));
  if (s.catName && index !== undefined && categories[index] != null) out.push(String(categories[index]));
  if (s.labels ?? chart.labels ?? (!s.serName && !s.catName && !s.pct)) out.push(valueLabel(value, s.numFmt, chart.date1904));
  if (s.pct && isNum(value)) {
    let total = 0; for (const v of s.values ?? []) if (isNum(v)) total += Math.abs(v);
    const ratio = s._pct && index !== undefined ? s._upper[index] - s._lower[index] : total ? value / total : 0;
    out.push(formatCode(ratio, s.percentFmt ?? '0%').text);
  }
  return out.join(s.labelSeparator ?? ', ');
}

function axisLabel(n, code, date1904 = false) {
  if (code && code !== 'General') { try { return formatCode(n, code, date1904).text; } catch { /* 기본 */ } }
  const a = Math.abs(n);
  if (a >= 1e6 && a < 1e15 && !code) return `${formatGeneral(Number((n / 1e6).toPrecision(4)))}M`;
  if (Number.isInteger(n)) return n.toLocaleString('en-US');
  return formatGeneral(Number(n.toPrecision(6)));
}

/** 데이터 레이블 글자 (계열 서식: 서식 코드 문자열 또는 셀 서식) */
function valueLabel(v, fmt, date1904 = false) {
  try {
    if (typeof fmt === 'string' && fmt && fmt !== 'General') return formatCode(v, fmt, date1904).text;
    if (fmt && typeof fmt === 'object' && (fmt.numFmt || fmt.code)) return formatValue(v, fmt, date1904).text;
  } catch { /* 기본 */ }
  if (Number.isInteger(v)) return v.toLocaleString('en-US');
  return formatGeneral(Number(v.toPrecision(6)));
}

function truncate(s, maxChars) {
  return [...s].length > maxChars ? `${[...s].slice(0, Math.max(1, maxChars - 1)).join('')}…` : s;
}

const MARKERS = {
  circle: (x, y, r, c) => `<circle cx="${x}" cy="${y}" r="${r}" fill="${c}"/>`,
  square: (x, y, r, c) => `<rect x="${x - r}" y="${y - r}" width="${2 * r}" height="${2 * r}" fill="${c}"/>`,
  diamond: (x, y, r, c) => `<path d="M${x},${y - r - 1}L${x + r + 1},${y}L${x},${y + r + 1}L${x - r - 1},${y}Z" fill="${c}"/>`,
  triangle: (x, y, r, c) => `<path d="M${x},${y - r - 1}L${x + r + 1},${y + r}L${x - r - 1},${y + r}Z" fill="${c}"/>`,
};

/**
 * chart: 차트 모델 (+ w, h), data: resolveChart 결과 → SVG 문자열
 */
/** 색 밝기 (0 검정 ~ 1 흰색) */
function lum(c) {
  let h = String(c ?? '').replace('#', '');
  if (/^[0-9a-f]{3}$/i.test(h)) h = [...h].map((v) => v + v).join('');
  if (!/^[0-9a-f]{6}$/i.test(h)) return 0.5;
  const [r, g, b] = [0, 2, 4].map((k) => parseInt(h.slice(k, k + 2), 16) / 255);
  return 0.2126 * r + 0.7152 * g + 0.0722 * b;
}
/** 글자 색이 배경과 거의 같으면 보이는 색으로 (흰 바탕 흰 글자 → 회색) */
function readable(fg, bg) {
  return Math.abs(lum(fg) - lum(bg)) < 0.25 ? (lum(bg) > 0.5 ? '#595959' : '#f2f2f2') : fg;
}
function plotCategoryLabels(categories, area, plot, font, cfg, width, slotCount = categories.length) {
  const options = { slotCount, length: area.w, depth: Math.max(font * 1.9, Math.min(160, plot.h * .44)), font,
    interval: cfg.labelInterval, rotation: cfg.labelRotation, reverse: cfg.reverse,
    startRoom: area.x - plot.x, endRoom: width - 10 - area.x - area.w, allowPadding: cfg.labelRotation === undefined };
  let plan = categoryAxisLayout(categories, options);
  if (plan.startPadding > .1 || plan.endPadding > .1) {
    const left = Math.ceil(plan.startPadding * 1.2 + 2), right = Math.ceil(plan.endPadding * 1.2 + 2);
    if (left + right < area.w * .3) {
      const next = { ...area, x: area.x + left, w: area.w - left - right };
      const fitted = categoryAxisLayout(categories, { ...options, length: next.w, startRoom: next.x - plot.x, endRoom: width - 10 - next.x - next.w, allowPadding: false });
      if (fitted.angle === plan.angle && fitted.fits) { Object.assign(area, next); plan = fitted; }
      else plan = categoryAxisLayout(categories, { ...options, allowPadding: false });
    }
  }
  return plan;
}
let svgSeq = 0;
export function renderChartSvg(chart, data) {
  chart = { ...chart, date1904: data.date1904 ?? chart.date1904 ?? false };
  const W = Math.max(120, chart.w);
  const H = Math.max(90, chart.h);
  const FONT = "font-family=\"'Malgun Gothic','Apple SD Gothic Neo','Noto Sans KR',sans-serif\"";
  // 글꼴 크기(px): 엑셀 기본값 — 제목 14pt, 축 · 범례 · 레이블 9pt (파일에 적힌 크기가 있으면 그것)
  const pt = (v) => Math.round(v * (4 / 3) * 10) / 10;
  const FS = { title: pt(chart.titleSize ?? 14), axis: pt(chart.axisSize ?? 9), legend: pt(chart.legendSize ?? 9) };
  const CW = FS.axis * 0.58; // 글자 하나의 대략 너비
  const LW = FS.legend * 0.58;
  const uid = `c${(++svgSeq).toString(36)}`;
  const defs = [];
  const pal = paletteOf(chart);
  const TXT = chart.textColor ?? '#595959';
  const GRID = chart.gridColor ?? '#d9d9d9';
  const overlays = [];
  const parts = [`<svg xmlns="http://www.w3.org/2000/svg" width="${W}" height="${H}" viewBox="0 0 ${W} ${H}" ${FONT}>`, '',
    chart.chartAreaFormat ? chartAreaSvg(chartAreaFormat(chart), { x: .5, y: .5, w: W - 1, h: H - 1 }, { id: uid + 'area', rounded: chart.rounded }) : `<rect x="0.5" y="0.5" width="${W - 1}" height="${H - 1}" rx="${chart.rounded ? 8 : 0}" fill="${chart.fill ?? '#fff'}"${chart.border ? ` stroke="${chart.border}"` : ''}/>`];
  let top = 10;
  if (chart.title) {
    const manual = manualChartLayout(chart.titleLayout);
    const tw = manual && Number.isFinite(manual.w) && manual.w > 0 ? manual.w * W : Math.min(W, Math.max(FS.title, [...String(chart.title)].length * FS.title * .9));
    const title = `<text x="${manual ? 0 : W / 2}" y="${manual ? 0 : Math.round(FS.title + 10)}"${manual ? ' dominant-baseline="text-before-edge"' : ' text-anchor="middle"'} font-size="${FS.title}"${chart.titleBold ? ' font-weight="700"' : ''} fill="${chart.titleColor ?? TXT}">${escSvg(truncate(chart.title, Math.floor((manual ? tw : W) / (FS.title * 0.62))))}</text>`;
    overlays.push(`<g data-el="title"${manual ? ` transform="translate(${manual.x * W},${manual.y * H})"` : ''}>${manual ? `<rect x="0" y="0" width="${tw}" height="${Math.max(FS.title * 1.4, Number.isFinite(manual.h) ? manual.h * H : 0)}" fill="transparent"/>` : ''}${title}</g>`);
    if (!chart.titleOverlay && !manual) top = Math.round(FS.title * 1.5 + 16); // 수동 위치는 그림 영역 안쪽으로도 이동할 수 있음
  }
  // 계열 채우기: 단색 또는 그라데이션 (url), 그림자 필터
  const fillOf = (s) => {
    if (!s.grad) return s.color;
    const id = `${uid}g${defs.length}`;
    const a = ((s.grad.ang ?? 90) * Math.PI) / 180;
    const x = Math.cos(a) / 2;
    const y = Math.sin(a) / 2;
    defs.push(`<linearGradient id="${id}" x1="${(0.5 - x).toFixed(3)}" y1="${(0.5 - y).toFixed(3)}" x2="${(0.5 + x).toFixed(3)}" y2="${(0.5 + y).toFixed(3)}">${s.grad.stops.map(([o, c]) => `<stop offset="${o}" stop-color="${c}"/>`).join('')}</linearGradient>`);
    return `url(#${id})`;
  };
  let shadowId = null;
  const shadowAttr = (s) => {
    if (!s.shadow) return '';
    if (!shadowId) { shadowId = `${uid}s`; defs.push(`<filter id="${shadowId}" x="-20%" y="-20%" width="140%" height="140%"><feDropShadow dx="0" dy="1.5" stdDeviation="2" flood-color="#000" flood-opacity="0.45"/></filter>`); }
    return ` filter="url(#${shadowId})"`;
  };
  const axisGroups = { x: [], y: [], y2: [] };
  const axisText = (axis, html) => { axisGroups[axis].push(html); return ''; };
  const finish = () => {
    if (defs.length) parts[1] = `<defs>${defs.join('')}</defs>`;
    for (const axis of ['x', 'y', 'y2']) if (axisGroups[axis].length) parts.push(`<g data-el="axis-${axis}">${axisGroups[axis].join('')}</g>`);
    pushAll(parts, overlays); parts.push('</svg>'); return parts.join('');
  };
  const { categories } = data;
  const baseType = chart.type === 'combo' ? 'column' : chart.type;
  // 다단계 항목 축 (세로 막대 · 꺾은선 · 콤보): 안쪽 이름 아래에 바깥 묶음 이름 줄
  const catLevels = baseType !== 'bar' && baseType !== 'scatter' && !chart.axes?.x?.hide ? data.catLevels ?? [] : [];
  const LEVEL_H = Math.round(FS.axis * 1.6);
  const series = chartStackValues(data.series.map((s, i) => { const defaults = chartComboDefaults(chart, i, data.series.length); return { ...defaults, ...s, type: s.type ?? defaults.type ?? baseType, axis: s.axis ?? defaults.axis ?? 0, color: s.color ?? pal[(s._fi ?? i) % pal.length], _fi: s._fi ?? i }; }), chart.grouping, categories.length);
  // 요소 고르기용 표시 (엑셀처럼 한 번 누르면 계열, 한 번 더 누르면 요소) — data-s = 계열 서식 번호, data-p = 항목 번호
  const tag = (s, p) => ` data-s="${s._fi}"${p === undefined ? '' : ` data-p="${chartPointIndex(s, p)}"`}`;
  const tagMk = (html, s, p) => html.replace(/^<(\w+)/, `<$1${tag(s, p)}`);
  const pieLike = baseType === 'pie' || baseType === 'doughnut';
  const pieColor = (s, i) => pointColor(s, i) ?? chart.seriesFmt?.[s?._fi ?? 0]?.color ?? pal[chartPointIndex(s, i) % pal.length];
  const special = SPECIAL[baseType] ?? ADVANCED_CHARTS[baseType] ?? MAP_CHARTS[baseType];
  const legendPos = chart.legend ?? (special?.legend === false ? 'none' : 'b');
  const legendItems = special?.legendItems ? special.legendItems(data, chart, pal) : baseType === 'treemap' && data.catLevels?.length ? Array.from(new Set(data.catLevels.at(-1).map((g) => g.text))).map((name, i) => ({ name, color: pal[i % pal.length], line: false }))
    : pieLike || baseType === 'treemap' || baseType === 'pieOfPie' || baseType === 'barOfPie' ? categories.map((c, i) => ({ name: c, color: pieColor(series[0], i), line: false }))
    : baseType === 'waterfall' ? [{ name: '증가', color: chart.upColor ?? pal[0] }, { name: '감소', color: chart.downColor ?? pal[1 % pal.length] }, { name: '합계', color: chart.totalColor ?? pal[2 % pal.length] }]
      : baseType === 'pareto' ? [{ name: series[0]?.name ?? '', color: series[0]?.color }, { name: '누적 %', color: pal[1 % pal.length], line: true }]
        : series.map((s) => ({ name: s.name, color: s.color, line: s.type === 'line' || s.type === 'radar' }));
  const showLegend = legendPos !== 'none' && legendItems.length > 0;
  const legendLayout = showLegend ? manualChartLayout(chart.legendLayout) : null;
  const sideLegend = showLegend && (legendPos === 'r' || legendPos === 'l');
  const legendW = sideLegend ? Math.min(180, Math.max(60, maxOf(legendItems.map((it) => [...String(it.name)].length * LW * 1.4 + 22)))) : 0;
  const legendH = showLegend && !sideLegend ? Math.round(FS.legend * 2.2) : 0;
  const plot = {
    x: 10 + (!legendLayout && legendPos === 'l' ? legendW : 0), y: top + (!legendLayout && legendPos === 't' ? legendH : 0),
    w: W - 20 - (legendLayout ? 0 : legendW), h: H - top - 10 - (legendLayout ? 0 : legendH),
  };
  // 데이터 표 (엑셀 [차트 요소 → 데이터 표]): 그림 영역 아래에 항목 × 계열 값 표 (범례 표지 포함)
  const DT_ROW = Math.round(FS.axis * 1.7);
  const specialTable = chart.dataTable && special ? specialDataTable(chart, series, categories) : null;
  const specialTableH = specialTable ? Math.min(DT_ROW * specialTable.length, Math.max(0, plot.h * .45)) : 0;
  const hasTable = !!chart.dataTable && !pieLike && !special && baseType !== 'scatter' && baseType !== 'bar' && series.length > 0;
  const dtH = hasTable ? DT_ROW * (series.length + 1) : 0;
  plot.h -= Math.max(0, dtH - Math.round(FS.axis * 1.9));
  if (specialTable) plot.h -= specialTableH + 6;

  if (!series.length || series.every((s) => s.values.every((v) => v === null))) {
    if (chart.plotAreaFormat) parts.push(chartAreaSvg(chartAreaFormat(chart, 'plot'), plot, { id: uid + 'plot', kind: 'plot' }));
    parts.push(`<text x="${W / 2}" y="${H / 2}" text-anchor="middle" font-size="12" fill="#999">표시할 숫자 데이터가 없습니다</text>`);
    return finish();
  }

  if (showLegend) {
    const legendStart = parts.length;
    const items = legendItems.slice(0, 16);
    // 범례 글자: 범례 색(legendColor) → 차트 글자 색. 배경과 거의 같은 색이면(흰 바탕에 흰 글자) 보이도록 회색
    const LTX = readable(chart.legendColor ?? TXT, chart.fill ?? '#ffffff');
    const light = (c) => lum(c) > 0.92;
    const key = (it, lx, ly) => (it.line
      ? `<line x1="${lx - 2}" y1="${ly - 4}" x2="${lx + 11}" y2="${ly - 4}" stroke="${it.color}" stroke-width="2.25"/><circle cx="${lx + 4.5}" cy="${ly - 4}" r="2.5" fill="${it.color}"/>`
      : `<rect x="${lx}" y="${ly - 8}" width="9" height="9" fill="${it.color}"${light(it.color) ? ' stroke="#bfbfbf" stroke-width="0.75"' : ''}/>`);
    const legendText = (name, maxChars) => {
      const full = String(name), shown = truncate(full, maxChars);
      return `${shown !== full ? `<title>${escSvg(full)}</title>` : ''}${escSvg(shown)}`;
    };
    parts.push(`<g data-el="legend"${legendLayout ? ` transform="translate(${legendLayout.x * W},${legendLayout.y * H})"` : ''}>`);
    if (legendLayout) {
      const rowH = Math.round(FS.legend * 1.6);
      const boxW = Number.isFinite(legendLayout.w) && legendLayout.w > 0 ? Math.max(24, legendLayout.w * W) : sideLegend ? legendW : Math.min(W - 20, 160 * items.length);
      const columns = sideLegend ? 1 : Math.max(1, Math.min(items.length, Math.floor(boxW / 80)));
      const itemW = boxW / columns;
      const boxH = Math.max(rowH * Math.ceil(items.length / columns), Number.isFinite(legendLayout.h) && legendLayout.h > 0 ? legendLayout.h * H : 0);
      parts.push(`<rect x="0" y="0" width="${boxW}" height="${boxH}" fill="transparent"/>`);
      items.forEach((it, i) => {
        const x = (i % columns) * itemW, y = Math.floor(i / columns) * rowH;
        parts.push(key(it, x + 2, y + 11), `<text x="${x + 17}" y="${y}" dominant-baseline="text-before-edge" font-size="${FS.legend}" fill="${LTX}"${chart.legendBold ? ' font-weight="700"' : ''}>${legendText(it.name, Math.max(1, Math.floor((itemW - 20) / (LW * 1.4))))}</text>`);
      });
    } else if (sideLegend) {
      const lx0 = legendPos === 'r' ? W - legendW - 4 : 8;
      let ly = Math.max(top + 12, H / 2 - (items.length * Math.round(FS.legend * 1.6)) / 2);
      items.forEach((it) => {
        parts.push(key(it, lx0, ly), `<text x="${lx0 + 15}" y="${ly}" font-size="${FS.legend}" fill="${LTX}"${chart.legendBold ? ' font-weight="700"' : ''}>${legendText(it.name, Math.floor((legendW - 20) / (LW * 1.4)))}</text>`);
        ly += Math.round(FS.legend * 1.6);
      });
    } else {
      const itemW = Math.min(160, Math.max(60, (W - 20) / items.length));
      const total = itemW * items.length;
      let lx = Math.max(10, (W - total) / 2);
      const ly = legendPos === 't' ? top + 12 : H - 14;
      items.forEach((it) => {
        parts.push(key(it, lx, ly), `<text x="${lx + 15}" y="${ly}" font-size="${FS.legend}" fill="${LTX}"${chart.legendBold ? ' font-weight="700"' : ''}>${legendText(it.name, Math.floor((itemW - 18) / (LW * 1.4)))}</text>`);
        lx += itemW;
      });
    }
    parts.push('</g>');
    pushAll(overlays, parts.splice(legendStart));
  }

  const wantLabels = (s) => (s.labels ?? chart.labels ?? false) || s.catName || s.serName || s.pct;
  const plotBackgroundIndex = parts.length;
  parts.push(chart.plotAreaFormat ? chartAreaSvg(chartAreaFormat(chart, 'plot'), plot, { id: uid + 'plot', kind: 'plot' }) : `<rect data-el="plot" x="${plot.x}" y="${plot.y}" width="${plot.w}" height="${plot.h}" fill="${chart.plotFill ?? 'transparent'}"/>`);
  if (special) {
    const ctx = { chart, data, series, categories, plot, parts, FS, TXT, GRID, pal, defs, uid, wantLabels, axisText, W, H };
    ctx.cartesian = (values, options) => cartesian(ctx, values, options);
    special.draw(ctx);
    if (specialTable) drawSpecialTable(ctx, specialTable, specialTableH);
    return finish();
  }

  if (pieLike) {
    const rings = baseType === 'doughnut' ? series : [series[0]];
    let maxExplosion = 0;
    for (const s of rings) for (let i = 0; i < s.values.length; i++) if (s.values[i] > 0) maxExplosion = Math.max(maxExplosion, chartPointExplosion(chart, s, i) / 100);
    for (const [ringIndex, s0] of rings.entries()) {
    const vals = s0.values.map((v) => (v && v > 0 ? v : 0));
    const sum = vals.reduce((a, b) => a + b, 0);
    const cx = plot.x + plot.w / 2;
    let cy = plot.y + plot.h / 2;
    // 항목 이름 레이블(엑셀의 원형 바깥 레이블)이 있으면 원을 줄여 둘레에 글자 자리를 둠
    const outsideLabel = (s) => ['out', 'outEnd'].includes(s.labelPos) || !s.labelPos && s.catName;
    const outside = ringIndex === rings.length - 1 && outsideLabel(s0);
    const outerLabels = outsideLabel(rings.at(-1));
    const radius = Math.max(2, (Math.min(plot.w, plot.h) / 2 - 6) / (1 + maxExplosion) * (outerLabels ? 0.72 : 1));
    const hole = baseType === 'doughnut' ? radius * Math.max(0.1, Math.min(0.9, (chart.hole ?? 50) / 100)) : 0;
    const ringWidth = (radius - hole) / rings.length;
    const r = baseType === 'doughnut' ? hole + ringWidth * (ringIndex + 1) : radius;
    const inner = baseType === 'doughnut' ? hole + ringWidth * ringIndex : 0;
    const threePie = chart.threeD && baseType === 'pie';
    const projection = threePie ? pieProjection3D(chart, r) : { squash: 1, depth: 0, rotation: 0 };
    cy -= projection.depth / 2;
    const pieTag = (i, angle) => `${tag(s0, i)} data-pie-cx="${cx}" data-pie-cy="${cy}" data-pie-r="${r}" data-pie-angle="${angle}" data-pie-squash="${projection.squash}"`;
    let a = -Math.PI / 2 + ((chart.firstAngle ?? 0) * Math.PI) / 180 + projection.rotation;
    if (threePie && sum) {
      let angle = a;
      const slices = [];
      vals.forEach((v, i) => {
        if (!v) return;
        const end = angle + v / sum * Math.PI * 2, mid = (angle + end) / 2;
        const ex = chartPointExplosion(chart, s0, i) / 100;
        slices.push({ a: angle, b: end, ox: Math.cos(mid) * r * ex, oy: Math.sin(mid) * r * ex * projection.squash, color: pieColor(s0, i), attrs: pieTag(i, mid) });
        angle = end;
      });
      parts.push(pieSolid3D(slices, cx, cy, r, projection));
    }
    vals.forEach((v, i) => {
      if (!v) return;
      const frac = v / sum;
      const a2 = a + frac * Math.PI * 2;
      const color = pieColor(s0, i);
      const mid0 = (a + a2) / 2;
      const ex = chartPointExplosion(chart, s0, i) / 100;
      const ox = ex ? Math.cos(mid0) * r * ex : 0;
      const oy = ex ? Math.sin(mid0) * r * ex * projection.squash : 0;
      if (!threePie && ex) parts.push(`<g transform="translate(${ox.toFixed(2)},${oy.toFixed(2)})">`);
      if (!threePie && frac >= 0.9999) {
        if (inner) parts.push(`<path d="M${cx + r},${cy}A${r},${r} 0 1 1 ${cx - r},${cy}A${r},${r} 0 1 1 ${cx + r},${cy}ZM${cx + inner},${cy}A${inner},${inner} 0 1 1 ${cx - inner},${cy}A${inner},${inner} 0 1 1 ${cx + inner},${cy}Z" fill="${color}" fill-rule="evenodd" stroke="${s0.outline ?? chart.fill ?? '#fff'}"${pieTag(i, mid0)}/>`);
        else parts.push(`<circle cx="${cx}" cy="${cy}" r="${r}" fill="${color}" stroke="${s0.outline ?? '#fff'}"${pieTag(i, mid0)}/>`);
      } else if (!threePie) {
        const large = a2 - a > Math.PI ? 1 : 0;
        const p = (ang, rad) => `${(cx + Math.cos(ang) * rad).toFixed(2)},${(cy + Math.sin(ang) * rad).toFixed(2)}`;
        const d = inner
          ? `M${p(a, r)}A${r},${r} 0 ${large} 1 ${p(a2, r)}L${p(a2, inner)}A${inner},${inner} 0 ${large} 0 ${p(a, inner)}Z`
          : `M${cx},${cy}L${p(a, r)}A${r},${r} 0 ${large} 1 ${p(a2, r)}Z`;
        parts.push(`<path d="${d}" fill="${color}" stroke="${s0.outline ?? chart.fill ?? '#fff'}" stroke-width="1.5"${pieTag(i, mid0)}/>`);
      }
      if (!threePie && ex) parts.push('</g>');
      // 레이블: false = 없음(파일에 없던 원형), pct = 백분율, labels = 값, 정하지 않음 = 백분율
      const pieLabelParts = [];
      if (s0.serName) pieLabelParts.push(String(s0.name ?? ''));
      if (s0.catName) pieLabelParts.push(String(categories[i] ?? ''));
      if (s0.labels ?? chart.labels ?? false) pieLabelParts.push(valueLabel(v, s0.numFmt, chart.date1904));
      if (s0.pct || s0.labels === undefined && chart.labels === undefined && !s0.catName && !s0.serName) pieLabelParts.push(formatCode(frac, s0.percentFmt ?? '0%').text);
      if (outside && pieLabelParts.length) {
        // 바깥 레이블: 항목 이름 + (백분율 · 값) 두 줄, 조각 색 글자
        const mid = (a + a2) / 2;
        const lr = r * 1.2 + 6;
        const lx = cx + ox + Math.cos(mid) * lr;
        const ly = cy + oy + Math.sin(mid) * lr * projection.squash;
        const anchor = Math.cos(mid) > 0.25 ? 'start' : Math.cos(mid) < -0.25 ? 'end' : 'middle';
        const fsz = s0.labelSize ? s0.labelSize * (4 / 3) : 11;
        const tc = s0.labelColor ?? color;
        pieLabelParts.forEach((t, k) => parts.push(`<text data-el="label"${tag(s0, i)} x="${lx.toFixed(1)}" y="${(ly + 4 + (k - (pieLabelParts.length - 1) / 2) * (fsz + 2)).toFixed(1)}" text-anchor="${anchor}" font-size="${fsz}" fill="${tc}" font-weight="${s0.labelBold ? 700 : 400}">${escSvg(t)}</text>`));
      } else if (frac >= 0.04 && pieLabelParts.length) {
        const mid = (a + a2) / 2;
        const lr = inner ? (r + inner) / 2 : r * 0.65;
        const txt = pieLabelParts.join(s0.labelSeparator ?? ', ');
        parts.push(`<text data-el="label"${tag(s0, i)} x="${(cx + ox + Math.cos(mid) * lr).toFixed(1)}" y="${(cy + oy + Math.sin(mid) * lr * projection.squash + 4).toFixed(1)}" text-anchor="middle" font-size="${pt(s0.labelSize ?? 8.25)}" fill="${s0.labelColor ?? '#fff'}" font-weight="${s0.labelBold === false ? 400 : 700}">${escSvg(txt)}</text>`);
      }
      a = a2;
    });
    }
    return finish();
  }

  const horizontal = baseType === 'bar';
  const n = categories.length;
  const trends = new Map(series.map(s => [s, baseType === 'scatter' ? [] : categoryTrendPoints(s)]));
  // Forecast periods reserve room on the category axis as well as the value axis.
  // Stacked/percentage series retain their existing scale semantics.
  let categorySlots = n;
  for (const s of series) if (!s._stacked && !chart.threeD) for (const [x] of trends.get(s)) categorySlots = Math.max(categorySlots, x + 1);
  const categoryIndex = (i) => chart.axes?.x?.reverse ? categorySlots - 1 - i : i;
  const barTypes = new Set(['column', 'bar']);

  // 축별 값 범위 (누적이면 합계 기준)
  const scaleFor = (axis) => {
    const ss = series.filter((s) => s.axis === axis && (s.type !== 'scatter' || baseType === 'scatter'));
    if (!ss.length) return null;
    let vals = ss.flatMap((s) => s._stacked ? [...s._lower, ...s._upper] : s.values).filter(isNum);
    if (ss.some((s) => s._pct)) {
      vals.push(ss.some((s) => s._pct && s.values.some((v) => isNum(v) && v < 0)) ? -1 : 0,
        ss.some((s) => s._pct && s.values.some((v) => isNum(v) && v > 0)) ? 1 : 0);
    }
    for (const s of ss) if (!s._stacked && !chart.threeD) for (const [, y] of trends.get(s)) vals.push(y);
    if (!vals.length) vals = [0, 1];
    const cfg = chart.axes?.[axis ? 'y2' : 'y'] ?? {};
    let min = Math.min(0, minOf(vals));
    let max = Math.max(0, maxOf(vals));
    if (ss.every((s) => s.type === 'line' && !s._stacked)) {
      min = minOf(vals);
      max = maxOf(vals);
      if (min > 0 && min < max * 0.5) min = 0;
    }
    const sc = niceScale(isNum(cfg.min) ? cfg.min : min, isNum(cfg.max) ? cfg.max : max);
    if (isNum(cfg.min)) sc.min = cfg.min;
    if (isNum(cfg.max)) sc.max = cfg.max;
    if (!(sc.max > sc.min)) Object.assign(sc, niceScale(min, max)); // 같은 값·뒤집힌 수동 범위는 0으로 나누지 않음
    if (isNum(cfg.major) && cfg.major > 0 && (sc.max - sc.min) / cfg.major <= 200) sc.step = cfg.major;
    const ticks = [];
    for (let t = sc.min; t <= sc.max + sc.step / 2; t += sc.step) ticks.push(Number(t.toPrecision(12)));
    const srcCode = ss.map((s) => s.numFmt).find((f) => typeof f === 'string' && f !== 'General') ?? null;
    return { ...sc, ticks, reverse: !!cfg.reverse, code: ss.every((s) => s._pct) ? '0%' : cfg.numFmt ?? srcCode };
  };
  const scale = scaleFor(0) ?? scaleFor(1);
  const scale2 = series.some((s) => s.axis === 1) && series.some((s) => s.axis === 0) ? scaleFor(1) : null;
  const primaryAxis = scaleFor(0) ? 0 : 1;
  const hideY = !!chart.axes?.[primaryAxis ? 'y2' : 'y']?.hide;
  const hideX = !!chart.axes?.x?.hide;
  const labelW = hideY || (!horizontal && primaryAxis === 1) ? 6 : Math.min(100, maxOf(scale.ticks.map((t) => axisLabel(t, scale.code, chart.date1904).length)) * CW + 8);
  const rightScale = scale2 ?? (!horizontal && primaryAxis === 1 ? scale : null);
  const label2W = rightScale && !chart.axes?.y2?.hide ? Math.min(100, maxOf(rightScale.ticks.map((t) => axisLabel(t, rightScale.code, chart.date1904).length)) * CW + 8) : 0;

  // 가로축 값 범위 (분산형)
  let xScale = null;
  if (baseType === 'scatter') {
    const xs = series.flatMap((s) => (s.x ?? s.values.map((_, i) => i + 1))).filter(isNum);
    xScale = chartValueScale(xs, chart.axes?.x ?? {}, { zero: false });
  }

  const categoryConfig = chart.axes?.x ?? {};
  const categoryTitleH = categoryConfig.title ? FS.axis * 1.7 : 0;
  let categoryLayout = horizontal && !hideX ? categoryAxisLayout(categories, { slotCount: categorySlots, horizontal: true, length: Math.max(1, plot.h - 24 - (chart.axes?.y?.title ? FS.axis * 1.7 : 0)), depth: Math.min(260, plot.w * .4), font: FS.axis, interval: categoryConfig.labelInterval, rotation: categoryConfig.labelRotation }) : null;
  const catLabelW = horizontal && !hideX ? (categoryLayout?.extent ?? 0) + categoryTitleH : 0;
  const area = horizontal
    ? { x: plot.x + catLabelW, y: plot.y, w: plot.w - catLabelW - 10, h: plot.h - 24 - (chart.axes?.y?.title ? FS.axis * 1.7 : 0) }
    : (() => {
      // 데이터 표가 있으면 왼쪽에 계열 이름이 들어갈 자리
      const lw = hasTable ? Math.max(labelW, Math.min(140, maxOf(series.map((s) => [...String(s.name)].length)) * CW * 1.25 + 24)) : labelW;
      return { x: plot.x + lw, y: plot.y + 4, w: plot.w - lw - 6 - label2W, h: plot.h - Math.round(FS.axis * 1.9) - (hasTable ? 0 : catLevels.length * LEVEL_H) };
    })();
  const threeD = chart.threeD && ['column', 'bar', 'area', 'line'].includes(chart.type);
  const depth = threeD ? chartDepth(chart, area.w, area.h) : { dx: 0, dy: 0 };
  if (threeD) {
    area.x -= Math.min(0, depth.dx); area.y -= Math.min(0, depth.dy);
    area.w -= Math.abs(depth.dx); area.h -= Math.abs(depth.dy);
  }
  if (horizontal && threeD && !hideX) {
    categoryLayout = categoryAxisLayout(categories, { slotCount: categorySlots, horizontal: true, length: area.h, depth: Math.min(260, plot.w * .4), font: FS.axis, interval: categoryConfig.labelInterval, rotation: categoryConfig.labelRotation });
    const delta = categoryLayout.extent + categoryTitleH - catLabelW;
    area.x += delta; area.w -= delta;
  }
  if (!horizontal && baseType !== 'scatter' && !hideX && !hasTable) {
    categoryLayout = plotCategoryLabels(categories, area, plot, FS.axis, categoryConfig, W, categorySlots);
    area.h -= categoryLayout.extent - Math.round(FS.axis * 1.9) + categoryTitleH;
  }
  if (area.w < 20 || area.h < 20) return finish();
  if (chart.plotAreaFormat) parts[plotBackgroundIndex] = chartAreaSvg(chartAreaFormat(chart, 'plot'), area, { id: uid + 'plot', kind: 'plot' });
  if (threeD) parts.push(chartWalls3D(area, depth, chart.plotFill));
  // 값 축 거꾸로 (엑셀 축 서식 '값을 거꾸로')
  const posFor = (sc) => (v) => {
    const f = (v - sc.min) / (sc.max - sc.min);
    return horizontal ? area.x + (sc.reverse ? 1 - f : f) * area.w : area.y + area.h - (sc.reverse ? 1 - f : f) * area.h;
  };
  const vpos = posFor(scale);
  const vpos2 = scale2 ? posFor(scale2) : vpos;
  const posOf = (s) => (s.axis === 1 && scale2 ? vpos2 : s.axis === primaryAxis || !scale2 ? vpos : vpos2);
  const scaleOf = (s) => s.axis === 1 && scale2 ? scale2 : scale;
  const zeroOf = (s) => { const sc = scaleOf(s); return posOf(s)(Math.max(sc.min, Math.min(0, sc.max))); };

  // 눈금선 + 값 축 레이블
  for (const t of scale.ticks) {
    const p = vpos(t).toFixed(1);
    if (horizontal) {
      parts.push(chart.gridY === false ? '' : `<line x1="${p}" y1="${area.y}" x2="${p}" y2="${area.y + area.h}" stroke="${GRID}"/>`,
        hideY ? '' : axisText(primaryAxis ? 'y2' : 'y', `<text x="${p}" y="${area.y + area.h + Math.round(FS.axis * 1.35)}" text-anchor="middle" font-size="${FS.axis}" fill="${TXT}">${escSvg(axisLabel(t, scale.code, chart.date1904))}</text>`));
    } else {
      parts.push(chart.gridY === false ? '' : `<line x1="${area.x}" y1="${p}" x2="${area.x + area.w}" y2="${p}" stroke="${GRID}"/>`,
        hideY || primaryAxis === 1 ? '' : axisText('y', `<text x="${area.x - 5}" y="${Number(p) + FS.axis * 0.35}" text-anchor="end" font-size="${FS.axis}" fill="${TXT}">${escSvg(axisLabel(t, scale.code, chart.date1904))}</text>`));
    }
  }
  if (rightScale && !horizontal && !chart.axes?.y2?.hide) {
    for (const t of rightScale.ticks) {
      const p = vpos2(t);
      axisText('y2', `<text x="${area.x + area.w + 5}" y="${(p + 3.5).toFixed(1)}" text-anchor="start" font-size="${FS.axis}" fill="${TXT}">${escSvg(axisLabel(t, rightScale.code, chart.date1904))}</text>`);
    }
  }
  const axisTitle = (txt, x, y, rot) => (txt ? `<text x="${x}" y="${y}" text-anchor="middle" font-size="11" fill="${TXT}"${rot ? ` transform="rotate(${rot} ${x} ${y})"` : ''}>${escSvg(txt)}</text>` : '');

  const base = vpos(Math.max(scale.min, Math.min(0, scale.max)));
  const labelsOut = [];
  const pushLabel = (x, y, v, s, anchor = 'middle', index) => {
    labelsOut.push(`<text data-el="label"${tag(s, index)} x="${x.toFixed(1)}" y="${y.toFixed(1)}" text-anchor="${anchor}" font-size="${pt(s.labelSize ?? 9)}"${s.labelBold ? ' font-weight="700"' : ''} fill="${s.labelColor ?? '#404040'}" paint-order="stroke" stroke="#fff" stroke-width="2.5">${escSvg(chartDataLabel(chart, s, categories, index, v))}</text>`);
  };

  if (baseType === 'scatter') {
    const xpos = (x) => { const f = (x - xScale.min) / (xScale.max - xScale.min); return area.x + (xScale.reverse ? 1 - f : f) * area.w; };
    if (!hideX) for (const t of xScale.ticks) axisText('x', `<text x="${xpos(t).toFixed(1)}" y="${area.y + area.h + Math.round(FS.axis * 1.35)}" text-anchor="middle" font-size="${FS.axis}" fill="${TXT}">${escSvg(axisLabel(t, xScale.code, chart.date1904))}</text>`);
    const clipId = `${uid}scatterPlot`; defs.push(`<clipPath id="${clipId}"><rect x="${area.x}" y="${area.y}" width="${area.w}" height="${area.h}"/></clipPath>`); parts.push(`<g clip-path="url(#${clipId})">`);
    const sty = chart.scatterStyle ?? 'marker';
    series.forEach((s) => {
      const pts = [];
      s.values.forEach((v, i) => {
        const x = s.x ? s.x[i] : i + 1;
        if (isNum(v) && isNum(x)) pts.push([xpos(x), posOf(s)(v), v, i]);
      });
      if (/line|smooth/i.test(sty) && pts.length > 1) {
        const d = (s.smooth ?? /smooth/i.test(sty)) ? smoothPath(pts) : pts.map((p, i) => `${i ? 'L' : 'M'}${p[0].toFixed(1)},${p[1].toFixed(1)}`).join('');
        parts.push(`<path d="${d}" fill="none" stroke="${s.color}" stroke-width="${s.lineWidth ?? 2}" stroke-linejoin="round"${dashAttr(s.dash, s.lineWidth ?? 2)}${shadowAttr(s)}${tag(s)}/>`);
      }
      const marker = s.marker ?? chart.marker, hasMarker = marker === undefined ? sty === 'marker' || /Marker$/.test(sty) : marker !== false && marker !== 'none';
      if (hasMarker) for (const p of pts) parts.push(tagMk((MARKERS[marker] ?? MARKERS.circle)(p[0].toFixed(1), p[1].toFixed(1), s.markerSize ? s.markerSize * 2 / 3 : 3.5, pointColor(s, p[3]) ?? s.markerColor ?? s.color), s, p[3]));
      if (wantLabels(s)) for (const p of pts) pushLabel(p[0], p[1] - 7, p[2], s, 'middle', p[3]);
    });
    parts.push('</g>');
  } else {
    // 항목 축 레이블
    const band = (horizontal ? area.h : area.w) / Math.max(1, categorySlots);
    if (chart.gridX) {
      for (let i = 1; i < n; i++) {
        const q = ((horizontal ? area.y : area.x) + band * (chart.axes?.x?.reverse ? categorySlots - i : i)).toFixed(1);
        parts.push(horizontal ? `<line x1="${area.x}" y1="${q}" x2="${area.x + area.w}" y2="${q}" stroke="${GRID}"/>` : `<line x1="${q}" y1="${area.y}" x2="${q}" y2="${area.y + area.h}" stroke="${GRID}"/>`);
      }
    }
    if (!hideX && !hasTable && categoryLayout) for (const label of categoryLayout.labels) {
      const mid = (horizontal ? area.y : area.x) + band * (categoryIndex(label.index) + .5);
      axisText('x', categoryAxisLabelSvg(categoryLayout, label, mid, horizontal ? area.x : area.y + area.h, TXT));
    }
    if (catLevels.length && !horizontal && !hasTable) {
      // 묶음 경계선은 항목 축에서 해당 단계 줄 아래까지
      const y0 = area.y + area.h;
      catLevels.forEach((spans, k) => {
        const yText = y0 + (categoryLayout?.extent ?? Math.round(FS.axis * 1.9)) + (k + .75) * LEVEL_H;
        const yEnd = y0 + (categoryLayout?.extent ?? Math.round(FS.axis * 1.9)) + (k + 1) * LEVEL_H;
        for (const sp of spans) {
          const x1 = area.x + band * (chart.axes?.x?.reverse ? categorySlots - 1 - sp.end : sp.start);
          const w = band * (sp.end - sp.start + 1);
          parts.push(`<line x1="${x1.toFixed(1)}" y1="${y0}" x2="${x1.toFixed(1)}" y2="${yEnd}" stroke="${GRID}"/>`);
          if (sp.text) axisText('x', `<text x="${(x1 + w / 2).toFixed(1)}" y="${yText}" text-anchor="middle" font-size="${FS.axis}" fill="${TXT}">${escSvg(truncate(sp.text, Math.max(2, Math.floor(w / (CW * 1.4)))))}</text>`);
        }
        const xr = area.x + band * (chart.axes?.x?.reverse ? categorySlots : n);
        parts.push(`<line x1="${xr.toFixed(1)}" y1="${y0}" x2="${xr.toFixed(1)}" y2="${yEnd}" stroke="${GRID}"/>`);
      });
    }
    // 막대 (묶은 · 누적)
    const bars = series.filter((s) => barTypes.has(s.type)), barStart = parts.length;
    if (bars.length) {
      const deep = threeD && chart.type === 'column' && chart.grouping === 'standard';
      const slots = [[],[]], slotOf = new Map();
      bars.forEach((s,bi) => { const key=s._stacked?s._stackGroup:`single/${bi}`, keys=slots[s.axis]; if(!keys.includes(key))keys.push(key); slotOf.set(s,keys.indexOf(key)); });
      const groupW = isNum(chart.gap) ? band / (1 + Math.max(0, chart.gap) / 100) : band * 0.65;
      // 계열 겹치기 (-100 ~ 100%): 양수면 막대가 겹치고 음수면 간격이 생김
      const paintBars = deep ? [...bars].reverse() : bars;
      const vary = chart.varyColors && bars.length === 1;
      paintBars.forEach((s) => {
        const bi=bars.indexOf(s), k=deep?1:Math.max(1,slots[s.axis].length);
        const ov=isNum(chart.overlap)&&k>1?Math.max(-100,Math.min(100,chart.overlap))/100:0;
        const barW=groupW/(k-(k-1)*ov),stepW=barW*(1-ov);
        const lane=deep?bi/Math.max(1,bars.length):0, localDepth=deep?{dx:depth.dx/Math.max(1,bars.length)*.7,dy:depth.dy/Math.max(1,bars.length)*.7}:depth;
        const vp = posOf(s);
        const bf = fillOf(s);
        const sh = shadowAttr(s);
        const b0 = zeroOf(s);
        s.values.forEach((raw, i) => {
          if (!isNum(raw)) return;
          const v = s._upper[i];
          const start = (horizontal ? area.y : area.x) + band * categoryIndex(i) + (band - groupW) / 2 + stepW * (deep ? 0 : slotOf.get(s));
          const pf = pointColor(s, i) ?? (vary && !chart.seriesFmt?.[s._fi]?.color ? pal[chartPointIndex(s, i) % pal.length] : bf);
          let from = b0;
          let to = vp(v);
          if (s._stacked) {
            from = vp(s._lower[i]);
            to = vp(s._upper[i]);
          }
          const a = Math.min(from, to);
          const len = Math.abs(to - from);
          const stroke = s.outline ? ` stroke="${s.outline}" stroke-width="1"` : '';
          if (threeD) {
            const x = (horizontal ? a : start)+depth.dx*lane, y = (horizontal ? start : a)+depth.dy*lane;
            const w = horizontal ? len : Math.max(1, barW - 1), h = horizontal ? Math.max(1, barW - 1) : len;
            parts.push(barSolid3D({x,y,w,h,from:from+(horizontal?depth.dx:depth.dy)*lane,to:to+(horizontal?depth.dx:depth.dy)*lane,horizontal,dx:localDepth.dx,dy:localDepth.dy,shape:s.barShape??chart.barShape,fill:pf,attrs:stroke+sh+tag(s,i)+(deep?` data-depth="${bi}"`:'')}));
          } else if (horizontal) parts.push(`<rect x="${a.toFixed(1)}" y="${start.toFixed(1)}" width="${len.toFixed(1)}" height="${Math.max(1, barW - 1).toFixed(1)}" fill="${pf}"${stroke}${sh}${tag(s, i)}/>`);
          else parts.push(`<rect x="${start.toFixed(1)}" y="${a.toFixed(1)}" width="${Math.max(1, barW - 1).toFixed(1)}" height="${len.toFixed(1)}" fill="${pf}"${stroke}${sh}${tag(s, i)}/>`);
          if (wantLabels(s)) {
            // 레이블 위치: 바깥쪽 끝(기본) · 가운데 · 안쪽 끝 · 안쪽 축
            const lp = s._stacked && !s.labelPos ? 'center' : s.labelPos ?? 'outEnd';
            if (horizontal) {
              const hi = Math.max(from, to);
              const lo = Math.min(from, to);
              const x = lp === 'center' ? (from + to) / 2 : lp === 'insideEnd' ? hi - 4 : lp === 'insideBase' ? lo + 4 : hi + 4;
              pushLabel(x, start + barW / 2 + 3.5, raw, s, lp === 'center' ? 'middle' : lp === 'insideEnd' ? 'end' : 'start', i);
            } else {
              const topY = Math.min(from, to);
              const botY = Math.max(from, to);
              const y = lp === 'center' ? (from + to) / 2 + 3.5 : lp === 'insideEnd' ? topY + 12 : lp === 'insideBase' ? botY - 4 : topY - 4;
              pushLabel(start + barW / 2, y, raw, s, 'middle', i);
            }
          }
        });
      });
    }
    const behindArea = series.some(s => s.type === 'area') ? parts.splice(barStart) : null;
    // 영역은 막대 뒤에, 꺾은선은 막대 위에 그린다.
    // 누적 꺾은선 · 영역: 앞 계열 값 위에 쌓음 (100% 기준이면 항목 합계로 나눔)
    for (const type of ['area', 'line']) {
      if (type === 'line' && behindArea) pushAll(parts, behindArea);
      series.filter((s) => s.type === type).forEach((s) => {
        const vp = posOf(s);
        const lower = s._stacked ? s._lower : null;
        const pts = s.values.map((v0, i) => {
          if (!isNum(v0)) return null;
          const v = s._upper[i];
          return [area.x + band * (categoryIndex(i) + 0.5), vp(v), v0, i];
        });
        const segs = [];
        let cur = [];
        for (const p of pts) { if (p) cur.push(p); else if (cur.length) { segs.push(cur); cur = []; } }
        if (cur.length) segs.push(cur);
        const b0 = zeroOf(s);
        for (const seg of segs) {
          const d = s.smooth ? smoothPath(seg) : seg.map((p, i) => `${i ? 'L' : 'M'}${p[0].toFixed(1)},${p[1].toFixed(1)}`).join('');
          if (type === 'area') {
            const back = lower
              ? seg.slice().reverse().map((p) => `L${p[0].toFixed(1)},${vp(lower[p[3]]).toFixed(1)}`).join('')
              : `L${seg.at(-1)[0].toFixed(1)},${b0.toFixed(1)}L${seg[0][0].toFixed(1)},${b0.toFixed(1)}`;
            if (threeD) {
              const lowerPoints = lower ? seg.slice().reverse().map((p) => [p[0], vp(lower[p[3]])]) : [[seg.at(-1)[0], b0], [seg[0][0], b0]];
              parts.push(extrudedPolygon([...seg.map((p) => [p[0], p[1]]), ...lowerPoints], depth.dx, depth.dy, fillOf(s), tag(s)));
            } else parts.push(`<path d="${d}${back}Z" fill="${fillOf(s)}" fill-opacity="${s._stacked ? 0.9 : 0.75}"${tag(s)}/>`);
          } else {
            if (threeD) for (let j = 1; j < seg.length; j++) {
              const a = seg[j - 1], b = seg[j];
              parts.push(`<polygon data-3d="ribbon" points="${a[0]},${a[1]} ${b[0]},${b[1]} ${b[0] + depth.dx},${b[1] + depth.dy} ${a[0] + depth.dx},${a[1] + depth.dy}" fill="${s.color}" fill-opacity="0.85"${tag(s)}/>`);
            }
            parts.push(`<path d="${d}" fill="none" stroke="${s.color}" stroke-width="${s.lineWidth ?? 2.25}" stroke-linejoin="round" stroke-linecap="round"${dashAttr(s.dash, s.lineWidth ?? 2.25)}${shadowAttr(s)}${tag(s)}/>`);
            const mkName = s.marker ?? chart.marker;
            if (mkName !== false && mkName !== 'none') {
              const mk = MARKERS[mkName] ?? MARKERS.circle;
              const mr = s.markerSize ? (s.markerSize * 4) / 3 / 2 : 3;
              for (const p of seg) parts.push(tagMk(mk(p[0].toFixed(1), p[1].toFixed(1), mr, pointColor(s, p[3]) ?? s.markerColor ?? s.color), s, p[3]));
            }
          }
          if (wantLabels(s)) for (const p of seg) pushLabel(p[0], p[1] - 7, p[2], s, 'middle', p[3]);
        }
      });
    }
  }
  // Clip after layout; manual axis bounds remain authoritative.
  const trendClip = `${uid}trendPlot`;
  if ([...trends.values()].some(points => points.length)) {
    defs.push(`<clipPath id="${trendClip}" clipPathUnits="userSpaceOnUse"><rect x="${area.x}" y="${area.y}" width="${area.w}" height="${area.h}"/></clipPath>`);
    const bandT = (horizontal ? area.h : area.w) / Math.max(1, categorySlots);
    for (const s of series) {
      const pts = trends.get(s);
      if (!pts.length) continue;
      const vp = posOf(s), at = i => (horizontal ? area.y : area.x) + bandT * (categoryIndex(i) + .5);
      const path = pts.map(([x, y], i) => `${i ? 'L' : 'M'}${(horizontal ? vp(y) : at(x)).toFixed(1)},${(horizontal ? at(x) : vp(y)).toFixed(1)}`).join('');
      parts.push(`<path data-trend="${s.trend}" d="${path}" fill="none" stroke="${s.trendColor ?? s.color}" stroke-width="1.5" stroke-dasharray="5 3" clip-path="url(#${trendClip})"/>`);
    }
  }
  // 데이터 표
  if (hasTable) {
    const y0 = area.y + area.h + 2;
    const bandT = area.w / Math.max(1, categorySlots);
    const x0 = plot.x;
    const rows = [['', ...categories], ...series.map((s) => [s.name, ...s.values.map((v) => (isNum(v) ? valueLabel(v, s.numFmt, chart.date1904) : ''))])];
    const lineC = '#d9d9d9';
    parts.push(`<rect x="${x0}" y="${y0}" width="${area.x + area.w - x0}" height="${DT_ROW * rows.length}" fill="none" stroke="${lineC}"/>`);
    rows.forEach((row, r) => {
      const y = y0 + DT_ROW * r;
      if (r) parts.push(`<line x1="${x0}" y1="${y}" x2="${area.x + area.w}" y2="${y}" stroke="${lineC}"/>`);
      // 첫 칸: 계열 이름 + 범례 표지
      if (r) {
        const s = series[r - 1];
        parts.push(s.type === 'line' ? `<line x1="${x0 + 3}" y1="${y + DT_ROW / 2}" x2="${x0 + 13}" y2="${y + DT_ROW / 2}" stroke="${s.color}" stroke-width="2"/>` : `<rect x="${x0 + 4}" y="${y + DT_ROW / 2 - 4}" width="8" height="8" fill="${s.color}"/>`);
        parts.push(T(x0 + 16, y + DT_ROW / 2 + FS.axis * 0.35, truncate(String(row[0]), Math.max(2, Math.floor((area.x - x0 - 18) / CW))), FS.axis, TXT, 'start'));
      }
      row.slice(1).forEach((cell, i) => {
        const cx = area.x + bandT * (categoryIndex(i) + 0.5);
        parts.push(T(cx, y + DT_ROW / 2 + FS.axis * 0.35, truncate(String(cell), Math.max(2, Math.floor(bandT / (CW * 1.1)))), FS.axis, TXT));
      });
    });
    for (let i = 0; i <= n; i++) {
      const x = area.x + bandT * (chart.axes?.x?.reverse ? categorySlots - i : i);
      parts.push(`<line x1="${x.toFixed(1)}" y1="${y0}" x2="${x.toFixed(1)}" y2="${y0 + DT_ROW * rows.length}" stroke="${lineC}"/>`);
    }
  }
  // 기준선
  if (horizontal) parts.push(`<line x1="${base}" y1="${area.y}" x2="${base}" y2="${area.y + area.h}" stroke="#bfbfbf"/>`);
  else parts.push(`<line x1="${area.x}" y1="${base}" x2="${area.x + area.w}" y2="${base}" stroke="#bfbfbf"/>`);
  pushAll(parts, labelsOut);
  if (horizontal) {
    if (!hideY) axisText(primaryAxis ? 'y2' : 'y', axisTitle(chart.axes?.[primaryAxis ? 'y2' : 'y']?.title, area.x + area.w / 2, plot.y + plot.h));
    if (!hideX) axisText('x', axisTitle(categoryConfig.title, plot.x + FS.axis, area.y + area.h / 2, -90));
  } else {
    if (!hideY && primaryAxis === 0) axisText('y', axisTitle(chart.axes?.y?.title, plot.x + 6, area.y + area.h / 2, -90));
    if (rightScale && !chart.axes?.y2?.hide) axisText('y2', axisTitle(chart.axes?.y2?.title, area.x + area.w + label2W - 2, area.y + area.h / 2, 90));
    if (!hideX) axisText('x', axisTitle(categoryConfig.title, area.x + area.w / 2, H - (legendH ? legendH + 2 : 2)));
  }
  return finish();
}

/** 선 종류 → stroke-dasharray (실선 · 파선 · 점선 · 일점쇄선 · 긴 파선) */
function dashAttr(dash, w = 2) {
  const k = Math.max(1, w);
  const pat = { dash: [4, 3], dot: [1, 2], dashDot: [4, 2, 1, 2], longDash: [8, 3], sysDash: [3, 1] }[dash];
  return pat ? ` stroke-dasharray="${pat.map((x) => x * k).join(' ')}"` : '';
}

/** 부드러운 곡선 (Catmull-Rom → 베지어) */
function smoothPath(pts) {
  let d = `M${pts[0][0].toFixed(1)},${pts[0][1].toFixed(1)}`;
  for (let i = 0; i < pts.length - 1; i++) {
    const p0 = pts[i - 1] ?? pts[i];
    const p1 = pts[i];
    const p2 = pts[i + 1];
    const p3 = pts[i + 2] ?? p2;
    const c1 = [p1[0] + (p2[0] - p0[0]) / 6, p1[1] + (p2[1] - p0[1]) / 6];
    const c2 = [p2[0] - (p3[0] - p1[0]) / 6, p2[1] - (p3[1] - p1[1]) / 6];
    d += `C${c1[0].toFixed(1)},${c1[1].toFixed(1)} ${c2[0].toFixed(1)},${c2[1].toFixed(1)} ${p2[0].toFixed(1)},${p2[1].toFixed(1)}`;
  }
  return d;
}

// ───────────── 통계 · 계층 차트 (엑셀 2016 이후 차트) 와 방사형 · 거품형 · 주식형 ─────────────
const T = (x, y, txt, size, fill, anchor = 'middle', extra = '') => `<text x="${x.toFixed(1)}" y="${y.toFixed(1)}" text-anchor="${anchor}" font-size="${size}" fill="${fill}"${extra}>${escSvg(txt)}</text>`;

/** 값 범위/주 단위를 제한해 잘못된 파일의 무한 눈금 반복을 막습니다. */
export function chartValueScale(values, cfg = {}, { zero = true, code = null } = {}) {
  let lo = Infinity, hi = -Infinity;
  for (const v of values) if (isNum(v)) { lo = Math.min(lo, v); hi = Math.max(hi, v); }
  if (!Number.isFinite(lo)) { lo = 0; hi = 1; }
  if (zero) { lo = Math.min(0, lo); hi = Math.max(0, hi); }
  const sc = niceScale(isNum(cfg.min) ? cfg.min : lo, isNum(cfg.max) ? cfg.max : hi);
  if (isNum(cfg.min)) sc.min = cfg.min;
  if (isNum(cfg.max)) sc.max = cfg.max;
  if (!(sc.max > sc.min)) Object.assign(sc, niceScale(lo, hi));
  if (isNum(cfg.major) && cfg.major > 0 && (sc.max - sc.min) / cfg.major <= 200) sc.step = cfg.major;
  const count = Math.min(200, Math.floor((sc.max - sc.min) / sc.step + 1e-9));
  const ticks = Array.from({ length: count + 1 }, (_, i) => Number((sc.min + i * sc.step).toPrecision(12)));
  return { ...sc, ticks, reverse: !!cfg.reverse, code: cfg.numFmt ?? code };
}

/** 특수 차트 공용 직교 영역. 축 숨김과 눈금선 표시는 별개입니다. */
function cartesian(ctx, vals, { horizontal = false, code = null, cats = null, zero = true, right = 0, secondary = null } = {}) {
  const { chart, plot, parts, FS, TXT, GRID, defs, uid } = ctx;
  const cfg = chart.axes?.y ?? {}, xc = chart.axes?.x ?? {}, yc2 = chart.axes?.y2 ?? {};
  const sc = chartValueScale(vals, cfg, { zero, code });
  const sc2 = secondary ? chartValueScale(secondary.values, { ...secondary.defaults, ...yc2 }, { zero: secondary.zero ?? false, code: secondary.code }) : null;
  const CW = FS.axis * .58, titleH = FS.axis * 1.7;
  const lw = (scale) => Math.min(100, maxOf(scale.ticks.map((t) => axisLabel(t, scale.code, chart.date1904).length)) * CW + 8);
  let categoryLayout = horizontal && cats && !xc.hide ? categoryAxisLayout(cats, { horizontal: true, length: Math.max(1, plot.h - FS.axis * 1.9 - (cfg.title ? titleH : 0)), depth: Math.min(260, plot.w * .4), font: FS.axis, interval: xc.labelInterval, rotation: xc.labelRotation }) : null;
  const catW = categoryLayout?.extent ?? 0;
  const left = horizontal ? (xc.hide ? 5 : catW + (xc.title ? titleH : 0)) : cfg.hide ? 5 : lw(sc) + (cfg.title ? titleH : 0);
  const bottom = horizontal ? (cfg.hide ? 4 : FS.axis * 1.9 + (cfg.title ? titleH : 0)) : xc.hide ? 4 : FS.axis * 1.9 + (xc.title ? titleH : 0);
  const rightW = sc2 && !yc2.hide ? lw(sc2) + (yc2.title ? titleH : 0) : right;
  const area = { x: plot.x + left, y: plot.y + 4, w: Math.max(4, plot.w - left - rightW - 6), h: Math.max(4, plot.h - bottom - 4) };
  if (!horizontal && cats && !xc.hide) {
    categoryLayout = plotCategoryLabels(cats, area, plot, FS.axis, xc, ctx.W);
    area.h = Math.max(4, area.h - categoryLayout.extent + FS.axis * 1.9);
  }
  const position = (v, scale) => { const q = (v - scale.min) / (scale.max - scale.min), t = scale.reverse ? 1 - q : q; return horizontal ? area.x + t * area.w : area.y + (1 - t) * area.h; };
  const pos = (v) => position(v, sc), pos2 = (v) => position(v, sc2);
  const title = (name, x, y, rotation = 0, axis = 'y') => { if (name) ctx.axisText(axis, T(x, y, name, FS.axis, TXT, 'middle', ` data-axis-title="${axis}"${rotation ? ` transform="rotate(${rotation} ${x} ${y})"` : ''}`)); };
  for (const t of sc.ticks) {
    const p = pos(t);
    if (chart.gridY !== false) parts.push(horizontal ? `<line data-grid="y" x1="${p.toFixed(1)}" y1="${area.y}" x2="${p.toFixed(1)}" y2="${area.y + area.h}" stroke="${GRID}"/>` : `<line data-grid="y" x1="${area.x}" y1="${p.toFixed(1)}" x2="${area.x + area.w}" y2="${p.toFixed(1)}" stroke="${GRID}"/>`);
    if (!cfg.hide) ctx.axisText('y', horizontal ? T(p, area.y + area.h + FS.axis * 1.35, axisLabel(t, sc.code, chart.date1904), FS.axis, TXT, 'middle', ' data-axis="y"') : T(area.x - 5, p + FS.axis * .35, axisLabel(t, sc.code, chart.date1904), FS.axis, TXT, 'end', ' data-axis="y"'));
  }
  if (!cfg.hide) horizontal ? title(cfg.title, area.x + area.w / 2, plot.y + plot.h) : title(cfg.title, plot.x + FS.axis, area.y + area.h / 2, -90);
  if (sc2 && !yc2.hide) {
    for (const t of sc2.ticks) ctx.axisText('y2', T(area.x + area.w + 5, pos2(t) + FS.axis * .35, axisLabel(t, sc2.code, chart.date1904), FS.axis, TXT, 'start', ' data-axis="y2"'));
    title(yc2.title, plot.x + plot.w - FS.axis * .35, area.y + area.h / 2, 90, 'y2');
  }
  const n = cats?.length ?? 0, band = (horizontal ? area.h : area.w) / Math.max(1, n);
  const mid = (i) => (horizontal ? area.y : area.x) + band * (xc.reverse ? n - i - .5 : i + .5);
  if (cats) {
    cats.forEach((c, i) => {
      const p = mid(i);
      if (chart.gridX) parts.push(horizontal ? `<line data-grid="x" x1="${area.x}" y1="${p}" x2="${area.x + area.w}" y2="${p}" stroke="${GRID}"/>` : `<line data-grid="x" x1="${p}" y1="${area.y}" x2="${p}" y2="${area.y + area.h}" stroke="${GRID}"/>`);
    });
    if (!xc.hide && categoryLayout) for (const label of categoryLayout.labels) ctx.axisText('x', categoryAxisLabelSvg(categoryLayout, label, mid(label.index), horizontal ? area.x : area.y + area.h, TXT));
  }
  if (!xc.hide) horizontal ? title(xc.title, plot.x + FS.axis, area.y + area.h / 2, -90, 'x') : title(xc.title, area.x + area.w / 2, plot.y + plot.h, 0, 'x');
  const clipId = `${uid}plot`;
  defs.push(`<clipPath id="${clipId}"><rect x="${area.x}" y="${area.y}" width="${area.w}" height="${area.h}"/></clipPath>`);
  return { area, pos, pos2, band, mid, sc, sc2, fmt: sc.code, base: pos(Math.max(sc.min, Math.min(0, sc.max))), clip: ` clip-path="url(#${clipId})"` };
}

const labelTxt = (ctx, s, v, x, y, anchor = 'middle') => ctx.parts.push(`<text x="${x.toFixed(1)}" y="${y.toFixed(1)}" text-anchor="${anchor}" font-size="${(s?.labelSize ?? 9) * 4 / 3}" fill="${s?.labelColor ?? '#404040'}" paint-order="stroke" stroke="#fff" stroke-width="2.5">${escSvg(valueLabel(v, s?.numFmt, ctx.chart.date1904))}</text>`);

/** Excel 상자수염은 홀수 표본의 중앙값을 양쪽 절반에 포함/제외한 뒤 절반의 중앙값을 씁니다.
 * https://learn.microsoft.com/en-us/openspecs/office_standards/ms-odrawxml/a1463864-276e-4cef-b772-d9b2a046e5a7
 */
export function boxWhiskerStats(values, method = 'inclusive') {
  const v = values.filter(isNum).sort((a, b) => a - b);
  if (!v.length) return null;
  const median = (xs) => xs.length % 2 ? xs[(xs.length - 1) / 2] : xs[xs.length / 2 - 1] / 2 + xs[xs.length / 2] / 2;
  const half = Math.floor(v.length / 2), include = method !== 'exclusive' && v.length % 2;
  const lower = v.slice(0, half + (include ? 1 : 0)), upper = v.slice(v.length - half - (include ? 1 : 0));
  const q1 = median(lower.length ? lower : v), q3 = median(upper.length ? upper : v);
  const iqr = q3 - q1, inside = [], out = [];
  for (const x of v) (x < q1 - 1.5 * iqr || x > q3 + 1.5 * iqr ? out : inside).push(x);
  return { count: v.length, q1, q3, med: median(v), mean: v.reduce((sum, x) => sum + x / v.length, 0), lo: inside[0] ?? q1, hi: inside.at(-1) ?? q3, inside, out, min: v[0], max: v.at(-1) };
}

/** 히스토그램 구간: 엑셀 자동 구간(스콧 규칙)과 비슷하게 */
export function histogramBins(values, count = null, width = null) {
  const v = values.filter(isNum).sort((a, b) => a - b);
  if (!v.length) return [];
  const min = v[0];
  const max = v.at(-1);
  let w = width;
  if (!w) {
    const n = count ?? Math.max(1, Math.min(30, Math.ceil(Math.sqrt(v.length))));
    w = (max - min) / n || 1;
    const mag = 10 ** Math.floor(Math.log10(w));
    w = Math.ceil(w / mag * 2) / 2 * mag;
  }
  const bins = [];
  for (let a = min; a <= max || !bins.length; a += w) {
    const b = a + w;
    bins.push({ from: a, to: b, count: 0 });
    if (bins.length > 200) break;
  }
  for (const x of v) {
    let k = Math.min(bins.length - 1, Math.floor((x - min) / w));
    if (k > 0 && x === bins[k].from) k -= 1; // (a, b] — 엑셀처럼 오른쪽 끝 포함
    bins[k].count++;
  }
  return bins;
}

const binLabel = (b, i) => `${i ? '(' : '['}${axisLabel(Number(b.from.toPrecision(10)))}, ${axisLabel(Number(b.to.toPrecision(10)))}]`;
const TOTAL_RE = /^(합계|총계|총합계|소계|전체|total|subtotal|grand total|net)$/i;

function waterfallBars(values, categories, chart) {
  const totals = new Set(chart.totals ?? categories.map((c, i) => TOTAL_RE.test(String(c).trim()) ? i : -1).filter((i) => i >= 0));
  let run = 0;
  return values.map((v0, i) => {
    const v = isNum(v0) ? v0 : 0;
    if (totals.has(i)) { if (i === 0) run = v; return { a: 0, b: run, v: run, kind: 'total' }; }
    const a = run; run += v;
    return { a, b: run, v, kind: v >= 0 ? 'up' : 'down' };
  });
}

function paretoItems(series, categories, chart) {
  const s = series[0];
  if (!s) return [];
  const items = categories.every((c, i) => String(c) === String(i + 1))
    ? histogramBins(s.values, chart.binCount ?? null, chart.binWidth ?? null).map((b, i) => ({ c: binLabel(b, i), v: b.count, i }))
    : s.values.map((v, i) => ({ c: categories[i], v: isNum(v) ? v : 0, i }));
  return items.sort((a, b) => b.v - a.v);
}

/** 통계 차트의 표에는 원자료 대신 화면에 집계된 값/요약 통계를 표시합니다. */
function specialDataTable(chart, series, categories) {
  if (!series.length) return null;
  const s = series[0], values = (row, code = s.numFmt) => row.map((v) => isNum(v) ? valueLabel(v, code, chart.date1904) : '');
  if (chart.type === 'waterfall') return [['', ...categories], [s.name, ...values(waterfallBars(s.values, categories, chart).map((b) => b.v))]];
  if (chart.type === 'histogram') { const bins = histogramBins(series.flatMap((s) => s.values), chart.binCount ?? null, chart.binWidth ?? null); return [['구간', ...bins.map(binLabel)], ['빈도', ...values(bins.map((b) => b.count), '0')]]; }
  if (chart.type === 'pareto') {
    const items = paretoItems(series, categories, chart), total = items.reduce((sum, p) => sum + Math.max(0, p.v), 0) || 1;
    let acc = 0;
    return [['', ...items.map((p) => p.c)], [s.name, ...values(items.map((p) => p.v))], ['누적 %', ...items.map((p) => valueLabel((acc += Math.max(0, p.v)) / total, '0.0%', chart.date1904))]];
  }
  if (chart.type === 'boxWhisker') return [['계열', '최솟값', 'Q1', '중앙값', 'Q3', '최댓값', '평균'], ...series.map((sr) => { const st = boxWhiskerStats(sr.values, chart.quartileMethod); return [sr.name, ...values(st ? [st.min, st.q1, st.med, st.q3, st.max, st.mean] : Array(6).fill(null), sr.numFmt)]; })];
  if (chart.type === 'stock') return [['', ...categories], ...series.map((sr) => [sr.name, ...values(sr.values, sr.numFmt)])];
  return null;
}

function drawSpecialTable(ctx, rows, height) {
  if (!rows.length || height <= 0) return;
  const { plot, parts, FS, TXT, GRID } = ctx, y = plot.y + plot.h + 6, h = height / rows.length;
  const count = Math.max(1, maxOf(rows.map((r) => r.length))), w = plot.w / count, font = Math.min(FS.axis, h * .65);
  parts.push(`<g data-el="dataTable"><rect x="${plot.x}" y="${y}" width="${plot.w}" height="${height}" fill="none" stroke="${GRID}"/>`);
  for (let c = 1; c < count; c++) parts.push(`<line x1="${plot.x + w * c}" y1="${y}" x2="${plot.x + w * c}" y2="${y + height}" stroke="${GRID}"/>`);
  rows.forEach((row, r) => {
    if (r) parts.push(`<line x1="${plot.x}" y1="${y + h * r}" x2="${plot.x + plot.w}" y2="${y + h * r}" stroke="${GRID}"/>`);
    row.forEach((v, c) => parts.push(T(plot.x + w * (c + .5), y + h * (r + .5) + font * .35, truncate(String(v ?? ''), Math.max(1, Math.floor((w - 5) / (font * .7)))), font, TXT)));
  });
  parts.push('</g>');
}

const specialBarWidth = (band, chart, fallback) => band / (1 + Math.max(0, Math.min(500, isNum(chart.gap) ? chart.gap : fallback)) / 100);

export const SPECIAL = {
  waterfall: {
    draw(ctx) {
      const { chart, series, categories, parts, pal } = ctx;
      const s = series[0];
      if (!s) return;
      const bars = waterfallBars(s.values, categories, chart);
      const { area, pos, band, mid, base, clip } = cartesian(ctx, bars.flatMap((b) => [b.a, b.b]), { cats: categories, code: s.numFmt });
      const w = specialBarWidth(band, chart, 61), direction = chart.axes?.x?.reverse ? -1 : 1;
      const col = { up: chart.upColor ?? pal[0], down: chart.downColor ?? pal[1 % pal.length], total: chart.totalColor ?? pal[2 % pal.length] };
      parts.push(`<g data-plot="waterfall"${clip}>`);
      bars.forEach((b, i) => {
        const x = mid(i) - w / 2, y1 = pos(b.a), y2 = pos(b.b);
        parts.push(`<rect x="${x.toFixed(1)}" y="${Math.min(y1, y2).toFixed(1)}" width="${w.toFixed(1)}" height="${Math.max(1, Math.abs(y2 - y1)).toFixed(1)}" fill="${s.pointColors?.[i] ?? col[b.kind]}" data-s="${s._fi}" data-p="${i}" data-value="${b.v}"/>`);
        if (i < bars.length - 1 && chart.connectors !== false) {
          const yb = pos(b.b);
          parts.push(`<line x1="${(mid(i) + direction * w / 2).toFixed(1)}" y1="${yb.toFixed(1)}" x2="${(mid(i + 1) - direction * w / 2).toFixed(1)}" y2="${yb.toFixed(1)}" stroke="#a6a6a6" stroke-dasharray="2,2"/>`);
        }
      });
      parts.push('</g>');
      bars.forEach((b, i) => { if ((s.labels ?? chart.labels) !== false) labelTxt(ctx, s, b.v, mid(i), Math.max(area.y + 12, Math.min(pos(b.a), pos(b.b)) - 4)); });
      if (!chart.axes?.y?.hide) parts.push(`<line x1="${area.x}" y1="${base.toFixed(1)}" x2="${area.x + area.w}" y2="${base.toFixed(1)}" stroke="#bfbfbf"/>`);
    },
  },
  funnel: {
    legend: false,
    draw(ctx) {
      const { chart, series, categories, plot, parts, FS, TXT } = ctx;
      const s = series[0];
      if (!s) return;
      const vals = s.values.map((v) => (isNum(v) && v > 0 ? v : 0));
      const max = Math.max(maxOf(vals), 1);
      const CW = FS.axis * 0.58;
      const lw = Math.min(160, maxOf(categories.map((c) => [...String(c)].length)) * CW * 1.4 + 10);
      const area = { x: plot.x + lw, y: plot.y + 2, w: plot.w - lw - 4, h: plot.h - 4 };
      const band = area.h / Math.max(1, vals.length);
      vals.forEach((v, i) => {
        const bw = (v / max) * area.w;
        const x = area.x + (area.w - bw) / 2;
        const y = area.y + band * i + band * 0.08;
        const tag = ` data-s="${s._fi}" data-p="${chartPointIndex(s, i)}"`;
        parts.push(`<rect x="${x.toFixed(1)}" y="${y.toFixed(1)}" width="${Math.max(1, bw).toFixed(1)}" height="${(band * 0.84).toFixed(1)}" fill="${pointColor(s, i) ?? s.color}"${tag}/>`);
        parts.push(T(area.x - 6, y + band * 0.42 + 4, truncate(String(categories[i] ?? ''), 20), FS.axis, TXT, 'end'));
        const labelSeries = { ...s, labels: s.labels ?? chart.labels ?? true };
        if (labelSeries.labels || s.catName || s.serName || s.pct) parts.push(`<text data-el="label"${tag} x="${(area.x + area.w / 2).toFixed(1)}" y="${(y + band * 0.42 + 4).toFixed(1)}" text-anchor="middle" font-size="${s.labelSize ? s.labelSize * 4 / 3 : FS.axis}" fill="${s.labelColor ?? '#fff'}"${s.labelBold !== false ? ' font-weight="700"' : ''}>${escSvg(chartDataLabel(chart, labelSeries, categories, i, s.values[i] ?? 0))}</text>`);
      });
    },
  },
  histogram: {
    legend: false,
    draw(ctx) {
      const { chart, series, parts } = ctx;
      const vals = series.flatMap((s) => s.values);
      const bins = histogramBins(vals, chart.binCount ?? null, chart.binWidth ?? null);
      const cats = bins.map(binLabel);
      const { area, pos, band, mid, base, clip } = cartesian(ctx, bins.map((b) => b.count), { cats });
      const s = series[0], w = specialBarWidth(band, chart, 0);
      parts.push(`<g data-plot="histogram"${clip}>`);
      bins.forEach((b, i) => {
        const x = mid(i) - w / 2, y = pos(b.count);
        parts.push(`<rect x="${(x + .5).toFixed(1)}" y="${Math.min(y, base).toFixed(1)}" width="${Math.max(1, w - 1).toFixed(1)}" height="${Math.abs(base - y).toFixed(1)}" fill="${s?.pointColors?.[i] ?? s?.color}" data-s="${s?._fi ?? 0}" data-p="${i}" data-value="${b.count}"/>`);
      });
      parts.push('</g>');
      if (ctx.wantLabels(s ?? {})) bins.forEach((b, i) => labelTxt(ctx, s, b.count, mid(i), Math.max(area.y + 12, Math.min(pos(b.count), base) - 4)));
    },
  },
  pareto: {
    draw(ctx) {
      const { chart, series, categories, parts, pal } = ctx, s = series[0];
      if (!s) return;
      const items = paretoItems(series, categories, chart), cats = items.map((x) => x.c);
      const total = items.reduce((sum, p) => sum + Math.max(0, p.v), 0) || 1;
      const { area, pos, pos2, band, mid, base, clip } = cartesian(ctx, items.map((p) => p.v), { cats, code: s.numFmt, secondary: { values: [0, 1], defaults: { min: 0, max: 1, major: .25 }, code: '0%' } });
      const w = specialBarWidth(band, chart, 11), pts = [];
      let acc = 0;
      parts.push(`<g data-plot="pareto"${clip}>`);
      items.forEach((p, i) => {
        const x = mid(i) - w / 2, y = pos(p.v);
        parts.push(`<rect x="${x.toFixed(1)}" y="${Math.min(y, base).toFixed(1)}" width="${w.toFixed(1)}" height="${Math.abs(base - y).toFixed(1)}" fill="${s.pointColors?.[p.i] ?? s.color}" data-s="${s._fi}" data-p="${p.i}" data-value="${p.v}"/>`);
        acc += Math.max(0, p.v); pts.push([mid(i), pos2(acc / total)]);
      });
      parts.push(`<path data-pareto="cumulative" d="${pts.map((p, i) => `${i ? 'L' : 'M'}${p[0].toFixed(1)},${p[1].toFixed(1)}`).join('')}" fill="none" stroke="${pal[1 % pal.length]}" stroke-width="2.25"/>`);
      for (const p of pts) parts.push(`<circle cx="${p[0].toFixed(1)}" cy="${p[1].toFixed(1)}" r="2.5" fill="${pal[1 % pal.length]}"/>`);
      parts.push('</g>');
      if (ctx.wantLabels(s)) items.forEach((p, i) => labelTxt(ctx, s, p.v, mid(i), Math.max(area.y + 12, Math.min(pos(p.v), base) - 4)));
    },
  },
  treemap: { legendItems: hierarchyLegend, draw: drawHierarchyTreemap },
  boxWhisker: {
    draw(ctx) {
      const { chart, series, parts } = ctx;
      const stats = series.map((s) => { const st = boxWhiskerStats(s.values, chart.quartileMethod); return st ? { ...st, s } : null; });
      // 표시하지 않는 이상값도 원자료 범위에 포함: 옵션 토글로 축이 바뀌어 분포가 왜곡되지 않습니다.
      const all = stats.filter(Boolean).flatMap((s) => [s.min, s.max]);
      const { pos, band, mid, clip } = cartesian(ctx, all, { cats: series.map((s) => s.name), zero: false, code: series[0]?.numFmt });
      parts.push(`<g data-plot="boxWhisker"${clip}>`);
      stats.forEach((st, i) => {
        if (!st) return;
        const cx = mid(i), w = specialBarWidth(band, chart, 100), c = st.s.color, attrs = ` data-s="${st.s._fi}"`;
        parts.push(`<line x1="${cx}" y1="${pos(st.hi).toFixed(1)}" x2="${cx}" y2="${pos(st.q3).toFixed(1)}" stroke="#595959"${attrs}/>`, `<line x1="${cx}" y1="${pos(st.q1).toFixed(1)}" x2="${cx}" y2="${pos(st.lo).toFixed(1)}" stroke="#595959"${attrs}/>`);
        for (const v of [st.hi, st.lo]) parts.push(`<line x1="${cx - w / 4}" y1="${pos(v).toFixed(1)}" x2="${cx + w / 4}" y2="${pos(v).toFixed(1)}" stroke="#595959"${attrs}/>`);
        parts.push(`<rect data-box="quartiles" data-q1="${st.q1}" data-q3="${st.q3}" x="${(cx - w / 2).toFixed(1)}" y="${Math.min(pos(st.q1), pos(st.q3)).toFixed(1)}" width="${w.toFixed(1)}" height="${Math.max(1, Math.abs(pos(st.q1) - pos(st.q3))).toFixed(1)}" fill="${c}" stroke="#595959" stroke-width=".75"${attrs}/>`);
        parts.push(`<line data-box="median" data-value="${st.med}" x1="${cx - w / 2}" y1="${pos(st.med).toFixed(1)}" x2="${cx + w / 2}" y2="${pos(st.med).toFixed(1)}" stroke="#fff" stroke-width="1.5"${attrs}/>`);
        if (chart.showMean !== false) { const my = pos(st.mean); parts.push(`<path data-box="mean" data-value="${st.mean}" d="M${cx - 3.5},${my - 3.5}L${cx + 3.5},${my + 3.5}M${cx + 3.5},${my - 3.5}L${cx - 3.5},${my + 3.5}" stroke="#404040" stroke-width="1.5"${attrs}/>`); }
        if (chart.showInnerPoints) for (const v of st.inside) parts.push(`<circle data-box="inner" data-value="${v}" cx="${cx}" cy="${pos(v).toFixed(1)}" r="2.5" fill="${c}" stroke="#fff"${attrs}/>`);
        if (chart.showOutliers !== false) for (const v of st.out) parts.push(`<circle data-box="outlier" data-value="${v}" cx="${cx}" cy="${pos(v).toFixed(1)}" r="2.5" fill="none" stroke="${c}"${attrs}/>`);
      });
      parts.push('</g>');
    },
  },
  stock: {
    legend: false,
    draw(ctx) {
      const { chart, series, categories, parts } = ctx;
      if (chart.volume) { drawVolumeStock(ctx); return; }
      const ohlc = chart.ohlc ?? series.length >= 4;
      const [o, h, l, c] = ohlc ? series : [null, ...series];
      if (!h || !l || !c || ohlc && !o) return;
      const vals = [...h.values, ...l.values, ...(c?.values ?? []), ...(o?.values ?? [])].filter(isNum);
      const { pos, band, mid, clip } = cartesian(ctx, vals, { cats: categories, zero: false, code: c?.numFmt });
      parts.push(`<g data-plot="stock"${clip}>`);
      categories.forEach((_, i) => {
        const cx = mid(i);
        const hi = h.values[i];
        const lo = l.values[i];
        if (isNum(hi) && isNum(lo)) parts.push(`<line x1="${cx.toFixed(1)}" y1="${pos(hi).toFixed(1)}" x2="${cx.toFixed(1)}" y2="${pos(lo).toFixed(1)}" stroke="#404040"/>`);
        if (ohlc && isNum(o.values[i]) && isNum(c.values[i])) {
          const up = c.values[i] >= o.values[i];
          const y1 = Math.min(pos(o.values[i]), pos(c.values[i]));
          const y2 = Math.max(pos(o.values[i]), pos(c.values[i]));
          const w = Math.min(14, band * 0.5);
          parts.push(`<rect x="${(cx - w / 2).toFixed(1)}" y="${y1.toFixed(1)}" width="${w.toFixed(1)}" height="${Math.max(1, y2 - y1).toFixed(1)}" fill="${up ? (chart.upColor ?? '#ffffff') : (chart.downColor ?? '#404040')}" stroke="#404040"/>`);
        } else if (c && isNum(c.values[i])) {
          parts.push(`<line x1="${cx.toFixed(1)}" y1="${pos(c.values[i]).toFixed(1)}" x2="${(cx + Math.min(8, band * 0.3)).toFixed(1)}" y2="${pos(c.values[i]).toFixed(1)}" stroke="#404040" stroke-width="2"/>`);
        }
      });
      parts.push('</g>');
    },
  },
  radar: {
    draw(ctx) {
      const { chart, series, categories, plot, parts, FS, TXT, GRID } = ctx;
      const n = categories.length;
      if (n < 3) return;
      const cx = plot.x + plot.w / 2;
      const cy = plot.y + plot.h / 2 + 4;
      const R = Math.max(10, Math.min(plot.w / 2 - 40, plot.h / 2 - 16));
      const vals = series.flatMap((s) => s.values).filter(isNum);
      const sc = niceScale(Math.min(0, minOf(vals)), Math.max(...vals, 1), 4);
      const ang = (i) => -Math.PI / 2 + (i / n) * Math.PI * 2;
      const at = (i, v) => { const r = ((v - sc.min) / (sc.max - sc.min)) * R; return [cx + Math.cos(ang(i)) * r, cy + Math.sin(ang(i)) * r]; };
      for (let t = sc.min; t <= sc.max + sc.step / 2; t += sc.step) {
        const pts = categories.map((_, i) => at(i, t));
        parts.push(`<path d="${pts.map((p, i) => `${i ? 'L' : 'M'}${p[0].toFixed(1)},${p[1].toFixed(1)}`).join('')}Z" fill="none" stroke="${GRID}"/>`);
        parts.push(T(cx - 4, at(0, t)[1] + 3.5, axisLabel(Number(t.toPrecision(12))), FS.axis * 0.9, TXT, 'end'));
      }
      categories.forEach((c, i) => {
        const [x, y] = at(i, sc.max);
        parts.push(`<line x1="${cx}" y1="${cy}" x2="${x.toFixed(1)}" y2="${y.toFixed(1)}" stroke="${GRID}"/>`);
        const [lx, ly] = [cx + Math.cos(ang(i)) * (R + 12), cy + Math.sin(ang(i)) * (R + 12)];
        parts.push(T(lx, ly + 3.5, truncate(String(c), 12), FS.axis, TXT, Math.abs(Math.cos(ang(i))) < 0.2 ? 'middle' : Math.cos(ang(i)) > 0 ? 'start' : 'end'));
      });
      const filled = chart.radarStyle === 'filled';
      series.forEach((s) => {
        const pts = s.values.map((v, i) => at(i, isNum(v) ? v : sc.min));
        parts.push(`<path d="${pts.map((p, i) => `${i ? 'L' : 'M'}${p[0].toFixed(1)},${p[1].toFixed(1)}`).join('')}Z" fill="${filled ? s.color : 'none'}" fill-opacity="${filled ? 0.55 : 0}" stroke="${s.color}" stroke-width="${s.lineWidth ?? 2}"${dashAttr(s.dash, s.lineWidth ?? 2)} data-s="${s._fi}"/>`);
        const marker = s.marker ?? chart.marker;
        const showMarker = marker === undefined ? chart.radarStyle === 'marker' : marker !== false && marker !== 'none';
        if (showMarker) pts.forEach((p, i) => {
          if (!isNum(s.values[i])) return;
          const markup = (MARKERS[marker] ?? MARKERS.circle)(Number(p[0].toFixed(1)), Number(p[1].toFixed(1)), s.markerSize ? s.markerSize * 2 / 3 : 3, pointColor(s, i) ?? s.markerColor ?? s.color);
          parts.push(markup.replace(/^<(\w+)/, `<$1 data-s="${s._fi}" data-p="${chartPointIndex(s, i)}"`));
        });
        if (ctx.wantLabels(s)) s.values.forEach((v, i) => { if (isNum(v)) { const p = at(i, v); labelTxt(ctx, s, v, p[0], p[1] - 6); } });
      });
    },
  },
  bubble: {
    draw(ctx) {
      const { series, plot, parts, FS, TXT, GRID, chart } = ctx;
      const xs = series.flatMap((s) => (s.x ?? s.values.map((_, i) => i + 1))).filter(isNum);
      const ys = series.flatMap((s) => s.values).filter(isNum);
      const sizes = series.flatMap((s) => s.size ?? []).filter(isNum);
      if (!xs.length || !ys.length) return;
      const xc = chart.axes?.x ?? {}, yc = chart.axes?.y ?? {};
      const xsc = chartValueScale(xs, xc, { zero: false });
      const ysc = chartValueScale(ys, yc);
      const CW = FS.axis * 0.58;
      const yt = ysc.ticks;
      const lw = yc.hide ? 5 : maxOf(yt.map((t) => axisLabel(t, ysc.code, chart.date1904).length)) * CW + 8 + (yc.title ? FS.axis * 1.7 : 0);
      const area = { x: plot.x + lw, y: plot.y + 4, w: plot.w - lw - 10, h: plot.h - FS.axis * (xc.title ? 3.6 : 1.9) };
      if (area.w < 10 || area.h < 10) return;
      const maxR = Math.min(area.w, area.h) * Math.min(.25, Math.max(.01, .12 * ((chart.bubbleScale ?? 100) / 100)));
      area.x += maxR; area.y += maxR; area.w -= 2 * maxR; area.h -= 2 * maxR;
      const X = (v) => { const f=(v-xsc.min)/(xsc.max-xsc.min); return area.x+(xsc.reverse?1-f:f)*area.w; };
      const Y = (v) => { const f=(v-ysc.min)/(ysc.max-ysc.min); return area.y+(ysc.reverse?f:1-f)*area.h; };
      for (const t of yt) {
        if (chart.gridY !== false) parts.push(`<line x1="${area.x}" y1="${Y(t).toFixed(1)}" x2="${area.x + area.w}" y2="${Y(t).toFixed(1)}" stroke="${GRID}"/>`);
        if (!yc.hide) ctx.axisText('y', T(area.x - 5, Y(t) + 3.5, axisLabel(t, ysc.code, chart.date1904), FS.axis, TXT, 'end'));
      }
      if (!xc.hide) for (const t of xsc.ticks) ctx.axisText('x', T(X(t), area.y + area.h + FS.axis * 1.35, axisLabel(t, xsc.code, chart.date1904), FS.axis, TXT));
      if (!xc.hide && xc.title) ctx.axisText('x', T(area.x+area.w/2,plot.y+plot.h,xc.title,FS.axis,TXT));
      if (!yc.hide && yc.title) ctx.axisText('y', T(plot.x+FS.axis,area.y+area.h/2,yc.title,FS.axis,TXT,'middle',` transform="rotate(-90 ${plot.x+FS.axis} ${area.y+area.h/2})"`));
      const maxS = Math.max(maxOf(sizes.map(Math.abs)), 1);
      const clipId=`${ctx.uid}bubblePlot`;ctx.defs.push(`<clipPath id="${clipId}"><rect x="${area.x-maxR}" y="${area.y-maxR}" width="${area.w+maxR*2}" height="${area.h+maxR*2}"/></clipPath>`);parts.push(`<g clip-path="url(#${clipId})">`);
      series.forEach((s) => s.values.forEach((v, i) => {
        const x = s.x ? s.x[i] : i + 1;
        if (!isNum(v) || !isNum(x)) return;
        const rawSize = s.size?.[i] ?? 1;
        if (!isNum(rawSize) || rawSize === 0 || rawSize < 0 && !chart.showNegBubbles) return;
        const sz = Math.abs(rawSize), r = Math.sqrt(sz / maxS) * maxR;
        let fill = pointColor(s, i) ?? s.color;
        if (chart.threeD) {
          const id = `${ctx.uid}-bubble-${s._fi}-${i}`;
          ctx.defs.push(`<radialGradient id="${id}" cx="35%" cy="30%" r="70%"><stop offset="0" stop-color="#fff"/><stop offset="0.32" stop-color="${escSvg(fill)}"/><stop offset="1" stop-color="#333"/></radialGradient>`);
          fill = `url(#${id})`;
        }
        parts.push(`<circle cx="${X(x).toFixed(1)}" cy="${Y(v).toFixed(1)}" r="${r.toFixed(2)}" fill="${fill}" fill-opacity="${chart.threeD ? 1 : .75}" stroke="${s.color}" data-s="${s._fi}" data-p="${chartPointIndex(s, i)}"${chart.threeD ? ' data-3d="bubble"' : ''}/>`);
        if (ctx.wantLabels(s)) parts.push(T(X(x),Y(v)+3.5,chartDataLabel(chart,s,ctx.categories,i,v),(s.labelSize??9)*4/3,s.labelColor??'#404040','middle',` data-el="label" data-s="${s._fi}" data-p="${chartPointIndex(s,i)}"`));
      }));
      parts.push('</g>');
    },
  },
};

/** 스퀘어리파이 트리맵 배치 (items: [{ v, i }] 큰 값부터) */
function squarify(items, box) {
  const out = [];
  const total = items.reduce((a, x) => a + x.v, 0);
  if (!total) return out;
  const scale = (box.w * box.h) / total;
  let rest = items.map((x) => ({ ...x, a: x.v * scale }));
  let b = { ...box };
  while (rest.length) {
    const short = Math.min(b.w, b.h);
    let row = [rest[0]];
    const worst = (r) => { const s = r.reduce((a, x) => a + x.a, 0); const mx = maxOf(r.map((x) => x.a)); const mn = minOf(r.map((x) => x.a)); return Math.max((short * short * mx) / (s * s), (s * s) / (short * short * mn)); };
    let k = 1;
    while (k < rest.length && worst([...row, rest[k]]) <= worst(row)) { row.push(rest[k]); k++; }
    const sum = row.reduce((a, x) => a + x.a, 0);
    if (b.w >= b.h) {
      const w = sum / b.h;
      let y = b.y;
      for (const x of row) { const h = x.a / w; out.push({ x: b.x, y, w, h, v: x.v, i: x.i }); y += h; }
      b = { x: b.x + w, y: b.y, w: b.w - w, h: b.h };
    } else {
      const h = sum / b.w;
      let x0 = b.x;
      for (const x of row) { const w = x.a / h; out.push({ x: x0, y: b.y, w, h, v: x.v, i: x.i }); x0 += w; }
      b = { x: b.x, y: b.y + h, w: b.w, h: b.h - h };
    }
    rest = rest.slice(k);
  }
  return out;
}
