// 국가/지역 단위 오프라인 색칠지도. 세부 주소·행정구역 지오코딩은 하지 않습니다.
import { MAP_COUNTRIES } from './chart-map-data.js';
import { formatGeneral, formatCode } from './format.js';
const finite = n => typeof n === 'number' && Number.isFinite(n);
const esc = v => String(v ?? '').replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&apos;' })[c]);
const key = v => String(v ?? '').normalize('NFKD').replace(/[\u0300-\u036f]/g, '').toLowerCase().replace(/[\s._,'’()\-]/g, '');
const countries = new Map(MAP_COUNTRIES.map(c => [c.id, c]));
let lookup;
function countryLookup() {
  if (lookup) return lookup;
  lookup = new Map();
  const add = (name, id) => { if (!name) return; const k = key(name); if (!lookup.has(k)) lookup.set(k, id); else if (lookup.get(k) !== id) lookup.set(k, null); };
  let displayKo, displayEn;
  try { displayKo = new Intl.DisplayNames(['ko'], { type: 'region' }); displayEn = new Intl.DisplayNames(['en'], { type: 'region' }); } catch { /* 데이터에 저장한 한/영 이름으로 동작 */ }
  for (const c of MAP_COUNTRIES) {
    for (const name of [c.ko, c.en, ...c.aliases]) add(name, c.id);
    if (c.iso2) { add(displayKo?.of(c.iso2), c.id); add(displayEn?.of(c.iso2), c.id); }
  }
  // 언어 이름의 충돌과 관계없이 코드 자체는 해당 코드의 국가/지역을 뜻합니다.
  for (const c of MAP_COUNTRIES) for (const code of [c.id, c.iso2, c.iso3]) if (code) lookup.set(key(code), c.id);
  for (const [id, names] of [
    ['KOR', ['한국', '남한', '대한민국', 'Republic of Korea', 'Korea, Republic of']],
    ['PRK', ['북한', 'North Korea', 'DPRK']],
    ['USA', ['미국', '미합중국', 'United States', 'USA', 'U.S.A.']],
    ['GBR', ['영국', 'UK', 'U.K.', 'United Kingdom', 'Great Britain', 'Britain']],
    ['CHN', ['중국', '중화인민공화국']],
  ]) for (const name of names) lookup.set(key(name), id);
  return lookup;
}
export function mapCountry(value) { const id = countryLookup().get(key(value)); return id ? countries.get(id) : null; }
/** 같은 국가/지역의 여러 행은 합계. 빈 값/오류/문자열을 0으로 바꾸지 않습니다. */
export function mapValues(data) {
  const groups = new Map(), unknown = [], invalid = [], values = data.series?.[0]?.values ?? [];
  for (let i = 0; i < (data.categories?.length ?? 0); i++) {
    const label = String(data.categories[i] ?? ''), country = mapCountry(label), value = values[i];
    if (!country) { unknown.push({ label, index: i }); continue; }
    if (!finite(value)) { invalid.push({ label, index: i }); continue; }
    let item = groups.get(country.id);
    if (!item) { item = { country, value: 0, points: [], count: 0 }; groups.set(country.id, item); }
    const sum = item.value === null ? null : item.value + value;
    if (sum === null || !finite(sum)) { item.value = null; invalid.push({ label, index: i, reason: '합계 범위 초과' }); }
    else item.value = sum;
    item.points.push(i); item.count++;
  }
  const items = [...groups.values()].filter(item => finite(item.value));
  let min = Infinity, max = -Infinity;
  for (const item of items) { min = Math.min(min, item.value); max = Math.max(max, item.value); }
  return { items, unknown, invalid, min: items.length ? min : 0, max: items.length ? max : 0, extraSeries: Math.max(0, (data.series?.length ?? 0) - 1) };
}
const hex = (value, fallback) => /^#[a-f\d]{6}$/i.test(String(value ?? '')) ? value : fallback;
export function mapColor(value, min, max, low = '#dbe9f6', high = '#1764ab', middle = '#ffffff') {
  if (!finite(value)) return '#e5e7eb';
  low = hex(low, '#dbe9f6'); high = hex(high, '#1764ab'); middle = hex(middle, '#ffffff');
  let from = low, to = high, t;
  if (min < 0 && max > 0) { from = value < 0 ? low : middle; to = value < 0 ? middle : high; t = value < 0 ? 1 - value / min : value / max; }
  else t = max > min ? (value - min) / (max - min) : .5;
  t = Math.max(0, Math.min(1, t));
  return '#' + [1, 3, 5].map(at => Math.round(parseInt(from.slice(at, at + 2), 16) * (1 - t) + parseInt(to.slice(at, at + 2), 16) * t).toString(16).padStart(2, '0')).join('');
}
function drawMap(ctx) {
  const { chart, data, series, plot, parts, defs, uid, TXT } = ctx;
  const result = mapValues(data), found = new Map(result.items.map(item => [item.country.id, item]));
  const s = series[0] ?? {}, seriesIndex = Number.isInteger(s._fi) && s._fi >= 0 ? s._fi : 0, low = hex(chart.mapLowColor, result.min < 0 && result.max > 0 ? '#d6614d' : '#dbe9f6'), high = hex(chart.mapHighColor, hex(s.color, '#1764ab')), middle = hex(chart.mapMidColor, '#ffffff');
  const color = value => mapColor(value, result.min, result.max, low, high, middle);
  const valueText = value => s.numFmt ? formatCode(value, s.numFmt).text : formatGeneral(value);
  const notes = [];
  if (result.unknown.length) notes.push(`인식 못한 지역 ${result.unknown.length}개`);
  if (result.invalid.length) notes.push(`숫자 아닌 값·합계 오류 ${result.invalid.length}개`);
  if (result.extraSeries) notes.push('첫 번째 값 계열 사용');
  const footer = notes.length ? 43 : 28, legend = chart.legend !== 'none' && result.items.length > 0;
  const legendH = legend ? 30 : 0;
  const width = Math.max(20, plot.w - 12), height = Math.max(20, plot.h - footer - legendH), k = Math.min(width / 360, height / 180);
  const x = plot.x + (plot.w - 360 * k) / 2, y = plot.y + (height - 180 * k) / 2;
  parts.push(`<g data-map="countries" transform="translate(${x},${y}) scale(${k})">`);
  for (const c of MAP_COUNTRIES) {
    const item = found.get(c.id), fill = item ? color(item.value) : '#e5e7eb';
    const label = `${c.ko} (${c.iso2 || c.id}): ${item ? valueText(item.value) + (item.count > 1 ? ` (${item.count}개 행 합계)` : '') : '데이터 없음'}`;
    const attrs = ` data-country="${esc(c.id)}"${item ? ` data-s="${seriesIndex}" data-p="${item.points[0]}" data-value="${item.value}"` : ''}`;
    if (c.path) parts.push(`<path d="${c.path}" fill="${fill}" fill-rule="evenodd" stroke="#ffffff" stroke-width="${.45 / k}"${attrs}><title>${esc(label)}</title></path>`);
  }
  // 낮은 해상도 세계 경계에 없는 작은 국가/지역은 점으로 구분해 표시합니다.
  for (const c of MAP_COUNTRIES) for (const [lon, lat] of c.points) {
    const item = found.get(c.id), label = `${c.ko} (${c.iso2 || c.id}): ${item ? valueText(item.value) : '데이터 없음'} · 작은 국가/지역 위치 표시`;
    parts.push(`<circle cx="${lon + 180}" cy="${90 - lat}" r="${(item ? 3 : 1.7) / k}" fill="${item ? color(item.value) : '#e5e7eb'}" stroke="#6b7280" stroke-width="${.6 / k}" data-map-point="true" data-country="${esc(c.id)}"${item ? ` data-s="${seriesIndex}" data-p="${item.points[0]}" data-value="${item.value}"` : ''}><title>${esc(label)}</title></circle>`);
  }
  parts.push('</g>');
  const labelText = (tx, ty, value, anchor = 'middle', size = 10, attrs = '') => `<text x="${tx}" y="${ty}" text-anchor="${anchor}" font-size="${size}" fill="${esc(TXT)}"${attrs}>${esc(value)}</text>`;
  if (legend) {
    const zeroOffset = result.min < 0 && result.max > 0 ? 1 / (1 + result.max / -result.min) : 0;
    const id = uid + '-map', lw = Math.min(200, plot.w * .5), lx = plot.x + (plot.w - lw) / 2, ly = plot.y + height + 5;
    const stops = result.min === result.max ? [[0, color(result.min)], [1, color(result.max)]] : result.min < 0 && result.max > 0 ? [[0, low], [zeroOffset, middle], [1, high]] : [[0, low], [1, high]];
    defs.push(`<linearGradient id="${id}">${stops.map(([offset, color]) => `<stop offset="${offset}" stop-color="${color}"/>`).join('')}</linearGradient>`);
    parts.push(`<g data-el="legend"><rect x="${lx}" y="${ly}" width="${lw}" height="9" fill="url(#${id})"/>`, labelText(lx, ly + 23, valueText(result.min), 'start'), labelText(lx + lw, ly + 23, valueText(result.max), 'end'));
    if (result.min < 0 && result.max > 0) parts.push(labelText(lx + lw * (zeroOffset), ly + 23, '0'));
    parts.push('</g>');
  }
  if (!result.items.length) parts.push(labelText(plot.x + plot.w / 2, plot.y + Math.max(16, height / 2), '국가/지역 이름과 숫자 값을 확인하세요.', 'middle', 12));
  const detail = [...result.unknown.map(u => u.label || '(빈 항목)'), ...result.invalid.map(u => u.label + (u.reason ? ': ' + u.reason : ''))].join(', ');
  const bottom = plot.y + plot.h;
  if (notes.length) parts.push(`<g data-map="unmatched"><title>${esc(detail)}</title>${labelText(plot.x + plot.w / 2, bottom - 22, notes.join(' · '), 'middle', 10)}</g>`);
  parts.push(labelText(plot.x + plot.w / 2, bottom - 7, '국가/지역 경계: Natural Earth · 작은 지역은 점 표시', 'middle', Math.min(10, Math.max(7, plot.w / 45))));
}
export const MAP_CHARTS = { map: { legend: false, legendItems: () => [], draw: drawMap } };
