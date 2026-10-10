// 리본 갤러리 (테마 · 전환 · 애니메이션 · 도형 스타일 · WordArt · 그림/표 스타일) 와 타이밍 입력 칸
import { S, curSlide, selObjects, selOne, run } from './state.js';
import { THEMES, cloneTheme, resolveColor } from './themes.js';
import { slideHtml, fillCss } from './render.js';
import { newPresentation, TRANSITIONS, ANIM_EFFECTS, plainText } from './model.js';
import { shapePath, SHAPE_GALLERY } from './shapes.js';
import { SHAPE_STYLES, WORDART, PICTURE_STYLES, TABLE_STYLES } from './presets.js';
import { el } from './ui.js';

const themeCache = new Map();
/** 테마 미리 보기 (제목 슬라이드) */
export function themePreview(t, w = 96) {
  const key = `${t.name}|${w}|${S.pres.size.w}x${S.pres.size.h}`;
  if (themeCache.has(key)) return themeCache.get(key);
  const p = newPresentation({ theme: t.name });
  p.size = { ...S.pres.size };
  p.theme = cloneTheme(t);
  const s = p.slides[0];
  s.objects[0].text.paras[0].runs = [{ t: '가나다 Aa' }];
  s.objects[1].text.paras[0].runs = [];
  const sc = w / p.size.w;
  const html = `<div class="tw" style="width:${w}px;height:${Math.round(p.size.h * sc)}px"><div class="tw-in" style="transform:scale(${sc})">${slideHtml(p, s, {})}</div></div>`;
  themeCache.set(key, html);
  return html;
}

function strip(items, { cls = '' } = {}) {
  return el('div', { class: `gallery ${cls}` }, el('div', { class: 'gallery-items' }, items));
}

const TRANS_GLYPH = { none: '∅', cut: '▮', fade: '◐', push: '⇧', wipe: '▤', split: '⇔', cover: '⬒', uncover: '⬓', zoom: '⊕', circle: '◯', dissolve: '░', flip: '⇋' };
const ANIM_GLYPH = { appear: '★', fade: '✧', fly: '➶', float: '⇡', split: '⇔', wipe: '▤', zoom: '⊕', grow: '↻', wheel: '◔', bounce: '⤵' };

export const GALLERIES = {
  themes: () => {
    const cur = S.pres.theme.name;
    return strip(THEMES.map((t) => el('button', { class: `gal-item theme${t.name === cur ? ' on' : ''}`, title: t.label, html: themePreview(t), onclick: () => run('applyTheme', t.name) })), { cls: 'themes' });
  },
  transitions: () => {
    const cur = curSlide()?.transition?.type ?? 'none';
    return strip(TRANSITIONS.map(([k, l]) => el('button', { class: `gal-item fx${k === cur ? ' on' : ''}`, title: l, onclick: () => run('setTransition', k) }, el('span', { class: 'fx-ic' }, TRANS_GLYPH[k] ?? '◇'), el('span', { class: 'fx-l' }, l))), { cls: 'fx' });
  },
  animations: () => {
    const o = selObjects()[0];
    const a = o ? curSlide().anims?.find((x) => x.obj === o.id) : null;
    const items = [el('button', { class: `gal-item fx${!a ? ' on' : ''}`, title: '없음', disabled: !o, onclick: () => run('removeAnim', { all: true }) }, el('span', { class: 'fx-ic' }, '∅'), el('span', { class: 'fx-l' }, '없음'))];
    for (const [k, l] of ANIM_EFFECTS.entr) items.push(el('button', { class: `gal-item fx entr${a?.cls === 'entr' && a.effect === k ? ' on' : ''}`, title: `나타내기: ${l}`, disabled: !o, onclick: () => run('setAnim', 'entr', k) }, el('span', { class: 'fx-ic' }, ANIM_GLYPH[k] ?? '★'), el('span', { class: 'fx-l' }, l)));
    for (const [k, l] of ANIM_EFFECTS.emph) items.push(el('button', { class: `gal-item fx emph${a?.cls === 'emph' && a.effect === k ? ' on' : ''}`, title: `강조: ${l}`, disabled: !o, onclick: () => run('setAnim', 'emph', k) }, el('span', { class: 'fx-ic' }, '✷'), el('span', { class: 'fx-l' }, l)));
    for (const [k, l] of ANIM_EFFECTS.exit) items.push(el('button', { class: `gal-item fx exit${a?.cls === 'exit' && a.effect === k ? ' on' : ''}`, title: `끝내기: ${l}`, disabled: !o, onclick: () => run('setAnim', 'exit', k) }, el('span', { class: 'fx-ic' }, '✦'), el('span', { class: 'fx-l' }, l)));
    return strip(items, { cls: 'fx' });
  },
  shapesMini: () => {
    const list = SHAPE_GALLERY.flatMap(([, l]) => l).slice(0, 24);
    const grid = el('div', { class: 'shape-mini' }, list.map(([k, l]) => el('button', { class: 'shape-btn', title: l, html: shapeIconSvg(k), onclick: () => run('drawShape', k) })));
    return el('div', { class: 'gallery shapes' }, grid, el('button', { class: 'gal-more', title: '모든 도형', onclick: (e) => run('shapesMenu', e.currentTarget) }, '▾'));
  },
  shapeStyles: () => {
    const items = SHAPE_STYLES.slice(0, 14).map((st, i) => styleChip(st, i));
    return el('div', { class: 'gallery styles' }, el('div', { class: 'gallery-items' }, items), el('button', { class: 'gal-more', title: '자세히', onclick: (e) => run('quickStylesMenu', e.currentTarget) }, '▾'));
  },
  wordart: () => strip(WORDART.map((w, i) => el('button', { class: 'gal-item wa', title: w.label, onclick: () => run('applyWordArt', i), html: `<span style="${wordArtCss(w)}">가</span>` })), { cls: 'wa' }),
  pictureStyles: () => strip(PICTURE_STYLES.map((p, i) => el('button', { class: 'gal-item pic', title: p.label, onclick: () => run('applyPictureStyle', i), html: picStyleSvg(p) })), { cls: 'pic' }),
  tableStyles: () => {
    const t = selOne();
    const cur = t?.type === 'table' ? t.style : null;
    return strip(TABLE_STYLES.map((ts, i) => el('button', { class: `gal-item tblst${cur && ((ts.kind === 'none' && cur.none) || (!cur.none && cur.accent === ts.accent && !!cur.light === (ts.kind === 'light'))) ? ' on' : ''}`, title: ts.label, onclick: () => run('applyTableStyle', i), html: tableStyleSvg(ts) })), { cls: 'tbl' });
  },
  transTiming: () => {
    const s = curSlide();
    const t = s?.transition ?? {};
    const dur = el('input', { type: 'number', step: '0.25', min: '0', value: String(t.dur ?? 1), class: 'rnum' });
    dur.addEventListener('change', () => run('transitionDuration', Number(dur.value)));
    const clk = el('input', { type: 'checkbox', checked: t.advClick !== false });
    clk.addEventListener('change', () => run('transitionAdvance', { advClick: clk.checked }));
    const aft = el('input', { type: 'checkbox', checked: t.advAfter != null });
    const sec = el('input', { type: 'number', step: '0.5', min: '0', value: String(t.advAfter ?? 5), class: 'rnum' });
    aft.addEventListener('change', () => run('transitionAdvance', { advAfter: aft.checked ? Number(sec.value) : null }));
    sec.addEventListener('change', () => { if (aft.checked) run('transitionAdvance', { advAfter: Number(sec.value) }); });
    for (const i of [dur, sec]) i.addEventListener('keydown', (e) => e.stopPropagation());
    return el('div', { class: 'rcol timing' },
      el('label', { class: 'rline' }, el('span', {}, '기간:'), dur, el('small', {}, '초')),
      el('label', { class: 'rline' }, clk, el('span', {}, '마우스를 클릭할 때')),
      el('label', { class: 'rline' }, aft, el('span', {}, '다음 시간 후:'), sec, el('small', {}, '초')));
  },
  animTiming: () => {
    const o = selObjects()[0];
    const a = o ? curSlide().anims?.find((x) => x.obj === o.id) : null;
    const start = el('select', { class: 'rselect', disabled: !a }, [['click', '클릭할 때'], ['with', '이전 효과와 함께'], ['after', '이전 효과 다음에']].map(([v, l]) => el('option', { value: v, selected: a?.start === v }, l)));
    start.addEventListener('change', () => run('animTiming', { start: start.value }));
    const dur = el('input', { type: 'number', step: '0.25', min: '0', value: String(a?.dur ?? 0.5), class: 'rnum', disabled: !a });
    dur.addEventListener('change', () => run('animTiming', { dur: Number(dur.value) }));
    const delay = el('input', { type: 'number', step: '0.25', min: '0', value: String(a?.delay ?? 0), class: 'rnum', disabled: !a });
    delay.addEventListener('change', () => run('animTiming', { delay: Number(delay.value) }));
    for (const i of [dur, delay]) i.addEventListener('keydown', (e) => e.stopPropagation());
    return el('div', { class: 'rcol timing' },
      el('label', { class: 'rline' }, el('span', {}, '시작:'), start),
      el('label', { class: 'rline' }, el('span', {}, '재생 시간:'), dur, el('small', {}, '초')),
      el('label', { class: 'rline' }, el('span', {}, '지연:'), delay, el('small', {}, '초'),
        el('button', { class: 'rbtn small', title: '앞으로 이동', disabled: !a, onclick: () => run('moveAnim', -1) }, '▲'),
        el('button', { class: 'rbtn small', title: '뒤로 이동', disabled: !a, onclick: () => run('moveAnim', 1) }, '▼')));
  },
  objSize: () => {
    const o = selObjects()[0];
    const cm = (px) => String(Math.round(px * 2.54 / 96 * 100) / 100);
    const h = el('input', { type: 'number', step: '0.1', class: 'rnum wide', value: o ? cm(o.h) : '', disabled: !o });
    const w = el('input', { type: 'number', step: '0.1', class: 'rnum wide', value: o ? cm(o.w) : '', disabled: !o });
    h.addEventListener('change', () => run('setSize', { h: Number(h.value) * 96 / 2.54 }));
    w.addEventListener('change', () => run('setSize', { w: Number(w.value) * 96 / 2.54 }));
    for (const i of [h, w]) i.addEventListener('keydown', (e) => e.stopPropagation());
    return el('div', { class: 'rcol timing' },
      el('label', { class: 'rline' }, el('span', { 'data-icon': 'rowInsert' }), el('span', {}, '높이:'), h, el('small', {}, 'cm')),
      el('label', { class: 'rline' }, el('span', { 'data-icon': 'colInsert' }), el('span', {}, '너비:'), w, el('small', {}, 'cm')));
  },
};

export function shapeIconSvg(kind, size = 18) {
  const line = kind === 'line' || kind === 'straightConnector1' || kind === 'arc' || kind === 'bracketPair' || kind === 'bracePair';
  const d = shapePath(kind, 14, kind === 'line' || kind === 'straightConnector1' ? 14 : 12);
  return `<svg viewBox="-1 -1 16 16" width="${size}" height="${size}"><path d="${d}" fill="${line ? 'none' : '#dbe6f4'}" stroke="#4a6a94" stroke-width="1" vector-effect="non-scaling-stroke"/>${kind === 'straightConnector1' ? '<path d="M14 14l-4-1 3-3z" fill="#4a6a94"/>' : ''}</svg>`;
}

export function styleChip(st, i) {
  const t = S.pres.theme;
  const bg = fillCss(t, st.fill);
  const border = st.line ? `2px solid ${resolveColor(t, st.line.color)}` : '2px solid transparent';
  return el('button', { class: 'gal-item style', title: st.label, style: { background: bg, border, color: resolveColor(t, st.text) }, onclick: () => run('applyQuickStyle', i) }, 'Abc');
}

export function wordArtCss(w) {
  const t = S.pres.theme;
  const r = w.run;
  return [`color:${resolveColor(t, r.color ?? '@tx1')}`, r.b ? 'font-weight:700' : '', r.shadow ? 'text-shadow:1px 1.5px 2px rgba(0,0,0,.45)' : '', r.outline ? `-webkit-text-stroke:${Math.min(1.2, r.outline.width)}px ${resolveColor(t, r.outline.color)}` : ''].filter(Boolean).join(';');
}

function picStyleSvg(p) {
  const s = p.set;
  const path = shapePath(s.shape ?? 'rect', 44, 30);
  const stroke = s.line ? resolveColor(S.pres.theme, s.line.color) : 'none';
  return `<svg viewBox="-4 -4 52 38" width="52" height="38"><defs><linearGradient id="pg" x1="0" y1="0" x2="1" y2="1"><stop offset="0" stop-color="#9cc3e6"/><stop offset="1" stop-color="#5b8c3a"/></linearGradient></defs>${s.shadow ? `<path d="${path}" transform="translate(2 2)" fill="rgba(0,0,0,.25)"/>` : ''}<path d="${path}" fill="url(#pg)" stroke="${stroke === '#FFFFFF' ? '#ddd' : stroke}" stroke-width="${s.line ? Math.min(4, s.line.width / 2) : 0}"/></svg>`;
}

function tableStyleSvg(ts) {
  const t = S.pres.theme;
  const acc = resolveColor(t, `@${ts.accent}`);
  const rows = [];
  for (let r = 0; r < 4; r++) {
    const fill = ts.kind === 'none' ? '#fff' : r === 0 ? acc : ts.kind === 'light' ? (r % 2 ? resolveColor(t, `@${ts.accent}:lm20:lo80`) : '#fff') : resolveColor(t, r % 2 ? `@${ts.accent}:lm20:lo80` : `@${ts.accent}:lm40:lo60`);
    rows.push(`<rect x="0" y="${r * 8}" width="44" height="8" fill="${fill}" stroke="${ts.kind === 'none' ? '#999' : ts.kind === 'light' ? acc : '#fff'}" stroke-width=".6"/>`);
  }
  return `<svg viewBox="0 0 44 32" width="46" height="34">${rows.join('')}<path d="M11 0v32M22 0v32M33 0v32" stroke="${ts.kind === 'none' ? '#999' : '#fff'}" stroke-width=".6"/></svg>`;
}

export function slideTitleText(s) { const t = s.objects.find((o) => o.ph === 'title' || o.ph === 'ctrTitle'); return t ? plainText(t.text) : ''; }
