// 스포이트 (PowerPoint 의 [도형 채우기 › 스포이트]): 슬라이드의 한 점을 눌러 그 색을 선택한 개체에 적용
// 브라우저 EyeDropper API 가 없는 사파리 · 아이폰에서도 되도록, 화면에 그려진 요소의 색(그림은 픽셀)을 직접 읽음
import { S, run, selObjects } from './state.js';
import { el, toast } from './ui.js';

const TARGETS = { fill: ['shapeFill', '도형 채우기'], line: ['shapeOutline', '도형 윤곽선'], font: ['fontColor', '글꼴 색'] };
let tip = null;
let lastHex = null;
let pickCb = null;
const imgCache = new Map();

const hex2 = (n) => Math.round(n).toString(16).padStart(2, '0');
/** 'rgb(1, 2, 3)' / 'rgba(…, a)' / '#abc' → '#RRGGBB' (투명이면 null) */
export function cssToHex(c) {
  if (!c) return null;
  c = c.trim();
  if (c.startsWith('#')) {
    if (c.length === 4) return `#${c[1]}${c[1]}${c[2]}${c[2]}${c[3]}${c[3]}`.toUpperCase();
    return c.slice(0, 7).toUpperCase();
  }
  const m = c.match(/rgba?\(\s*([\d.]+)[ ,]+([\d.]+)[ ,]+([\d.]+)(?:[ ,/]+([\d.]+%?))?/);
  if (!m) return null;
  if (m[4] != null) { const a = m[4].endsWith('%') ? parseFloat(m[4]) / 100 : parseFloat(m[4]); if (a < 0.05) return null; }
  return `#${hex2(m[1])}${hex2(m[2])}${hex2(m[3])}`.toUpperCase();
}

function imageCanvas(img) {
  let cv = imgCache.get(img.src);
  if (cv) return cv;
  cv = document.createElement('canvas');
  cv.width = img.naturalWidth || 1;
  cv.height = img.naturalHeight || 1;
  try { cv.getContext('2d', { willReadFrequently: true }).drawImage(img, 0, 0); } catch { return null; }
  imgCache.set(img.src, cv);
  if (imgCache.size > 20) imgCache.delete(imgCache.keys().next().value);
  return cv;
}

function gradientStop(svgEl, prop) {
  const v = getComputedStyle(svgEl)[prop];
  const m = v && v.match(/url\(["']?#([^"')]+)/);
  if (!m) return null;
  const g = svgEl.ownerSVGElement?.getElementById?.(m[1]) ?? document.getElementById(m[1]);
  const stop = g?.querySelector('stop');
  return stop ? cssToHex(stop.getAttribute('stop-color') || getComputedStyle(stop).stopColor) : null;
}

/** 화면 좌표의 색 (슬라이드 안만) */
export function sampleAt(x, y, target = 'fill') {
  const layer = document.querySelector('#stage .slide-layer');
  if (!layer) return null;
  for (const e of document.elementsFromPoint(x, y)) {
    if (!layer.contains(e)) continue;
    if (target === 'font' && e.childNodes.length && [...e.childNodes].some((n) => n.nodeType === 3 && n.textContent.trim())) {
      const c = cssToHex(getComputedStyle(e).color);
      if (c) return c;
    }
    if (e.tagName === 'IMG') {
      const cv = imageCanvas(e);
      if (!cv) continue;
      const r = e.getBoundingClientRect();
      const px = Math.min(cv.width - 1, Math.max(0, Math.floor((x - r.left) / r.width * cv.width)));
      const py = Math.min(cv.height - 1, Math.max(0, Math.floor((y - r.top) / r.height * cv.height)));
      try {
        const d = cv.getContext('2d').getImageData(px, py, 1, 1).data;
        if (d[3] > 12) return `#${hex2(d[0])}${hex2(d[1])}${hex2(d[2])}`.toUpperCase();
      } catch { /* 다른 출처 그림 */ }
      continue;
    }
    if (e instanceof SVGElement && e.tagName !== 'svg' && e.tagName !== 'g') {
      const cs = getComputedStyle(e);
      const f = cs.fill && cs.fill !== 'none' ? (cs.fill.startsWith('url') ? gradientStop(e, 'fill') : cssToHex(cs.fill)) : null;
      if (f && parseFloat(cs.fillOpacity || '1') > 0.05) return f;
      const s = cs.stroke && cs.stroke !== 'none' ? cssToHex(cs.stroke) : null;
      if (s && e.tagName === 'line') return s;
      continue;
    }
    if (e instanceof HTMLElement) {
      const cs = getComputedStyle(e);
      const bg = cssToHex(cs.backgroundColor);
      if (bg) return bg;
      const gi = cs.backgroundImage;
      if (gi && gi.includes('gradient')) { const m = gi.match(/(rgba?\([^)]*\)|#[0-9a-f]{3,8})/i); const c = m && cssToHex(m[1]); if (c) return c; }
    }
  }
  return '#FFFFFF';
}

function onMove(e) {
  if (!tip) return;
  const c = sampleAt(e.clientX, e.clientY, S.eyedrop);
  lastHex = c;
  tip.style.left = `${e.clientX + 16}px`;
  tip.style.top = `${e.clientY + 16}px`;
  tip.hidden = !c;
  if (c) { tip.firstChild.style.background = c; tip.lastChild.textContent = `${c}  RGB(${parseInt(c.slice(1, 3), 16)}, ${parseInt(c.slice(3, 5), 16)}, ${parseInt(c.slice(5, 7), 16)})`; }
}
function onDown(e) {
  const stage = document.getElementById('stage');
  if (!stage.contains(e.target)) { stopEyedrop(); return; }
  e.preventDefault();
  e.stopPropagation();
  const c = sampleAt(e.clientX, e.clientY, S.eyedrop) ?? lastHex;
  const t = S.eyedrop;
  const cb = pickCb;
  stopEyedrop();
  if (!c) return;
  if (cb) cb(c);
  else { run(TARGETS[t][0], c); toast(`${TARGETS[t][1]}: ${c}`); }
}
function onKey(e) { if (e.key === 'Escape') { e.preventDefault(); e.stopPropagation(); stopEyedrop(); } }

/** target: 'fill' | 'line' | 'font'; onPick 이 있으면 색만 넘겨 줌 (색 메뉴의 [스포이트]) */
export function startEyedrop(target = 'fill', onPick = null) {
  if (!onPick && !selObjects().length && !(target === 'font' && S.editing)) { toast('먼저 색을 바꿀 개체를 선택하세요'); return; }
  stopEyedrop();
  pickCb = onPick;
  S.eyedrop = target;
  document.body.classList.add('eyedropping');
  tip = el('div', { class: 'eyedrop-tip', hidden: true }, el('i'), el('span'));
  document.body.append(tip);
  addEventListener('pointermove', onMove, true);
  addEventListener('pointerdown', onDown, true);
  addEventListener('keydown', onKey, true);
  toast(`스포이트 (${TARGETS[target][1]}): 슬라이드에서 원하는 색을 누르세요 · Esc 취소`);
}
export function stopEyedrop() {
  S.eyedrop = null;
  pickCb = null;
  document.body.classList.remove('eyedropping');
  tip?.remove();
  tip = null;
  removeEventListener('pointermove', onMove, true);
  removeEventListener('pointerdown', onDown, true);
  removeEventListener('keydown', onKey, true);
}
