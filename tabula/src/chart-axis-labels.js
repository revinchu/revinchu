// DOM 없는 범주축 배치. 자동 간격은 이름 길이가 아니라 실제 줄 높이/회전 공간으로 결정한다.
const esc = (s) => String(s).replace(/[&<>"']/g, (c) => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&apos;'}[c]));
const finite = (v) => typeof v === 'number' && Number.isFinite(v);

export function axisTextWidth(text, font = 12) {
  let width = 0;
  for (const c of String(text)) {
    if (/\p{Mark}/u.test(c)) continue;
    width += /[\u1100-\u11ff\u2e80-\ua4cf\uac00-\ud7af\uf900-\ufaff\uff00-\uffef]/u.test(c) || c.codePointAt(0) > 0xffff ? 1.04 : /[MW@#%&]/.test(c) ? .92 : /[ilI1.,:;!'| ]/.test(c) ? .34 : .62;
  }
  return width * font;
}

function wrap(text, width, font, limit) {
  const lines = [];
  for (const paragraph of String(text).split(/\r?\n/)) {
    let line = '';
    const add = (token) => {
      if (line && axisTextWidth(line + token, font) > width) { lines.push(line.trimEnd()); line = ''; }
      line += line ? token : token.trimStart();
    };
    for (const token of paragraph.match(/\s+|\S+/gu) ?? []) {
      if (axisTextWidth(token, font) <= width) add(token);
      else for (const c of token) add(c);
    }
    lines.push(line.trimEnd());
  }
  const complete = lines.length <= limit;
  if (!complete) {
    lines.length = Math.max(1, limit);
    let last = lines.at(-1);
    while (last && axisTextWidth(last + '…', font) > width) last = Array.from(last).slice(0, -1).join('');
    lines[lines.length - 1] = last + '…';
  }
  return { lines, complete, width: lines.reduce((max, s) => Math.max(max, axisTextWidth(s, font)), 0) };
}

export function categoryAxisLayout(categories, options = {}) {
  const font = Math.max(1, options.font ?? 12), lineHeight = font * 1.28;
  const length = Math.max(1, options.length ?? 400), depth = Math.max(lineHeight, options.depth ?? 120);
  const horizontal = !!options.horizontal, n = Math.max(categories.length, finite(options.slotCount) ? options.slotCount : 0), band = length / Math.max(1, n);
  const specified = finite(options.interval) && options.interval >= 1;
  // 긴 항목 하나 때문에 다른 항목까지 생략하지 않고, 최소 줄 높이만 확보한다.
  const every = specified ? Math.floor(options.interval) : Math.max(1, Math.ceil((horizontal ? lineHeight + 3 : font * 1.55) / band));
  const space = band * every, selected = [];
  categories.forEach((c, i) => { if (i % every === 0) selected.push({ index: i, text: String(c ?? '') }); });
  if (horizontal) {
    const angle = finite(options.rotation) ? Math.max(-90, Math.min(90, options.rotation)) : 0;
    const rad = Math.abs(angle) * Math.PI / 180, sin = Math.sin(rad), cos = Math.cos(rad);
    const limit = angle ? 1 : Math.max(1, Math.floor((space - 3) / lineHeight));
    const width = angle ? Math.max(font, Math.min((depth - 8 - lineHeight * sin) / Math.max(.001, cos), (space - 3 - lineHeight * cos) / sin)) : Math.max(font, depth - 8);
    const labels = selected.map((item) => ({ ...item, ...wrap(item.text, width, font, limit) }));
    return { labels, every, angle, font, lineHeight, horizontal, extent: labels.reduce((max, l) => Math.max(max, l.width * cos + l.lines.length * lineHeight * sin), 0) + 8 };
  }
  const manual = finite(options.rotation), angles = manual ? [Math.max(-90, Math.min(90, options.rotation))] : [0, -45, -90];
  const candidates = angles.map((angle) => {
    const rad = Math.abs(angle) * Math.PI / 180, sin = Math.sin(rad), cos = Math.cos(rad);
    const maxLines = angle === 0 ? Math.max(1, Math.min(manual ? 8 : 3, Math.floor((depth - 5) / lineHeight))) : Math.max(1, Math.min(3, Math.floor((space * sin - 4) / lineHeight)));
    const height = maxLines * lineHeight;
    const width = angle === 0 ? Math.max(font, space - 6) : Math.max(font, (depth - 5 - height * cos) / sin);
    const labels = selected.map((item) => {
      const visualIndex = options.reverse ? n - 1 - item.index : item.index;
      const mid = band * (visualIndex + .5);
      const room = angle < 0 ? mid + (options.startRoom ?? 0) : length - mid + (options.endRoom ?? 0);
      const allowed = manual && angle && cos > .001 ? Math.max(font, Math.min(width, (room - height * sin / 2) / cos)) : width;
      let wrapped = wrap(item.text, allowed, font, maxLines);
      if (angle && !manual && wrapped.lines.length > 1 && wrapped.complete) {
        const balanced = wrap(item.text, Math.min(allowed, axisTextWidth(item.text, font) / wrapped.lines.length * 1.1), font, maxLines);
        if (balanced.complete && balanced.lines.length <= wrapped.lines.length) wrapped = balanced;
      }
      return { ...item, ...wrapped };
    });
    let extent = lineHeight, fits = !angle || space * sin >= lineHeight + 3 || specified, complete = 0, startPadding = 0, endPadding = 0;
    for (const l of labels) {
      const h = l.lines.length * lineHeight;
      extent = Math.max(extent, l.width * sin + h * cos);
      const visualIndex = options.reverse ? n - 1 - l.index : l.index;
      const mid = band * (visualIndex + .5);
      const overhang = l.width * cos + h * sin / 2;
      if (angle < 0) startPadding = Math.max(startPadding, overhang - mid - (options.startRoom ?? 0));
      if (angle > 0) endPadding = Math.max(endPadding, overhang - length + mid - (options.endRoom ?? 0));
      if (l.complete) complete++;
    }
    const padding = Math.max(startPadding, endPadding);
    if (padding > (options.allowPadding ? Math.min(80, length * .22) : .1)) fits = false;
    return { labels, angle, extent: Math.ceil(extent + 7), fits: angle === 0 || fits, complete, startPadding, endPadding };
  });
  // 가로쓰기→기울임→세로 회전 순서. 공간이 부족하면 표시 가능한 글자를 최대한 유지한다.
  const full = candidates.find((c) => c.fits && c.complete === selected.length);
  const best = full ?? candidates.filter((c) => c.fits).sort((a, b) => b.complete - a.complete || b.labels.reduce((s, l) => s + l.lines.join('').length, 0) - a.labels.reduce((s, l) => s + l.lines.join('').length, 0))[0] ?? candidates[0];
  return { ...best, every, font, lineHeight, horizontal: false };
}

export function categoryAxisLabelSvg(plan, label, mid, edge, color = '#595959') {
  const { font, lineHeight, angle } = plan, h = label.lines.length * lineHeight;
  let transform, x = 0, y = font, anchor;
  if (plan.horizontal) {
    const rad = Math.abs(angle) * Math.PI / 180, sin = Math.sin(rad), cos = Math.cos(rad);
    const shiftX = angle < 0 ? h * sin : 0;
    const centerY = (h * cos + (angle < 0 ? 1 : -1) * label.width * sin) / 2;
    transform = `translate(${edge - 6 - shiftX} ${mid - centerY})${angle ? ` rotate(${angle})` : ''}`; anchor = 'end';
  }
  else if (angle) {
    const shift = h * Math.sin(Math.abs(angle) * Math.PI / 180) / 2;
    transform = `translate(${mid + (angle < 0 ? -shift : shift)} ${edge + 5}) rotate(${angle})`;
    anchor = angle < 0 ? 'end' : 'start';
  } else { transform = `translate(${mid} ${edge + 5})`; anchor = 'middle'; }
  const spans = label.lines.length === 1 ? esc(label.lines[0]) : label.lines.map((line, i) => `<tspan x="${x}" y="${y + i * lineHeight}">${esc(line)}</tspan>`).join('');
  return `<text data-axis="x" data-axis-label="${label.index}" aria-label="${esc(label.text)}" x="${x}" y="${y}" transform="${transform}" text-anchor="${anchor}" font-size="${font}" fill="${esc(color)}">${spans}</text>`;
}
