import { shortcutCode } from './keyboard-shortcuts.js';
// A document-independent draft editor. Workbook transactions and object ownership stay in app.js.
const shapeRunKeys = ['b', 'i', 'u', 's', 'sz', 'font', 'color'];
const shapeTextAliases = { bold: 'b', italic: 'i', underline: 'u', strike: 's', size: 'sz', fontSize: 'sz' };
const shapeTextCopy = value => JSON.parse(JSON.stringify(value));
const shapeRunAttrs = run => { const out = { ...run }; delete out.t; return out; };
const shapeParaAttrs = para => { const out = { ...para }; delete out.runs; return out; };
function shapeAppendRun(runs, run) {
  if (!run.t && runs.length) return;
  const prev = runs[runs.length - 1];
  if (prev && JSON.stringify(shapeRunAttrs(prev)) === JSON.stringify(shapeRunAttrs(run))) prev.t += run.t;
  else runs.push({ ...run });
}
export function shapeTextParagraphs(shape) {
  if (Array.isArray(shape.paras)) return shape.paras.length ? shape.paras.map(p => ({ ...shapeTextCopy(p), runs: (p.runs ?? []).map(r => ({ ...shapeTextCopy(r), t: String(r.t ?? '') })) })) : [{ runs: [] }];
  return String(shape.text ?? '').replace(/\r\n?/g, '\n').split('\n').map(t => ({ runs: t ? [{ t }] : [] }));
}
export function shapeTextString(paras) { return paras.map(p => p.runs.map(r => r.t).join('')).join('\n'); }
function shapeTextLength(para) { return para.runs.reduce((n, r) => n + r.t.length, 0); }
function shapeTextAt(paras, position) {
  let left = Math.max(0, Number(position) || 0);
  for (let p = 0; p < paras.length; p++) { const n = shapeTextLength(paras[p]); if (left <= n || p === paras.length - 1) return { p, offset: Math.min(left, n) }; left -= n + 1; }
  return { p: 0, offset: 0 };
}
function shapeRunSlice(runs, start, end) {
  const out = []; let at = 0;
  for (const r of runs) { const a = Math.max(0, start - at), b = Math.min(r.t.length, end - at); if (b > a) shapeAppendRun(out, { ...r, t: r.t.slice(a, b) }); at += r.t.length; }
  return out;
}
function shapeNormalizePatch(patch) {
  const out = {};
  for (const [key, value] of Object.entries(patch ?? {})) {
    const k = shapeTextAliases[key] ?? key;
    if (['b', 'i', 'u', 's'].includes(k)) out[k] = Boolean(value);
    else if (k === 'sz' && Number.isFinite(Number(value)) && Number(value) > 0) out.sz = Math.max(1, Math.min(4096, Number(value)));
    else if (['font', 'color'].includes(k) && typeof value === 'string') out[k] = value;
    else if (k === 'align' && ['left', 'center', 'right', 'justify'].includes(value)) out.align = value;
    else if (k === 'valign' && ['top', 'middle', 'bottom'].includes(value)) out.valign = value;
  }
  return out;
}
export function replaceShapeText(paras, selection, text, style = {}) {
  const a = shapeTextAt(paras, Math.min(selection.start, selection.end)), b = shapeTextAt(paras, Math.max(selection.start, selection.end));
  const parts = String(text).replace(/\r\n?/g, '\n').split('\n'), attrs = {};
  for (const k of shapeRunKeys) if (style[k] !== undefined) attrs[k] = style[k];
  const prefix = shapeRunSlice(paras[a.p].runs, 0, a.offset), suffix = shapeRunSlice(paras[b.p].runs, b.offset, Infinity), inserted = [];
  for (let i = 0; i < parts.length; i++) {
    const runs = i === 0 ? prefix : [];
    if (parts[i]) shapeAppendRun(runs, { ...attrs, t: parts[i] });
    if (i === parts.length - 1) for (const run of suffix) shapeAppendRun(runs, run);
    const original = i === 0 || parts.length === 1 ? paras[a.p] : i === parts.length - 1 ? paras[b.p] : paras[a.p];
    inserted.push({ ...shapeParaAttrs(original), runs });
  }
  return [...paras.slice(0, a.p).map(shapeTextCopy), ...inserted, ...paras.slice(b.p + 1).map(shapeTextCopy)];
}
export function formatShapeText(paras, selection, patch) {
  const out = shapeTextCopy(paras), props = shapeNormalizePatch(patch), start = Math.min(selection.start, selection.end), end = Math.max(selection.start, selection.end);
  const a = shapeTextAt(out, start), b = shapeTextAt(out, end); let offset = 0;
  for (let p = 0; p < out.length; p++) {
    const para = out[p], n = shapeTextLength(para), selected = p >= a.p && p <= b.p && (start === end || p < b.p || b.offset > 0);
    if (selected && props.align !== undefined) para.align = props.align;
    const runPatch = {}; for (const k of shapeRunKeys) if (props[k] !== undefined) runPatch[k] = props[k];
    if (Object.keys(runPatch).length && start < end && end > offset && start < offset + n) {
      const lo = Math.max(0, start - offset), hi = Math.min(n, end - offset), runs = shapeRunSlice(para.runs, 0, lo);
      for (const run of shapeRunSlice(para.runs, lo, hi)) shapeAppendRun(runs, { ...run, ...runPatch });
      for (const run of shapeRunSlice(para.runs, hi, Infinity)) shapeAppendRun(runs, run);
      para.runs = runs;
    } else if (selected && n === 0 && Object.keys(runPatch).length) para.runs = [{ ...(para.runs[0] ?? {}), ...runPatch, t: '' }];
    offset += n + 1;
  }
  return out;
}
export function shapeSelectionStyle(shape, paras, selection) {
  const base = { b: !!shape.bold, i: !!shape.italic, u: !!shape.underline, s: !!shape.strike, sz: shape.size ?? 11, font: shape.font ?? '맑은 고딕', color: shape.color ?? '#000000', align: shape.align ?? (shape.kind === 'textbox' ? 'left' : 'center'), valign: shape.valign ?? (shape.kind === 'textbox' ? 'top' : 'middle') };
  const start = Math.min(selection.start, selection.end), end = Math.max(selection.start, selection.end), styles = []; let offset = 0;
  const caret = shapeTextAt(paras, start);
  for (let p = 0; p < paras.length; p++) {
    const para = paras[p], n = shapeTextLength(para), ps = { ...base, ...(para.sz !== undefined ? { sz: para.sz } : {}), ...(para.align !== undefined ? { align: para.align } : {}) }; let at = 0;
    for (const r of para.runs) {
      const len = r.t.length, selected = start === end ? p === caret.p && (caret.offset > at && caret.offset <= at + len || caret.offset === 0 && at === 0) : offset + at < end && offset + at + len > start;
      if (selected) { const s = { ...ps }; for (const k of shapeRunKeys) if (r[k] !== undefined) s[k] = r[k]; styles.push(s); }
      at += len;
    }
    if (n === 0 && (start === end ? p === caret.p : offset >= start && offset < end)) { const s = { ...ps }; for (const k of shapeRunKeys) if (para.runs[0]?.[k] !== undefined) s[k] = para.runs[0][k]; styles.push(s); }
    offset += n + 1;
  }
  const result = { ...(styles[0] ?? base) };
  for (const s of styles.slice(1)) for (const key of Object.keys(result)) if (result[key] !== s[key]) result[key] = undefined;
  return { ...result, bold: result.b, italic: result.i, underline: result.u, strike: result.s, size: result.sz, fontSize: result.sz };
}

/** Coordinates are viewport CSS pixels; zoom is a ratio. Escape commits; cancel explicitly discards. */
export function createShapeTextEditor({ container, shape, onCommit = () => {}, onCancel = () => {}, onSelectionChange = () => {}, onSave = () => {}, onDraftChange = () => {}, onClipboardError = () => {}, onContextMenu = () => {}, isValid = () => true, fontStack = font => font } = {}) {
  const doc = container?.ownerDocument ?? document, win = doc.defaultView, host = container ?? doc.body, original = shapeTextCopy(shape ?? {});
  let paras = shapeTextParagraphs(original), attrs = {}, selection = { start: 0, end: 0 }, typing = {}, closed = false, composing = false, pendingCommit = false, pendingSave = false, renderDepth = 0, compositionBefore = null, selectionRevision = 0, draftRevision = 0, nativeStyle = null;
  const initial = JSON.stringify({ paras, attrs }), listeners = [], runData = new Map(), paraData = new Map(); let nextId = 0, lastGroup = '', lastTime = 0, historyIndex = 0;
  const snapshot = () => ({ paras: shapeTextCopy(paras), attrs: { ...attrs }, selection: { ...selection }, typing: { ...typing } });
  const history = [snapshot()];
  const element = doc.createElement('div'), frame = doc.createElement('div'), content = doc.createElement('div');
  element.className = 'shape-text-editor'; frame.className = 'shape-text-editor-frame'; content.className = 'shape-text-editor-content';
  content.contentEditable = 'true'; content.setAttribute('role', 'textbox'); content.setAttribute('aria-label', '도형 텍스트 편집'); content.setAttribute('aria-multiline', 'true'); content.setAttribute('spellcheck', 'false'); content.setAttribute('autocorrect', 'off'); content.setAttribute('autocapitalize', 'off'); content.tabIndex = 0;
  frame.append(content); element.append(frame); host.append(element);
  const baseStyle = () => ({ ...original, ...attrs });
  const valid = () => !closed && isValid() !== false;
  const listen = (node, event, fn, options) => { node.addEventListener(event, fn, options); listeners.push(() => node.removeEventListener(event, fn, options)); };
  function style() { const selected = shapeSelectionStyle(baseStyle(), paras, selection); return { ...selected, ...typing, ...Object.fromEntries(Object.entries(shapeTextAliases).map(([alias, key]) => [alias, typing[key] ?? selected[key]])) }; }
  function notify() { if (!closed && !composing) onSelectionChange(style()); }
  function applyStyle(node, s) {
    if (s.b !== undefined) node.style.fontWeight = s.b ? '700' : '400';
    if (s.i !== undefined) node.style.fontStyle = s.i ? 'italic' : 'normal';
    if (s.u !== undefined || s.s !== undefined) node.style.textDecoration = [s.u ? 'underline' : '', s.s ? 'line-through' : ''].filter(Boolean).join(' ') || 'none';
    if (s.sz !== undefined) node.style.fontSize = s.sz + 'pt';
    if (s.font) node.style.fontFamily = fontStack(s.font);
    if (s.color) node.style.color = s.color;
  }
  function render() {
    renderDepth++; runData.clear(); paraData.clear(); content.replaceChildren();
    const base = baseStyle(); content.style.fontFamily = fontStack(base.font ?? '맑은 고딕'); content.style.fontSize = (base.size ?? 11) + 'pt'; content.style.color = base.color ?? '#000000'; content.style.textAlign = base.align ?? (base.kind === 'textbox' ? 'left' : 'center'); content.style.whiteSpace = base.nowrap ? 'pre' : 'pre-wrap';
    content.style.fontWeight = '400'; content.style.fontStyle = 'normal'; content.style.textDecoration = 'none';
    frame.style.justifyContent = { top: 'flex-start', middle: 'center', bottom: 'flex-end' }[base.valign ?? (base.kind === 'textbox' ? 'top' : 'middle')];
    frame.style.padding = (Array.isArray(base.pad) ? base.pad : [4.8, 9.6, 4.8, 9.6]).map(n => Math.max(0, Number(n) || 0) + 'px').join(' ');
    for (const p of paras) {
      const block = doc.createElement('div'), pid = String(++nextId); block.dataset.shapeParagraph = pid; paraData.set(pid, shapeParaAttrs(p)); if (p.align) block.style.textAlign = p.align; if (p.sz) block.style.fontSize = p.sz + 'pt';
      const baseRun = { b: !!base.bold, i: !!base.italic, u: !!base.underline, s: !!base.strike, sz: p.sz ?? base.size ?? 11, font: base.font ?? '맑은 고딕', color: base.color ?? '#000000' };
      for (const r of p.runs) { const span = doc.createElement('span'), rid = String(++nextId); span.dataset.shapeRun = rid; runData.set(rid, shapeRunAttrs(r)); span.textContent = r.t; applyStyle(span, { ...baseRun, ...r }); block.append(span); }
      if (!shapeTextLength(p)) block.append(doc.createElement('br')); content.append(block);
    }
    renderDepth--; restoreSelection();
  }
  // Native composition can replace or split spans. Read text nodes, never injected HTML.
  function readDOM() {
    const result = [], points = new WeakMap(), groups = []; let inline = [];
    const isBlock = n => n.nodeType === 1 && /^(DIV|P|LI|H[1-6])$/.test(n.tagName);
    const flush = () => { if (inline.length) { groups.push({ nodes: inline }); inline = []; } };
    for (const node of content.childNodes) { if (isBlock(node)) { flush(); groups.push({ nodes: [...node.childNodes], block: node }); } else inline.push(node); } flush();
    if (!groups.length) groups.push({ nodes: [] }); let position = 0;
    for (const group of groups) {
      if (result.length) position++;
      const p = { ...(paraData.get(group.block?.dataset.shapeParagraph) ?? {}), runs: [] }, empty = group.nodes.every(n => !n.textContent) && group.nodes.filter(n => n.nodeName === 'BR').length <= 1;
      function walk(node, inherited = {}) {
        const start = position;
        if (node.nodeType === 3) { const t = node.nodeValue ?? ''; if (t) shapeAppendRun(p.runs, { ...inherited, t }); position += t.length; points.set(node, { start, end: position }); return; }
        if (node.nodeType !== 1) { points.set(node, { start, end: start, boundaries: [start] }); return; }
        const declared = { ...inherited, ...(runData.get(node.dataset.shapeRun) ?? {}) };
        if (/^(B|STRONG)$/.test(node.tagName)) declared.b = true; if (/^(I|EM)$/.test(node.tagName)) declared.i = true; if (node.tagName === 'U') declared.u = true; if (/^(S|STRIKE)$/.test(node.tagName)) declared.s = true;
        if (node.tagName === 'BR') { if (!empty) { shapeAppendRun(p.runs, { ...declared, t: '\n' }); position++; } points.set(node, { start, end: position, boundaries: [start] }); return; }
        const boundaries = [position]; for (const child of node.childNodes) { walk(child, declared); boundaries.push(position); } points.set(node, { start, end: position, boundaries });
      }
      const start = position, boundaries = [position]; for (const node of group.nodes) { walk(node); boundaries.push(position); }
      if (group.block) points.set(group.block, { start, end: position, boundaries });
      if (!p.runs.length && group.block) { const old = paras[result.length]; if (old && !shapeTextLength(old)) p.runs = shapeTextCopy(old.runs); }
      result.push(p);
    }
    const boundaries = [...content.childNodes].map(n => points.get(n)?.start ?? 0); boundaries.push(position); points.set(content, { start: 0, end: position, boundaries });
    return { paras: result, points };
  }
  function rememberSelection() {
    if (closed || renderDepth) return { ...selection };
    const sel = doc.getSelection(); if (!sel?.anchorNode || !content.contains(sel.anchorNode) || !content.contains(sel.focusNode)) return { ...selection };
    const { points } = readDOM(), at = (node, offset) => { const p = points.get(node); return !p ? 0 : node.nodeType === 3 ? Math.min(p.end, p.start + offset) : p.boundaries?.[offset] ?? p.end; };
    const anchor = at(sel.anchorNode, sel.anchorOffset), focus = at(sel.focusNode, sel.focusOffset), next = { start: Math.min(anchor, focus), end: Math.max(anchor, focus), backward: focus < anchor };
    if (next.start !== selection.start || next.end !== selection.end || next.backward !== selection.backward) { selectionRevision++; typing = {}; lastGroup = ''; }
    selection = next; return { ...selection };
  }
  function pointAt(position) {
    const p = shapeTextAt(paras, position), block = content.children[p.p] ?? content; let left = p.offset;
    const walk = doc.createTreeWalker(block, 4); let node, last;
    while ((node = walk.nextNode())) { last = node; if (left <= node.nodeValue.length) return [node, left]; left -= node.nodeValue.length; }
    return last ? [last, last.nodeValue.length] : [block, 0];
  }
  function restoreSelection() {
    if (closed || composing) return;
    const sel = doc.getSelection(), a = pointAt(selection.start), b = pointAt(selection.end); if (!sel) return;
    const range = doc.createRange(); range.setStart(...a); range.setEnd(...b); sel.removeAllRanges(); sel.addRange(range);
    if (selection.backward && sel.setBaseAndExtent) sel.setBaseAndExtent(...b, ...a);
  }
  function focus(options = {}) {
    if (!valid()) return false; content.focus({ preventScroll: true });
    if (options.selectAll) { selectionRevision++; selection = { start: 0, end: shapeTextString(paras).length }; }
    if (options.point) {
      const { x, y } = options.point, caret = doc.caretPositionFromPoint?.(x, y), range = !caret ? doc.caretRangeFromPoint?.(x, y) : null, node = caret?.offsetNode ?? range?.startContainer, offset = caret?.offset ?? range?.startOffset;
      if (node && content.contains(node)) { const r = doc.createRange(); r.setStart(node, offset); r.collapse(true); doc.getSelection().removeAllRanges(); doc.getSelection().addRange(r); rememberSelection(); }
    }
    restoreSelection(); notify(); return true;
  }
  function record(group = '') {
    const current = snapshot(), changed = JSON.stringify({ paras, attrs }) !== JSON.stringify({ paras: history[historyIndex].paras, attrs: history[historyIndex].attrs });
    if (!changed) return;
    history.splice(historyIndex + 1); const now = Date.now();
    if (group && group === lastGroup && now - lastTime < 1000 && historyIndex > 0) history[historyIndex] = current;
    else { history.push(current); historyIndex++; }
    if (history.length > 200) { history.shift(); historyIndex--; }
    lastGroup = group; lastTime = now;
  }
  function saveHistorySelection() { history[historyIndex].selection = { ...selection }; history[historyIndex].typing = { ...typing }; }
  function draftChanged() { const patch = getPatch(); onDraftChange({ dirty: Object.keys(patch).length > 0, patch }); }
  function change(next, nextSelection, group = '', priorSelection = selection) { draftRevision++; history[historyIndex].selection = { ...priorSelection }; history[historyIndex].typing = { ...typing }; paras = next; selection = nextSelection; record(group); render(); focus(); draftChanged(); }
  function insertText(text) {
    if (!valid() || composing) return false; rememberSelection(); const value = String(text).replace(/\r\n?/g, '\n'), start = selection.start, next = replaceShapeText(paras, selection, value, style());
    change(next, { start: start + value.length, end: start + value.length }, value.includes('\n') ? '' : 'typing'); return true;
  }
  function format(patchOrFn) {
    if (!valid() || composing) return false; rememberSelection(); const patch = shapeNormalizePatch(typeof patchOrFn === 'function' ? patchOrFn(style()) : patchOrFn); lastGroup = '';
    saveHistorySelection(); draftRevision++; if (patch.valign !== undefined) attrs.valign = patch.valign;
    paras = formatShapeText(paras, selection, patch);
    if (selection.start === selection.end) for (const key of shapeRunKeys) if (patch[key] !== undefined) typing[key] = patch[key];
    record(); render(); focus(); draftChanged(); return true;
  }
  function moveHistory(delta) {
    if (!valid() || composing) return false; const i = historyIndex + delta; if (i < 0 || i >= history.length) return false;
    draftRevision++; historyIndex = i; const state = shapeTextCopy(history[i]); paras = state.paras; attrs = state.attrs; selection = state.selection; typing = state.typing; lastGroup = ''; render(); focus(); draftChanged(); return true;
  }
  function getPatch() {
    if (!closed && !composing) rememberSelection();
    return JSON.stringify({ paras, attrs }) === initial ? {} : { text: shapeTextString(paras), paras: shapeTextCopy(paras), ...attrs };
  }
  function destroy() { if (closed) return; closed = true; for (const remove of listeners) remove(); element.remove(); }
  function cancel() { if (closed) return false; destroy(); onCancel(); return true; }
  function commit() {
    if (closed) return false; if (!valid()) { cancel(); return false; }
    if (composing) { pendingCommit = true; return false; }
    const patch = getPatch(); destroy(); onCommit(patch); return true;
  }
  function reposition(rect) {
    if (closed || !rect) return;
    const z = Math.max(.01, Number(rect.zoom) || 1), width = Math.max(1, Number(rect.width) || 1), height = Math.max(1, Number(rect.height) || 1);
    Object.assign(element.style, { left: (Number(rect.x) || 0) + 'px', top: (Number(rect.y) || 0) + 'px', width: width + 'px', height: height + 'px', transform: `rotate(${Number(rect.rotation) || 0}deg)` });
    const textRotation = [90, 270].includes(Number(original.textRot)) ? Number(original.textRot) : 0, fw = (textRotation ? height : width) / z, fh = (textRotation ? width : height) / z;
    Object.assign(frame.style, { width: fw + 'px', height: fh + 'px', left: (width - fw) / 2 + 'px', top: (height - fh) / 2 + 'px', transformOrigin: 'center', transform: `rotate(${textRotation}deg) scale(${z})` });
  }
  function clipboardToken() { rememberSelection(); return { draft: draftRevision, selection: selectionRevision }; }
  function clipboardCurrent(token) { rememberSelection(); return valid() && !composing && token.draft === draftRevision && token.selection === selectionRevision; }
  async function copy({ cut = false } = {}) {
    if (!valid() || composing) return false;
    const token = clipboardToken(), text = shapeTextString(paras).slice(selection.start, selection.end); if (!text) return false;
    try {
      if (!win.navigator.clipboard?.writeText) throw new Error('clipboard unavailable');
      await win.navigator.clipboard.writeText(text);
      if (cut) { if (!clipboardCurrent(token)) return false; insertText(''); }
      return true;
    } catch (error) { if (!closed) onClipboardError('클립보드에 복사할 수 없습니다. 텍스트를 선택하고 Ctrl+C 또는 Command+C를 눌러 주세요.', error); return false; }
  }
  async function paste() {
    if (!valid() || composing) return false; const token = clipboardToken();
    try {
      if (!win.navigator.clipboard?.readText) throw new Error('clipboard unavailable');
      const text = await win.navigator.clipboard.readText(); if (!clipboardCurrent(token)) return false; return insertText(text);
    } catch (error) { if (!closed) onClipboardError('클립보드를 읽을 수 없습니다. Ctrl+V 또는 Command+V로 붙여넣어 주세요.', error); return false; }
  }
  function contextMenu(point) { if (!valid() || composing) return; rememberSelection(); const r = content.getBoundingClientRect(); onContextMenu(point ?? { x: Math.max(0, r.left + 12), y: Math.max(0, r.top + 12) }); }
  function deleteText(direction, word = false) {
    rememberSelection(); const previous = { ...selection };
    if (selection.start === selection.end) {
      const text = shapeTextString(paras), pos = selection.start; let next = pos;
      if (word) next = direction < 0 ? text.slice(0, pos).replace(/\s*\S+\s*$/, '').length : pos + (text.slice(pos).match(/^\s*\S+\s*/)?.[0].length ?? 0);
      else if (typeof Intl.Segmenter === 'function') { const parts = new Intl.Segmenter(undefined, { granularity: 'grapheme' }).segment(text); for (const part of parts) { if (direction < 0 && part.index < pos) next = part.index; if (direction > 0 && part.index + part.segment.length > pos) { next = part.index + part.segment.length; break; } } }
      else next = direction < 0 ? pos - (Array.from(text.slice(0, pos)).pop()?.length ?? 0) : pos + (Array.from(text.slice(pos))[0]?.length ?? 0);
      selection = { start: Math.min(pos, next), end: Math.max(pos, next) };
    }
    const start = selection.start, next = replaceShapeText(paras, selection, ''); typing = {}; change(next, { start, end: start }, 'delete', previous);
  }
  listen(content, 'beforeinput', e => {
    e.stopPropagation(); if (!valid()) { e.preventDefault(); cancel(); return; } if (composing || e.isComposing) return; rememberSelection(); nativeStyle = style();
    if (e.inputType === 'insertText' && e.data != null) { e.preventDefault(); insertText(e.data); }
    else if (['insertParagraph', 'insertLineBreak'].includes(e.inputType)) { e.preventDefault(); insertText('\n'); }
    else if (/^delete(Content|Word)(Backward|Forward)$/.test(e.inputType)) { e.preventDefault(); deleteText(e.inputType.endsWith('Backward') ? -1 : 1, e.inputType.includes('Word')); }
    else if (e.inputType === 'historyUndo' || e.inputType === 'historyRedo') { e.preventDefault(); moveHistory(e.inputType === 'historyUndo' ? -1 : 1); }
  });
  listen(content, 'input', e => {
    e.stopPropagation(); if (composing || e.isComposing || closed) return; rememberSelection();
    const before = shapeTextString(paras), after = shapeTextString(readDOM().paras);
    // A final non-composing input can follow compositionend. Identical text must keep the rich model just restored there.
    if (before !== after) {
      let start = 0, end = 0;
      while (start < before.length && start < after.length && before[start] === after[start]) start++;
      if (start > 0 && /[\uDC00-\uDFFF]/.test(before[start] ?? after[start] ?? '') && /[\uD800-\uDBFF]/.test(before[start - 1] ?? after[start - 1] ?? '')) start--;
      while (end < before.length - start && end < after.length - start && before[before.length - end - 1] === after[after.length - end - 1]) end++;
      if (end > 0 && /[\uDC00-\uDFFF]/.test(before[before.length - end] ?? '') && /[\uD800-\uDBFF]/.test(before[before.length - end - 1] ?? '')) end--;
      paras = replaceShapeText(paras, { start, end: before.length - end }, after.slice(start, after.length - end), nativeStyle ?? style()); draftRevision++;
    }
    nativeStyle = null; record('typing'); render(); notify(); draftChanged();
  });
  listen(content, 'compositionstart', e => { e.stopPropagation(); rememberSelection(); saveHistorySelection(); draftRevision++; compositionBefore = { paras: shapeTextCopy(paras), style: style(), text: shapeTextString(paras) }; composing = true; lastGroup = ''; });
  listen(content, 'compositionend', e => { e.stopPropagation(); composing = false; rememberSelection(); const native = readDOM().paras;
    if (compositionBefore) {
      const before = compositionBefore.text, after = shapeTextString(native); let start = 0, end = 0;
      while (start < before.length && start < after.length && before[start] === after[start]) start++;
      while (end < before.length - start && end < after.length - start && before[before.length - end - 1] === after[after.length - end - 1]) end++;
      paras = replaceShapeText(compositionBefore.paras, { start, end: before.length - end }, after.slice(start, after.length - end), compositionBefore.style); compositionBefore = null;
    } else paras = native;
    record(); draftChanged(); win.requestAnimationFrame(() => { if (closed || composing) return; render(); notify(); if (pendingCommit) { const save = pendingSave; pendingCommit = false; pendingSave = false; if (commit() && save) onSave(save); } }); });
  listen(content, 'keydown', e => {
    e.stopPropagation();
    const mod = (e.ctrlKey || e.metaKey) && !e.altKey, code = shortcutCode(e), key = /^Key[A-Z]$/.test(code) ? code.slice(3).toLowerCase() : (e.key || '').toLowerCase();
    // Save must also suppress the browser command during composition, when shortcutCode deliberately returns empty.
    const saveCode = /^Key[A-Z]$/.test(e.code || '') ? e.code : shortcutCode({ code: e.code, key: e.key, keyCode: e.keyCode === 229 ? 0 : e.keyCode, which: e.which });
    if (mod && (saveCode === 'KeyS' || !saveCode && key === 's')) { e.preventDefault(); const options = { saveAs: !!e.shiftKey }; if (composing || e.isComposing) { pendingCommit = true; pendingSave = options; } else if (commit()) onSave(options); return; }
    if (composing || e.isComposing || e.keyCode === 229) return;
    if (e.key === 'ContextMenu' || e.shiftKey && e.key === 'F10') { e.preventDefault(); contextMenu(); }
    else if (mod && ['b', 'i', 'u'].includes(key)) { e.preventDefault(); format(s => ({ [key]: !s[key] })); }
    else if (mod && (key === 'z' || key === 'y')) { e.preventDefault(); moveHistory(key === 'y' || e.shiftKey ? 1 : -1); }
    else if (e.key === 'Escape' || e.key === 'Tab' || mod && e.key === 'Enter') { e.preventDefault(); commit(); }
    else if (e.key === 'Enter') { e.preventDefault(); insertText('\n'); }
  });
  listen(content, 'paste', e => { e.preventDefault(); e.stopPropagation(); if (e.clipboardData) insertText(e.clipboardData.getData('text/plain')); });
  listen(content, 'copy', e => { e.stopPropagation(); if (closed || !e.clipboardData) return; rememberSelection(); if (selection.start === selection.end) return; e.preventDefault(); e.clipboardData.setData('text/plain', shapeTextString(paras).slice(selection.start, selection.end)); });
  listen(content, 'cut', e => { e.stopPropagation(); if (!valid() || composing) return; rememberSelection(); if (selection.start === selection.end || !e.clipboardData) return; e.preventDefault(); e.clipboardData.setData('text/plain', shapeTextString(paras).slice(selection.start, selection.end)); insertText(''); });
  listen(content, 'drop', e => { e.preventDefault(); e.stopPropagation(); if (e.dataTransfer) { focus({ point: { x: e.clientX, y: e.clientY } }); insertText(e.dataTransfer.getData('text/plain')); } });
  for (const event of ['pointerdown', 'mousedown']) listen(element, event, e => { e.stopPropagation(); if (e.button === 0 && (e.target === element || e.target === frame)) { e.preventDefault(); focus({ point: { x: e.clientX, y: e.clientY } }); } });
  for (const event of ['pointerup', 'mouseup', 'click', 'dblclick', 'keyup', 'dragstart']) listen(element, event, e => e.stopPropagation());
  listen(element, 'contextmenu', e => { e.preventDefault(); e.stopPropagation(); contextMenu({ x: e.clientX, y: e.clientY }); });
  listen(doc, 'selectionchange', () => { if (!closed && !renderDepth && !composing) { rememberSelection(); notify(); } });
  render(); reposition({ x: original.x ?? 0, y: original.y ?? 0, width: original.w ?? 160, height: original.h ?? 80 });
  return { element, focus, format, style, commit, cancel, contains: target => !!target && element.contains(target), reposition, undo: () => moveHistory(-1), redo: () => moveHistory(1), getPatch, insertText, copy, cut: () => copy({ cut: true }), paste, rememberSelection, hasUndo: () => historyIndex > 0, hasRedo: () => historyIndex + 1 < history.length, destroy };
}
