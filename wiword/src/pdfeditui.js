// PDF 편집 창 — 쪽 위의 글자를 그 자리에서 고치기 (이음오피스의 "PDF 쪽 위의 글자 고치기" 를 참고해 직접 구현)
// 쪽은 DOM 없는 pdftools.editablePage 로 읽어 SVG 로 그리고, 글 상자는 두 번 눌러 고침. 저장하면 applyTextEdits 로 새 PDF.
import { register } from './state.js';
import { openPdf, editablePage, applyTextEdits, canEmbedFont, pdfInfo } from './pdftools.js';
import { pickFile, askPassword, saveBytes, fmtSize } from './fileio.js';
import { fontFileFor, askLocalFontsOnce } from './fontlib.js';
import { el, openDialog, toast, progressOverlay, alertDialog } from './ui.js';

const PX = 96 / 72;
const esc = (s) => String(s ?? '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
/** 그림 RGBA → data URL (편집 화면용) */
function canvasEncode(w, h, rgba, hasAlpha) {
  try {
    const c = document.createElement('canvas'); c.width = w; c.height = h;
    c.getContext('2d').putImageData(new ImageData(rgba, w, h), 0, 0);
    return hasAlpha ? c.toDataURL('image/png') : c.toDataURL('image/jpeg', 0.85);
  } catch { return null; }
}
/** 쪽 → SVG (고치는 상자의 원래 글은 빼고) */
function pageSvg(v, hide) {
  const parts = [`<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 ${v.w} ${v.h}" width="100%" height="100%" class="pe-svg"><rect width="${v.w}" height="${v.h}" fill="#fff"/>`];
  for (const p of v.paths) parts.push(`<path d="${p.d}" fill="${p.fill ?? 'none'}"${p.fill && p.alpha < 1 ? ` fill-opacity="${p.alpha}"` : ''}${p.evenOdd ? ' fill-rule="evenodd"' : ''}${p.stroke ? ` stroke="${p.stroke}" stroke-width="${Math.max(0.5, p.lw * PX)}"` : ''}/>`);
  for (const im of v.images) {
    const t = `${im.flipH || im.flipV ? ` transform="translate(${im.flipH ? 2 * im.x + im.w : 0} ${im.flipV ? 2 * im.y + im.h : 0}) scale(${im.flipH ? -1 : 1} ${im.flipV ? -1 : 1})"` : ''}`;
    parts.push(`<image href="${esc(im.url)}" x="${im.x}" y="${im.y}" width="${im.w}" height="${im.h}" preserveAspectRatio="none"${t}/>`);
  }
  for (const t of v.texts) {
    if (t.block >= 0 && hide.has(t.block)) continue;
    const rot = Math.abs(t.angle) > 0.5 ? ` transform="rotate(${t.angle} ${t.x0} ${t.y0})"` : '';
    parts.push(`<text x="${t.x0}" y="${t.y0}" font-size="${t.size}" fill="${t.color}"${t.alpha < 1 ? ` fill-opacity="${t.alpha}"` : ''} textLength="${Math.max(1, Math.hypot(t.x1 - t.x0, (t.y1 ?? t.y0) - t.y0))}" lengthAdjust="spacingAndGlyphs" font-family="${esc(t.font)}, 'Malgun Gothic', 'Noto Sans KR', sans-serif"${t.b ? ' font-weight="700"' : ''}${t.i ? ' font-style="italic"' : ''}${rot} xml:space="preserve">${esc(t.text)}</text>`);
  }
  parts.push('</svg>');
  return parts.join('');
}

/** PDF 편집 창. file 을 주면 바로 그 파일 */
export async function pdfEditDialog(file = null) {
  file ??= await pickFile('.pdf,application/pdf');
  if (!file) return;
  const bytes = new Uint8Array(await file.arrayBuffer());
  let password = ''; let wrong = false; let src;
  for (;;) {
    try { src = openPdf(bytes, { password }); break; } catch (e) {
      if (e.code !== 'PASSWORD') { alertDialog('PDF 편집', e.message); return; }
      const pw = await askPassword(file.name, wrong || !!e.wrong);
      if (pw == null) return;
      password = pw; wrong = true;
    }
  }
  const views = new Map();
  const view = (i) => { if (!views.has(i)) views.set(i, editablePage(src, i, { encodeImage: canvasEncode })); return views.get(i); };
  const edits = new Map(); // key 'p:b' 또는 'p:n<k>' → { page, block?, x?, y?, text, size, color, bold }
  const history = [];
  let page = 0; let sel = null; let adding = false; let newSeq = 0;
  const N = src.pages.length;
  const host = el('div', { class: 'pe-page' });
  const scroller = el('div', { class: 'pe-scroll' }, host);
  const pageLabel = el('span', { class: 'pe-pageno' });
  const sizeIn = el('input', { type: 'number', min: '4', max: '200', step: '0.5', title: '글자 크기 (pt)', style: { width: '64px' } });
  const colorIn = el('input', { type: 'color', title: '글자 색' });
  const boldIn = el('input', { type: 'checkbox', title: '굵게' });
  const status = el('span', { class: 'muted pe-status' });
  let scale = 1;
  const record = (key, val) => { history.push([key, edits.has(key) ? { ...edits.get(key) } : null]); if (val) edits.set(key, val); else edits.delete(key); render(); };
  const current = (key) => {
    const e = edits.get(key);
    if (e) return e;
    const [pg, bk] = key.split(':');
    const b = view(Number(pg)).blocks[Number(bk)];
    return { page: Number(pg), block: Number(bk), text: b.text, size: b.size, color: b.color, bold: b.b };
  };
  const syncTools = () => {
    const e = sel ? current(sel) : null;
    for (const i of [sizeIn, colorIn, boldIn]) i.disabled = !e;
    if (!e) return;
    sizeIn.value = String(e.size); colorIn.value = /^#[0-9a-f]{6}$/i.test(e.color) ? e.color : '#000000'; boldIn.checked = !!e.bold;
  };
  const changeSel = (patch) => { if (!sel) return; record(sel, { ...current(sel), ...patch }); };
  sizeIn.addEventListener('change', () => changeSel({ size: Math.max(4, Number(sizeIn.value) || 11) }));
  colorIn.addEventListener('change', () => changeSel({ color: colorIn.value }));
  boldIn.addEventListener('change', () => changeSel({ bold: boldIn.checked }));
  const startEdit = (key, box, after = null) => {
    const e = current(key);
    const ta = el('textarea', { class: 'pe-ta', spellcheck: 'false' });
    ta.value = e.text;
    Object.assign(ta.style, { left: `${box.x * scale}px`, top: `${box.y * scale}px`, width: `${Math.max(120, box.w * scale + 24)}px`, height: `${Math.max(28, box.h * scale + 12)}px`, fontSize: `${e.size * PX * scale}px`, color: e.color, fontWeight: e.bold ? '700' : '400' });
    host.append(ta);
    ta.focus(); ta.select();
    let done = false;
    const finish = (ok) => {
      if (done) return; done = true; ta.remove();
      if (ok && ta.value !== e.text) record(key, { ...e, text: ta.value });
      else render();
      after?.();
    };
    ta.addEventListener('blur', () => finish(true));
    ta.addEventListener('keydown', (ev) => { ev.stopPropagation(); if (ev.key === 'Escape') finish(false); if (ev.key === 'Enter' && (ev.ctrlKey || ev.metaKey)) finish(true); });
  };
  function render() {
    const v = view(page);
    const width = Math.min(scroller.clientWidth - 24 || 820, 1000);
    scale = width / v.w;
    host.style.width = `${v.w * scale}px`; host.style.height = `${v.h * scale}px`;
    const hide = new Set([...edits.values()].filter((e) => e.page === page && e.block != null).map((e) => e.block));
    host.innerHTML = pageSvg(v, hide);
    const add = (key, box, e, isEdited) => {
      const d = el('div', { class: `pe-block${isEdited ? ' edited' : ''}${sel === key ? ' sel' : ''}`, title: '두 번 눌러 고치기' });
      Object.assign(d.style, { left: `${box.x * scale}px`, top: `${box.y * scale}px`, width: `${box.w * scale}px`, height: `${box.h * scale}px` });
      if (isEdited && e.text) {
        const t = el('div', { class: 'pe-newtext' }, e.text);
        const b = e.block != null ? v.blocks[e.block] : null;
        Object.assign(t.style, { fontSize: `${e.size * PX * scale}px`, color: e.color, fontWeight: e.bold ? '700' : '400', lineHeight: `${(b && b.lines.length > 1 ? b.gap : e.size * PX * 1.25) * scale}px`, textAlign: b?.align === 'ctr' ? 'center' : b?.align === 'r' ? 'right' : 'left', fontFamily: `${b?.font ? `'${b.font}', ` : ''}'Malgun Gothic', 'Noto Sans KR', sans-serif` });
        d.append(t);
      }
      d.addEventListener('mousedown', (ev) => { ev.stopPropagation(); sel = key; syncTools(); for (const x of host.querySelectorAll('.pe-block.sel')) x.classList.remove('sel'); d.classList.add('sel'); });
      d.addEventListener('dblclick', (ev) => { ev.stopPropagation(); startEdit(key, box); });
      host.append(d);
    };
    v.blocks.forEach((b, bi) => {
      const key = `${page}:${bi}`;
      const e = edits.get(key);
      add(key, b, e ?? { text: b.text }, !!e);
    });
    for (const [key, e] of edits) {
      if (e.page !== page || e.block != null) continue;
      const lines = String(e.text).split('\n');
      add(key, { x: e.x, y: e.y, w: Math.max(...lines.map((l) => [...l].reduce((a, ch) => a + (ch.charCodeAt(0) > 0x2e80 ? 1 : 0.6), 0))) * e.size * PX + 8, h: lines.length * e.size * PX * 1.25 + 4 }, e, true);
    }
    pageLabel.textContent = `${page + 1} / ${N} 쪽`;
    status.textContent = edits.size ? `고친 곳 ${edits.size}개` : '글 상자를 두 번 눌러 고치세요';
    syncTools();
  }
  host.addEventListener('mousedown', (ev) => {
    if (!adding) { sel = null; for (const x of host.querySelectorAll('.pe-block.sel')) x.classList.remove('sel'); syncTools(); return; }
    ev.preventDefault(); // 누른 자리로 포커스가 옮겨 가 새 입력 칸이 바로 닫히지 않게
    const r = host.getBoundingClientRect();
    const x = (ev.clientX - r.left) / scale; const y = (ev.clientY - r.top) / scale;
    adding = false; host.classList.remove('adding');
    const key = `${page}:n${++newSeq}`;
    const e = { page, x, y, text: '', size: 11, color: '#000000', bold: false };
    edits.set(key, e); sel = key;
    render();
    startEdit(key, { x, y, w: 160, h: e.size * PX * 1.5 }, () => { if (!String(edits.get(key)?.text ?? '').trim()) { edits.delete(key); sel = null; render(); } });
  });
  const go = (d) => { const n = Math.max(0, Math.min(N - 1, page + d)); if (n !== page) { page = n; sel = null; render(); scroller.scrollTop = 0; } };
  const del = () => {
    if (!sel) { toast('지울 글 상자를 먼저 누르세요'); return; }
    if (sel.includes(':n')) { record(sel, null); sel = null; return; }
    record(sel, { ...current(sel), text: '' });
  };
  const undo = () => { const h = history.pop(); if (!h) return; if (h[1]) edits.set(h[0], h[1]); else edits.delete(h[0]); render(); };
  const save = async () => {
    if (!edits.size) { toast('고친 곳이 없습니다'); return; }
    const list = [...edits.values()];
    let font = null;
    const needFont = list.some((e) => /[^\x20-\x7e\n]/.test(e.text));
    if (needFont) {
      await askLocalFontsOnce().catch(() => false);
      const fams = [...new Set(list.map((e) => (e.block != null ? view(e.page).blocks[e.block]?.font : null)).filter(Boolean)), '맑은 고딕', 'Malgun Gothic', 'Apple SD Gothic Neo', 'Noto Sans KR', 'NanumGothic', '나눔고딕'];
      for (const fam of fams) { const r = await fontFileFor(fam).catch(() => null); if (r?.bytes && canEmbedFont(r.bytes, r.ps)) { font = r; break; } }
    }
    const bar = progressOverlay('PDF 만드는 중…');
    try {
      const out = await applyTextEdits(bytes, list, { password, font: font?.bytes, fontName: font?.ps });
      bar.close();
      const name = `${file.name.replace(/\.pdf$/i, '')}_수정.pdf`;
      const ok = await saveBytes(out, name);
      if (ok !== false) toast(`저장했습니다: ${name} (${fmtSize(out.length)})${needFont && !font ? ' — 한글 글꼴을 찾지 못해 보는 프로그램의 글꼴로 표시됩니다' : ''}`);
    } catch (e) { alertDialog('PDF 편집', e.message); } finally { bar.close(); }
  };
  const tb = el('div', { class: 'pe-toolbar' },
    el('button', { class: 'btn', onclick: () => go(-1), title: '이전 쪽 (PageUp)' }, '◀'), pageLabel, el('button', { class: 'btn', onclick: () => go(1), title: '다음 쪽 (PageDown)' }, '▶'),
    el('span', { class: 'pe-sep' }),
    el('button', { class: 'btn', onclick: () => { adding = true; host.classList.add('adding'); toast('글을 넣을 자리를 누르세요'); } }, '＋ 글 상자'),
    el('button', { class: 'btn', onclick: del }, '지우기'),
    el('button', { class: 'btn', onclick: undo, title: '실행 취소 (Ctrl+Z)' }, '되돌리기'),
    el('span', { class: 'pe-sep' }),
    el('label', { class: 'pe-tool' }, '크기 ', sizeIn), el('label', { class: 'pe-tool' }, '색 ', colorIn), el('label', { class: 'pe-tool' }, boldIn, ' 굵게'),
    el('span', { class: 'pe-grow' }), status,
    el('button', { class: 'btn primary', onclick: save }, 'PDF 로 저장…'));
  const info = pdfInfo(bytes, { password });
  const dlg = openDialog({
    title: `PDF 편집 — ${file.name} (${info.pages.length}쪽)`, width: Math.min(1180, window.innerWidth - 24),
    body: el('div', { class: 'pe-wrap' }, tb, scroller,
      el('div', { class: 'muted pe-help' }, '글 상자를 두 번 누르면 고칩니다 (Ctrl+Enter 마침, Esc 취소). 저장하면 원래 글을 지우고 같은 자리에 새 글을 그린 새 PDF 를 만듭니다. 스캔한 그림 속 글자는 고칠 수 없습니다.')),
    onClose: () => views.clear(),
  });
  dlg.root.addEventListener('keydown', (ev) => {
    if (ev.target.tagName === 'TEXTAREA' || ev.target.tagName === 'INPUT') return;
    if (ev.key === 'PageDown') { ev.preventDefault(); go(1); }
    if (ev.key === 'PageUp') { ev.preventDefault(); go(-1); }
    if (ev.key === 'Delete') { ev.preventDefault(); del(); }
    if ((ev.ctrlKey || ev.metaKey) && ev.key.toLowerCase() === 'z') { ev.preventDefault(); undo(); }
  });
  requestAnimationFrame(render);
}

register({ pdfEdit: () => pdfEditDialog() }, { pdfEdit: { label: 'PDF 편집', desc: 'PDF 쪽 위의 글자를 그 자리에서 고치기' } });
