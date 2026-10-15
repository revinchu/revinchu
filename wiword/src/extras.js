// 나머지 명령: 스타일 수정/새 스타일 · Word 옵션 · 정보 · 새로운 기능 · 바로 가기 키 · 아이콘 · 차트 · 공유
import { S, run, register, change, curSel, emit } from './state.js';
import { blockAt, effRun, effPara, uid } from './model.js';
import { el, openDialog, alertDialog, toast } from './ui.js';
import { SHORTCUTS } from './panels.js';
import { FONTS } from './ribbon.js';

export const APP_VERSION = '1.0.0';

// ───────────── 스타일 수정 · 새 스타일 ─────────────
function styleDialog(id = null) {
  const st = id ? S.doc.styles[id] : null;
  const cur = blockAt(S.doc, curSel().f.p);
  const base = st ?? { name: '스타일1', type: 'para', basedOn: cur?.style ?? 'Normal', next: 'Normal', pPr: {}, rPr: {} };
  const fake = { type: 'p', id: 'x', style: id ?? base.basedOn, pPr: {}, runs: [{ text: 'x' }] };
  const er = effRun(S.doc, fake, fake.runs[0]);
  const ep = effPara(S.doc, fake);
  const name = el('input', { type: 'text', value: base.name ?? '' });
  const type = el('select', { disabled: !!st }, [['para', '단락'], ['char', '문자']].map(([v, l]) => el('option', { value: v, selected: v === base.type }, l)));
  const paraStyles = Object.values(S.doc.styles).filter((s) => s.type === 'para' && s.id !== id);
  const basedOn = el('select', {}, [el('option', { value: '' }, '(스타일 없음)'), ...paraStyles.map((s) => el('option', { value: s.id, selected: s.id === base.basedOn }, s.name))]);
  const next = el('select', {}, paraStyles.map((s) => el('option', { value: s.id, selected: s.id === (base.next ?? 'Normal') }, s.name)));
  const font = el('input', { type: 'text', value: (base.rPr?.fontEa ?? base.rPr?.font ?? '').replace(/^\+m[jn]$/, ''), list: 'sd-fonts', placeholder: er.fontEa?.startsWith('+') ? '(테마 글꼴)' : er.fontEa });
  const size = el('input', { type: 'number', min: '1', max: '400', step: '0.5', value: String(base.rPr?.sz ?? er.sz ?? 10), style: { width: '64px' } });
  const tog = (k, l) => { const b = el('button', { type: 'button', class: `rbtn small${(base.rPr?.[k] ?? er[k]) ? ' on' : ''}`, title: l }, l); b.addEventListener('click', () => b.classList.toggle('on')); b.dataset.k = k; return b; };
  const B = tog('b', '굵게'); const I = tog('i', '기울임'); const U = tog('u', '밑줄');
  const color = el('input', { type: 'color', value: base.rPr?.color ?? er.color ?? '#000000' });
  const align = el('select', {}, [['', '(상속)'], ['left', '왼쪽'], ['center', '가운데'], ['right', '오른쪽'], ['justify', '양쪽']].map(([v, l]) => el('option', { value: v, selected: v === (base.pPr?.align ?? '') }, l)));
  const pt = (px) => Math.round(((px ?? 0) * 0.75) * 10) / 10;
  const before = el('input', { type: 'number', step: '1', value: String(pt(base.pPr?.spc?.before ?? ep.spc?.before)), style: { width: '64px' } });
  const after = el('input', { type: 'number', step: '1', value: String(pt(base.pPr?.spc?.after ?? ep.spc?.after)), style: { width: '64px' } });
  const line = el('input', { type: 'number', step: '0.05', min: '0.5', value: String(base.pPr?.spc?.line ?? ep.spc?.line ?? 1), style: { width: '64px' } });
  const quick = el('input', { type: 'checkbox', checked: base.quick !== false });
  const preview = el('div', { class: 'sd-prev', style: { border: '1px solid #ccc', padding: '10px', minHeight: '60px', background: '#fff' } });
  const draw = () => { preview.style.font = `${B.classList.contains('on') ? 'bold ' : ''}${I.classList.contains('on') ? 'italic ' : ''}${size.value}pt '${font.value || '맑은 고딕'}', sans-serif`; preview.style.color = color.value; preview.style.textAlign = align.value || 'left'; preview.style.textDecoration = U.classList.contains('on') ? 'underline' : 'none'; preview.textContent = `${name.value || '스타일'} — 가나다라마바사 AaBbCc 123`; };
  for (const x of [name, font, size, color, align]) x.addEventListener('input', draw);
  for (const x of [B, I, U]) x.addEventListener('click', draw);
  const row = (l, ...c) => el('label', { class: 'dlg-row' }, el('span', {}, l), ...c);
  openDialog({
    title: st ? `스타일 수정 — ${st.name}` : '새 스타일 만들기', width: 560,
    body: el('div', { class: 'dlg-form' },
      el('datalist', { id: 'sd-fonts' }, FONTS.filter((f) => !f.startsWith('+')).map((f) => el('option', { value: f }))),
      row('이름:', name), row('스타일 형식:', type), row('스타일 기준:', basedOn), row('다음 단락의 스타일:', next),
      el('h4', {}, '서식'),
      el('div', { class: 'dlg-row' }, font, size, B, I, U, color),
      el('div', { class: 'dlg-row' }, el('span', {}, '맞춤'), align, el('span', {}, '앞(pt)'), before, el('span', {}, '뒤(pt)'), after, el('span', {}, '줄 간격(배)'), line),
      preview,
      el('label', { class: 'pane-check' }, quick, '스타일 갤러리에 추가')),
    buttons: [{ label: '확인', primary: true, action: () => {
      const nm = name.value.trim();
      if (!nm) { toast('이름을 입력하세요'); return false; }
      const sid = id ?? (`s${nm.replace(/[^A-Za-z0-9가-힣]/g, '')}` || uid('s'));
      const rPr = { ...(st?.rPr ?? {}) };
      if (font.value) { rPr.font = font.value; rPr.fontEa = font.value; }
      rPr.sz = Number(size.value) || undefined;
      for (const b of [B, I, U]) rPr[b.dataset.k] = b.classList.contains('on') ? true : undefined;
      if (color.value && color.value !== '#000000') rPr.color = color.value.toUpperCase(); else delete rPr.color;
      for (const k of Object.keys(rPr)) if (rPr[k] === undefined) delete rPr[k];
      const pPr = { ...(st?.pPr ?? {}) };
      if (align.value) pPr.align = align.value; else delete pPr.align;
      pPr.spc = { ...(pPr.spc ?? {}), before: Number(before.value) / 0.75, after: Number(after.value) / 0.75, line: Number(line.value) || 1, rule: pPr.spc?.rule === 'pct' ? 'pct' : 'auto' };
      const def = { ...(st ?? {}), id: sid, name: nm, type: type.value, basedOn: basedOn.value || undefined, next: next.value, quick: quick.checked, order: st?.order ?? 35, pPr: type.value === 'para' ? pPr : undefined, rPr };
      change((d) => { const s = d.prop('styles'); s[sid] = def; return S.sel; });
      if (!id && type.value === 'para') run('applyStyle', sid);
      return true;
    } }, { label: '취소' }],
  });
  draw();
}

// ───────────── Word 옵션 ─────────────
const OPT_KEY = 'wiword:opts';
export function loadOpts() { try { Object.assign(S.opts, JSON.parse(localStorage.getItem(OPT_KEY) ?? '{}')); } catch { /* 없음 */ } }
function saveOpts() { try { localStorage.setItem(OPT_KEY, JSON.stringify(S.opts)); } catch { /* 저장 불가 */ } }
function optionsDialog() {
  const o = S.opts;
  const chk = (k, l, def = false) => { const c = el('input', { type: 'checkbox', checked: o[k] ?? def }); c.dataset.k = k; c.dataset.def = String(def); return el('label', { class: 'pane-check' }, c, l); };
  const user = el('input', { type: 'text', value: o.userName ?? '' });
  const fmt = el('select', {}, [['docx', 'Word 문서 (.docx)'], ['hwpx', '한글 문서 (.hwpx)'], ['md', '마크다운 (.md)'], ['wiword', 'WIWORD 문서']].map(([v, l]) => el('option', { value: v, selected: v === (o.saveFormat ?? 'docx') }, l)));
  const body = el('div', { class: 'dlg-form' },
    el('h4', {}, '일반'), el('label', { class: 'dlg-row' }, el('span', {}, '사용자 이름 (메모 · 변경 내용):'), user),
    chk('miniToolbar', '선택할 때 미니 도구 모음 표시', true),
    el('h4', {}, '표시'), chk('fieldShade', '필드 음영 표시', true), chk('showBoundaries', '텍스트 경계 표시'), chk('linkClick', 'Ctrl 없이 클릭해도 하이퍼링크 열기'),
    el('h4', {}, '언어 교정'), chk('autoCorrect', '입력할 때 자동 고침 (--, (c), 인용 부호 …)', true), chk('autoList', '자동 글머리 기호 · 번호 (- , 1. 로 시작)', true), chk('mdShortcuts', '마크다운 단축 입력 (# 제목, **굵게**)', true), chk('autoLink', '인터넷 주소를 하이퍼링크로', true),
    el('h4', {}, '고급'), chk('mdPaste', '마크다운 글을 붙여넣으면 서식으로 바꾸기', true), chk('smartCut', '잘라내기 · 붙여넣기 때 공백 자동 조정', true), chk('autosave', '자동 저장 (이 브라우저에 1.5초마다)', true),
    el('h4', {}, '저장'), el('label', { class: 'dlg-row' }, el('span', {}, '기본 저장 형식:'), fmt));
  openDialog({
    title: 'Word 옵션', width: 520, body,
    buttons: [{ label: '확인', primary: true, action: () => {
      for (const c of body.querySelectorAll('input[type=checkbox][data-k]')) o[c.dataset.k] = c.checked;
      o.userName = user.value.trim();
      o.saveFormat = fmt.value;
      saveOpts();
      document.body.classList.toggle('show-bounds', !!o.showBoundaries);
      emit('change', { scope: 'view' });
      return true;
    } }, { label: '취소' }],
  });
}

// ───────────── 아이콘 (WIXEL · WIPOINT 와 같은 아이콘 모음) ─────────────
let iconLib = null;
async function loadIconLib() {
  if (iconLib) return iconLib;
  let res = await fetch('assets/iconlib.json.gz').catch(() => null);
  if (!res?.ok) res = await fetch('assets/iconlib.json').catch(() => null);
  if (!res?.ok) throw new Error(`아이콘 모음을 불러오지 못했습니다 (${res?.status ?? '연결 안 됨'}).`);
  let buf = new Uint8Array(await res.arrayBuffer());
  if (buf[0] === 0x1f && buf[1] === 0x8b) buf = new Uint8Array(await new Response(new Blob([buf]).stream().pipeThrough(new DecompressionStream('gzip'))).arrayBuffer());
  iconLib = JSON.parse(new TextDecoder().decode(buf)).map(([cat, items]) => [cat, items.map(([name, body, vb]) => ({ name, body, vb: vb ?? '0 0 96 96', cat }))]);
  return iconLib;
}
const svgUrl = (text) => { const u = new TextEncoder().encode(text); let s = ''; for (let i = 0; i < u.length; i += 0x8000) s += String.fromCharCode.apply(null, u.subarray(i, i + 0x8000)); return `data:image/svg+xml;base64,${btoa(s)}`; };
async function iconsDialog() {
  let lib;
  try { lib = await loadIconLib(); } catch (e) { alertDialog('아이콘', `${e.message} 배포 파일(assets/iconlib.json.gz)을 확인하세요.`); return; }
  const all = lib.flatMap(([, items]) => items);
  const picked = new Set();
  let cat = '';
  const search = el('input', { type: 'search', placeholder: '아이콘 검색 (예: 사람, 화살표, 돈)', class: 'ic-search' });
  const cats = el('div', { class: 'ic-cats' });
  const grid = el('div', { class: 'ic-grid' });
  let list = []; let shown = 0;
  const more = () => {
    for (const ic of list.slice(shown, shown + 200)) {
      const b = el('button', { type: 'button', class: `ic-cell${picked.has(ic) ? ' on' : ''}`, title: ic.name, html: `<svg viewBox="${ic.vb}">${ic.body}</svg>` });
      b.addEventListener('click', () => { if (picked.has(ic)) picked.delete(ic); else picked.add(ic); b.classList.toggle('on', picked.has(ic)); });
      grid.append(b);
    }
    shown += 200;
  };
  const draw = () => {
    const q = search.value.trim().toLowerCase();
    const seen = new Set();
    list = (cat ? lib.find((x) => x[0] === cat)[1] : all).filter((ic) => (!q || ic.name.toLowerCase().includes(q) || ic.cat.toLowerCase().includes(q)) && (seen.has(ic.body) ? false : (seen.add(ic.body), true)));
    grid.replaceChildren(); shown = 0; more();
    cats.querySelectorAll('button').forEach((b) => b.classList.toggle('on', (b.dataset.cat ?? '') === cat));
  };
  grid.addEventListener('scroll', () => { if (grid.scrollTop + grid.clientHeight > grid.scrollHeight - 200 && shown < list.length) more(); });
  cats.append(el('button', { type: 'button', 'data-cat': '', onclick: () => { cat = ''; draw(); } }, '전체'), ...lib.map(([c, items]) => el('button', { type: 'button', 'data-cat': c, onclick: () => { cat = c; draw(); } }, `${c} `, el('small', {}, String(items.length)))));
  let t = 0;
  search.addEventListener('input', () => { clearTimeout(t); t = setTimeout(draw, 150); });
  search.addEventListener('keydown', (e) => e.stopPropagation());
  openDialog({
    title: '아이콘 삽입', width: 760,
    body: el('div', { class: 'ic-dlg' }, search, el('div', { class: 'ic-main' }, cats, grid)),
    buttons: [{ label: '삽입', primary: true, action: () => {
      if (!picked.size) return false;
      for (const ic of picked) run('insertImage', svgUrl(`<svg xmlns="http://www.w3.org/2000/svg" viewBox="${ic.vb}">${ic.body}</svg>`), 72, 72, ic.name);
      return true;
    } }, { label: '취소' }],
  });
  draw();
}

// ───────────── 차트 (SVG 그림으로 넣기 · 데이터는 그림에 보관) ─────────────
const PALETTE = ['#4472C4', '#ED7D31', '#A5A5A5', '#FFC000', '#5B9BD5', '#70AD47'];
export function chartSvg(kind, rows, { w = 480, h = 288, title = '' } = {}) {
  const head = rows[0] ?? [];
  const cats = rows.slice(1).map((r) => r[0]);
  const series = head.slice(1).map((name, si) => ({ name, vals: rows.slice(1).map((r) => Number(String(r[si + 1]).replace(/,/g, '')) || 0) }));
  const esc = (s) => String(s ?? '').replace(/[&<>"]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' })[c]);
  const T = title ? 28 : 10; const L = 44; const R = 12; const B = 46;
  let s = `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 ${w} ${h}" font-family="Malgun Gothic, sans-serif" font-size="11"><rect width="${w}" height="${h}" fill="#fff"/>`;
  if (title) s += `<text x="${w / 2}" y="18" text-anchor="middle" font-size="14" fill="#595959">${esc(title)}</text>`;
  if (kind === 'pie') {
    const vals = series[0]?.vals ?? [];
    const sum = vals.reduce((a, b) => a + Math.max(0, b), 0) || 1;
    const cx = w / 2 - 50; const cy = (h + T - B + 20) / 2; const r = Math.min(w, h) / 2 - 40;
    let a0 = -Math.PI / 2;
    vals.forEach((v, i) => {
      const a1 = a0 + (Math.max(0, v) / sum) * Math.PI * 2;
      const large = a1 - a0 > Math.PI ? 1 : 0;
      s += `<path d="M${cx},${cy} L${cx + r * Math.cos(a0)},${cy + r * Math.sin(a0)} A${r},${r} 0 ${large} 1 ${cx + r * Math.cos(a1)},${cy + r * Math.sin(a1)} Z" fill="${PALETTE[i % 6]}" stroke="#fff"/>`;
      a0 = a1;
    });
    cats.forEach((c, i) => { s += `<rect x="${w - 130}" y="${T + 10 + i * 18}" width="10" height="10" fill="${PALETTE[i % 6]}"/><text x="${w - 115}" y="${T + 19 + i * 18}" fill="#595959">${esc(c)}</text>`; });
    return `${s}</svg>`;
  }
  const all = series.flatMap((x) => x.vals);
  const max = Math.max(0, ...all); const min = Math.min(0, ...all);
  const step = niceStep((max - min) / 5 || 1);
  const top = Math.ceil(max / step) * step || step; const bot = Math.floor(min / step) * step;
  const pw = w - L - R; const ph = h - T - B;
  const y = (v) => T + ph - ((v - bot) / (top - bot)) * ph;
  for (let v = bot; v <= top + 1e-9; v += step) s += `<line x1="${L}" x2="${w - R}" y1="${y(v)}" y2="${y(v)}" stroke="#D9D9D9"/><text x="${L - 6}" y="${y(v) + 4}" text-anchor="end" fill="#595959">${+v.toFixed(6)}</text>`;
  const n = cats.length || 1;
  const gw = pw / n;
  if (kind === 'line') {
    series.forEach((se, si) => { s += `<polyline fill="none" stroke="${PALETTE[si % 6]}" stroke-width="2.25" points="${se.vals.map((v, i) => `${L + gw * (i + 0.5)},${y(v)}`).join(' ')}"/>`; });
  } else {
    const bw = (gw * 0.7) / Math.max(1, series.length);
    series.forEach((se, si) => se.vals.forEach((v, i) => { const x = L + gw * i + gw * 0.15 + bw * si; s += `<rect x="${x}" y="${Math.min(y(v), y(0))}" width="${bw - 1}" height="${Math.abs(y(v) - y(0))}" fill="${PALETTE[si % 6]}"/>`; }));
  }
  cats.forEach((c, i) => { s += `<text x="${L + gw * (i + 0.5)}" y="${T + ph + 16}" text-anchor="middle" fill="#595959">${esc(c)}</text>`; });
  let lx = L;
  series.forEach((se, si) => { s += `<rect x="${lx}" y="${h - 16}" width="10" height="10" fill="${PALETTE[si % 6]}"/><text x="${lx + 14}" y="${h - 7}" fill="#595959">${esc(se.name)}</text>`; lx += 24 + String(se.name).length * 8; });
  return `${s}</svg>`;
}
function niceStep(x) { const p = 10 ** Math.floor(Math.log10(x)); const f = x / p; return (f <= 1 ? 1 : f <= 2 ? 2 : f <= 5 ? 5 : 10) * p; }
function chartDialog(kind = 'col') {
  const ta = el('textarea', { rows: 8, style: { width: '100%', font: '12px Consolas, monospace' } }, '항목\t계열 1\t계열 2\n1월\t4.3\t2.4\n2월\t2.5\t4.4\n3월\t3.5\t1.8\n4월\t4.5\t2.8');
  ta.addEventListener('keydown', (e) => e.stopPropagation());
  const type = el('select', {}, [['col', '세로 막대형'], ['bar', '가로 막대형 (세로로 그림)'], ['line', '꺾은선형'], ['pie', '원형']].map(([v, l]) => el('option', { value: v, selected: v === kind }, l)));
  const title = el('input', { type: 'text', value: '차트 제목' });
  const prev = el('div', { style: { border: '1px solid #ddd', marginTop: '8px' } });
  const rows = () => ta.value.trim().split(/\r?\n/).map((l) => l.split(/\t|,/));
  const draw = () => { prev.innerHTML = chartSvg(type.value === 'bar' ? 'col' : type.value, rows(), { title: title.value }); };
  for (const x of [ta, type, title]) x.addEventListener('input', draw);
  openDialog({
    title: '차트 삽입', width: 560,
    body: el('div', {}, el('div', { class: 'dlg-row' }, el('span', {}, '종류'), type, el('span', {}, '제목'), title), el('div', { class: 'muted' }, '데이터 (Excel 에서 복사해 붙여넣거나 탭 · 쉼표로 구분, 첫 줄 = 계열 이름):'), ta, prev),
    buttons: [{ label: '삽입', primary: true, action: () => {
      const svg = chartSvg(type.value === 'bar' ? 'col' : type.value, rows(), { title: title.value });
      run('insertImage', svgUrl(svg), 480, 288, `${title.value} 차트`);
      return true;
    } }, { label: '취소' }],
  });
  draw();
}

/** 단락 위 · 아래로 옮기기 (Alt+Shift+↑/↓): 최상위 블록 */
function moveBlock(dir) {
  const sel = curSel();
  if (typeof sel.f.p[0] !== 'number') return;
  const i = sel.f.p[0];
  const j = i + dir;
  if (j < 0 || j >= S.doc.body.length) return;
  change((d) => { const body = d.list([]); [body[i], body[j]] = [body[j], body[i]]; return { a: { ...sel.a, p: [j, ...sel.a.p.slice(1)] }, f: { ...sel.f, p: [j, ...sel.f.p.slice(1)] } }; });
}

// ───────────── 정보 · 새로운 기능 · 바로 가기 키 ─────────────
const WHATS_NEW = [
  'Word 와 같은 리본 · 단축키 · 키 팁(Alt) · 빠른 실행 도구 모음 · 탐색 창 · 눈금자 · 상태 표시줄',
  '열기: Word(.docx · .doc) · 한글(.hwp · .hwpx) · PDF(편집 가능한 문서로) · 마크다운 · HTML · RTF · ODT · 텍스트',
  '저장: Word(.docx) · 한글(.hwpx) · PDF(인쇄) · 마크다운 · HTML · 텍스트',
  '스타일 · 목차 · 머리글/바닥글 · 각주/미주 · 구역 · 다단 · 표(병합 · 스타일 · 머리글 행 반복) · 수식 · 아이콘 · 차트',
  '변경 내용 추적 · 메모 · 맞춤법(한국어 규칙) · 문서 비교 · 편지 병합 · 참고 문헌 · 색인',
  '마크다운 창: 문서 ↔ 마크다운 양방향 편집',
  'AI 연결 (MCP): Claude 같은 AI 가 문서를 읽고 · 고치고 · 저장 (파일 또는 열린 창)',
];
register({
  modifyStyleDialog: (id) => styleDialog(id ?? blockAt(S.doc, curSel().f.p)?.style ?? 'Normal'),
  newStyleDialog: () => styleDialog(null),
  optionsDialog: () => optionsDialog(),
  insertIcons: () => iconsDialog(),
  insertChartDialog: (kind) => chartDialog(kind),
  share: () => run('backstage', 'share'),
  moveUp: () => moveBlock(-1),
  moveDown: () => moveBlock(1),
  about: () => alertDialog('WIWORD 정보', `WIWORD (위워드) ${APP_VERSION}\nMicrosoft Word 와 같은 화면의 웹 문서 편집기 — WIXEL · WIPOINT 와 같은 가족\n\n열기: .docx .doc .hwp .hwpx .pdf .md .html .rtf .odt .txt\n저장: .docx .hwpx .md .html .txt · PDF(인쇄)\n\n참고한 공개 코드: kordoc (MIT, 한글 문서 구조) · hwp-parser (HWP 5.0 예제) · wordcraft (화면 구성 연구)\n의존성 없는 순수 JavaScript — 서버 없이 브라우저에서 동작합니다.`),
  whatsNew: () => openDialog({ title: '새로운 기능', width: 520, body: el('ul', {}, WHATS_NEW.map((t) => el('li', { style: { margin: '6px 0' } }, t))), buttons: [{ label: '닫기', primary: true }] }),
  shortcuts: () => openDialog({ title: '바로 가기 키', width: 560, body: el('table', { class: 'recent' }, SHORTCUTS.map(([k, t]) => el('tr', {}, el('td', { style: { padding: '3px 12px 3px 0' } }, el('b', {}, k)), el('td', {}, t)))), buttons: [{ label: '닫기', primary: true }] }),
}, { modifyStyleDialog: { label: '스타일 수정', args: ['styleId'] }, newStyleDialog: { label: '새 스타일' }, insertIcons: { label: '아이콘' }, insertChartDialog: { label: '차트', args: ['col|line|pie'] }, about: { label: 'WIWORD 정보' } });
