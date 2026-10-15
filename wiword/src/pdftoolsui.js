// PDF 도구 창 — 할 일(카드)을 먼저 고르고, 그 일에 필요한 파일과 옵션만 보여 줌 (이음오피스의 PDF 도구 창을 참고)
// 모든 처리는 이 브라우저 안에서만 하고(AI · 서버 없음), 결과는 항상 새 파일. 계산은 DOM 없는 pdftools.js.
import { register, run } from './state.js';
import { ANNOT_TYPES, analyzePages, canEmbedFont, compactPdf, decryptPdf, deletePages, duplexMerge, extractPages, imagesToPdf, mergePdfs, nUp, overlayPdf, parsePages, pdfInfo, removeAnnotations, reversePages, rotatePages, selectPages, setPdfInfo, splitPdf, stampPdf } from './pdftools.js';
import { readPdf } from './pdf.js';
import { readBytes } from './formats.js';
import { docText } from './model.js';
import { zip } from './zip.js';
import { pickFile, fmtSize, askPassword, saveBytes, openFile } from './fileio.js';
import { fontFileFor, askLocalFontsOnce } from './fontlib.js';
import { el, openDialog, toast, progressOverlay, alertDialog } from './ui.js';

const PAGES_HINT = '예: 1-3, 5, 8- (비우면 모든 페이지)';
const pagesField = (label = '페이지') => ({ name: 'pages', label, value: '', hint: PAGES_HINT });
const PAPER = { type: 'select', options: [['A4', 'A4'], ['A3', 'A3'], ['B4', 'B4'], ['B5', 'B5'], ['Letter', 'Letter'], ['Legal', 'Legal']] };

/** 도구 목록 — files: 'one' | 'many' | 'two' | 'images' */
export const PDF_TOOLS = [
  { id: 'merge', cat: '페이지 정리', label: 'PDF 병합', desc: '여러 PDF 를 하나로. 파일마다 책갈피를 만들고 원래 책갈피는 그 아래로', files: 'many', fields: [{ name: 'bookmarks', label: '파일마다 책갈피', type: 'checkbox', value: true }] },
  { id: 'split', cat: '페이지 정리', label: 'PDF 분할', desc: '정한 페이지 수마다, 또는 범위마다 나눔 (여러 파일은 ZIP 하나로)', files: 'one', fields: [{ name: 'every', label: '몇 페이지마다', type: 'number', value: 1 }, { name: 'ranges', label: '또는 범위', value: '', hint: '예: 1-3, 4-10, 11- (쉼표마다 파일 하나, 비우면 위 쪽수로)' }] },
  { id: 'extract', cat: '페이지 정리', label: '페이지 추출', desc: '고른 페이지만 새 파일로 (지정한 순서대로)', files: 'one', fields: [pagesField('남길 페이지')] },
  { id: 'delete', cat: '페이지 정리', label: '페이지 삭제', desc: '고른 페이지를 뺀 새 파일', files: 'one', fields: [{ ...pagesField('지울 페이지'), hint: '예: 2, 5-7' }] },
  { id: 'reorder', cat: '페이지 정리', label: '페이지 순서 바꾸기', desc: '적은 순서대로 다시 묶음 (같은 페이지를 두 번 넣을 수도 있음)', files: 'one', fields: [{ name: 'pages', label: '새 순서', value: '', hint: '예: 3, 1, 2, 4-' }] },
  { id: 'reverse', cat: '페이지 정리', label: '역순 정렬', desc: '스캔 순서가 거꾸로 된 문서를 바로잡음', files: 'one', fields: [] },
  { id: 'rotate', cat: '페이지 정리', label: '페이지 회전', desc: '고른 페이지를 시계 방향으로 돌림', files: 'one', fields: [{ name: 'deg', label: '각도', type: 'select', options: [['90', '오른쪽으로 90°'], ['180', '180°'], ['270', '왼쪽으로 90°']] }, pagesField()] },
  { id: 'blank', cat: '페이지 정리', label: '빈 페이지 찾아 지우기', desc: '빈 페이지 후보를 보여 주고, 고른 것만 지움 (스캔한 빈 종이도 찾음)', files: 'one', fields: [] },
  { id: 'nup', cat: '인쇄 준비', label: '페이지 모아찍기', desc: '한 장에 2 · 4 · 6 · 8 · 9 · 16 페이지', files: 'one', fields: [{ name: 'n', label: '한 장에', type: 'select', options: [['2', '2쪽'], ['4', '4쪽'], ['6', '6쪽'], ['8', '8쪽'], ['9', '9쪽'], ['16', '16쪽']] }, { name: 'paper', label: '용지', ...PAPER }, { name: 'order', label: '순서', type: 'select', options: [['row', '가로 먼저 (Z)'], ['col', '세로 먼저 (N)']] }, { name: 'border', label: '테두리', type: 'checkbox', value: false }] },
  { id: 'duplex', cat: '인쇄 준비', label: '양면 인쇄 맞추기', desc: '파일마다 앞면에서 시작하도록 빈 페이지를 넣어 하나로 합침', files: 'many', fields: [{ name: 'perSheet', label: '모아찍기', type: 'select', options: [['1', '안 함'], ['2', '한 장에 2쪽 (장 단위로 맞춤)'], ['4', '한 장에 4쪽 (장 단위로 맞춤)']] }] },
  { id: 'stamp', cat: '꾸미기', label: '쪽 번호 · 워터마크', desc: '쪽 번호, 비스듬한 워터마크, 머리글 · 바닥글 글자를 넣음', files: 'one', fields: [
    { name: 'numbers', label: '쪽 번호', type: 'checkbox', value: true }, { name: 'numFormat', label: '번호 모양', value: '- {n} -', hint: '{n} = 쪽 번호, {total} = 전체 쪽수' }, { name: 'numPos', label: '위치', type: 'select', options: [['bc', '아래 가운데'], ['br', '아래 오른쪽'], ['bl', '아래 왼쪽'], ['tc', '위 가운데'], ['tr', '위 오른쪽'], ['tl', '위 왼쪽']] }, { name: 'numStart', label: '시작 번호', type: 'number', value: 1 }, { name: 'numFrom', label: '번호를 붙일 첫 페이지', type: 'number', value: 1 },
    { name: 'wm', label: '워터마크 글', value: '', hint: '예: 대외비, 사본, DRAFT' }, { name: 'wmOpacity', label: '워터마크 진하기 %', type: 'number', value: 20 },
    { name: 'header', label: '머리글 (가운데)', value: '' }, { name: 'footer', label: '바닥글 (왼쪽)', value: '', hint: '{date} = 오늘 날짜, {file} = 파일 이름' }, pagesField('적용할 페이지')] },
  { id: 'overlay', cat: '꾸미기', label: '배경 겹치기', desc: '다른 PDF 한 페이지(레터헤드 · 양식)를 모든 페이지 뒤(또는 앞)에. 돌린 페이지에서도 똑바로', files: 'two', fields: [{ name: 'behind', label: '놓을 곳', type: 'select', options: [['1', '내용 뒤 (배경)'], ['0', '내용 앞 (덮기)']] }, { name: 'bgPage', label: '배경 PDF 의 몇 페이지', type: 'number', value: 1 }, { name: 'stretch', label: '페이지 크기에 맞춰 늘이기', type: 'checkbox', value: false }, pagesField('적용할 페이지')] },
  { id: 'annots', cat: '정리하기', label: '주석 한꺼번에 지우기', desc: '메모 · 형광펜 · 도장 등을 종류별로 지움. 링크 · 입력란은 기본으로 남김 (지운 글은 파일에 남지 않음)', files: 'one', fields: [pagesField()] },
  { id: 'compact', cat: '정리하기', label: '용량 줄이기 (정리)', desc: '쓰지 않는 개체를 버리고 압축 안 된 부분을 압축해 다시 씀', files: 'one', fields: [] },
  { id: 'decrypt', cat: '정리하기', label: '암호 해제', desc: '열기 암호를 아는 PDF 를 암호 없는 사본으로 저장 (권한 제한도 풀림)', files: 'one', fields: [] },
  { id: 'info', cat: '정리하기', label: '문서 정보 바꾸기', desc: '제목 · 만든 이 · 주제 · 키워드', files: 'one', fields: [{ name: 'title', label: '제목', value: '' }, { name: 'author', label: '만든 이', value: '' }, { name: 'subject', label: '주제', value: '' }, { name: 'keywords', label: '키워드', value: '' }] },
  { id: 'images', cat: '변환', label: '이미지로 PDF 만들기', desc: 'JPG · PNG (그 밖의 그림은 PNG 로 바꿔서) 여러 장을 PDF 하나로. 사진의 회전 정보 반영', files: 'images', fields: [{ name: 'paper', label: '용지', type: 'select', options: [['A4', 'A4'], ['fit', '그림 크기 그대로'], ['A3', 'A3'], ['B5', 'B5'], ['Letter', 'Letter']] }, { name: 'margin', label: '여백 (mm)', type: 'number', value: 8 }] },
  { id: 'text', cat: '변환', label: '텍스트로 저장', desc: 'PDF 의 글을 문단으로 이어 .txt 로 저장', files: 'one', fields: [] },
  { id: 'edit', cat: '변환', label: '문서로 열어 편집', desc: 'PDF 를 문서로 바꿔 이 편집기에서 열기 (표 · 두 단 인식)', files: 'one', fields: [] },
];

/** 파일 하나 읽기 + 암호가 있으면 물어봄 → { name, bytes, password, pages } | null */
async function loadPdfFile(f) {
  const bytes = new Uint8Array(await f.arrayBuffer());
  let password = ''; let wrong = false;
  for (;;) {
    try {
      const info = pdfInfo(bytes, { password });
      return { name: f.name, bytes, password, info };
    } catch (e) {
      if (e.code !== 'PASSWORD') { alertDialog('PDF 도구', `'${f.name}': ${e.message}`); return null; }
      const pw = await askPassword(f.name, wrong || !!e.wrong);
      if (pw == null) return null;
      password = pw; wrong = true;
    }
  }
}

/** 브라우저가 바로 넣을 수 없는 그림(GIF · WebP · BMP …)은 PNG 로 바꿈 */
async function imageBytes(f) {
  const bytes = new Uint8Array(await f.arrayBuffer());
  const jpg = bytes[0] === 0xff && bytes[1] === 0xd8; const png = bytes[0] === 0x89 && bytes[1] === 0x50;
  if (jpg || png) return bytes;
  const bmp = await createImageBitmap(new Blob([bytes]));
  const c = document.createElement('canvas'); c.width = bmp.width; c.height = bmp.height;
  c.getContext('2d').drawImage(bmp, 0, 0);
  const blob = await new Promise((r) => c.toBlob(r, 'image/png'));
  return new Uint8Array(await blob.arrayBuffer());
}

/** 그림만 있는 페이지가 거의 흰색인지 (스캔한 빈 종이) */
async function whitePages(bytes, password, idxs) {
  const out = new Set();
  if (!idxs.length) return out;
  const stats = new Map(); // dataURL → 흰 비율
  const ratio = (rgba) => { let w = 0; const n = rgba.length / 4; for (let i = 0; i < rgba.length; i += 4) if (rgba[i + 3] < 16 || (rgba[i] > 235 && rgba[i + 1] > 235 && rgba[i + 2] > 235)) w++; return w / Math.max(1, n); };
  let k = 0;
  const { pres } = readPdf(bytes, { password, maxPages: 2000, encodeImage: (w, h, rgba) => { const key = `w:${k++}`; stats.set(key, ratio(rgba)); return key; } });
  for (const i of idxs) {
    const slide = pres.slides[i];
    if (!slide) continue;
    const ims = slide.objects.filter((o) => o.type === 'image' && o.media);
    if (!ims.length) continue;
    let all = true;
    for (const o of ims) {
      const src = pres.media[o.media];
      let r = stats.get(src);
      if (r == null && typeof src === 'string' && src.startsWith('data:image/')) {
        try {
          const bmp = await createImageBitmap(await (await fetch(src)).blob());
          const s = Math.min(1, 200 / Math.max(bmp.width, bmp.height));
          const c = document.createElement('canvas'); c.width = Math.max(1, Math.round(bmp.width * s)); c.height = Math.max(1, Math.round(bmp.height * s));
          const g = c.getContext('2d'); g.drawImage(bmp, 0, 0, c.width, c.height);
          r = ratio(g.getImageData(0, 0, c.width, c.height).data);
        } catch { r = 0; }
      }
      if (!(r >= 0.995)) { all = false; break; }
    }
    if (all) out.add(i);
  }
  return out;
}

/** 결과 저장 (여러 파일이면 ZIP) */
async function saveResults(results, base) {
  if (results.length === 1) {
    const r = results[0];
    const ok = await saveBytes(r.bytes, r.name, r.mime ? { mime: r.mime, ext: r.name.slice(r.name.lastIndexOf('.')), description: r.desc ?? '파일' } : {});
    if (ok !== false) toast(`저장했습니다: ${r.name} (${fmtSize(r.bytes.length)})`);
    return;
  }
  const z = zip(Object.fromEntries(results.map((r) => [r.name, r.bytes])));
  const ok = await saveBytes(z, `${base}.zip`, { mime: 'application/zip', ext: '.zip', description: 'ZIP 압축 파일' });
  if (ok !== false) toast(`파일 ${results.length}개를 ZIP 으로 저장했습니다 (${fmtSize(z.length)})`);
}

/** PDF 도구 창 (toolId 를 주면 그 도구를 바로 엶) */
export function pdfToolsDialog(toolId = null) {
  const side = el('div', { class: 'pt-side' });
  const main = el('div', { class: 'pt-main' });
  const dlg = openDialog({ title: 'PDF 도구', width: 860, body: el('div', { class: 'pt-wrap' }, side, main) });
  const cats = [...new Set(PDF_TOOLS.map((t) => t.cat))];
  const buttons = new Map();
  for (const c of cats) {
    side.append(el('div', { class: 'pt-cat' }, c));
    for (const t of PDF_TOOLS.filter((x) => x.cat === c)) {
      const b = el('button', { class: 'pt-item', title: t.desc, onclick: () => show(t) }, t.label);
      buttons.set(t.id, b);
      side.append(b);
    }
  }
  const home = () => {
    main.replaceChildren(el('h3', {}, 'PDF 도구'), el('p', { class: 'muted' }, '할 일을 왼쪽에서 고르세요. 파일은 이 브라우저 안에서만 처리하고(서버 · AI 를 쓰지 않음), 결과는 항상 새 파일로 저장합니다. 원본은 바뀌지 않습니다.'),
      el('div', { class: 'pt-cards' }, PDF_TOOLS.map((t) => el('button', { class: 'pt-card', onclick: () => show(t) }, el('b', {}, t.label), el('span', {}, t.desc)))),
      el('h4', {}, '이렇게 써 보세요'),
      el('ul', { class: 'pt-tips' },
        el('li', {}, '여러 문서를 양면 인쇄하는데 문서마다 앞면에서 시작해야 할 때 → 양면 인쇄 맞추기 (+ 모아찍기)'),
        el('li', {}, '스캔한 문서의 순서가 거꾸로이거나 몇 장이 뒤집혔을 때 → 역순 정렬 · 페이지 회전'),
        el('li', {}, '회사 레터헤드를 모든 페이지에 입힐 때 → 배경 겹치기'),
        el('li', {}, '내보내기 전에 검토 메모를 정리할 때 → 주석 한꺼번에 지우기 (링크 · 입력란은 남김)'),
        el('li', {}, '사진 여러 장을 PDF 하나로 제출할 때 → 이미지로 PDF 만들기')));
  };
  let files = [];
  const show = (t) => {
    for (const [id, b] of buttons) b.classList.toggle('on', id === t.id);
    files = [];
    const list = el('div', { class: 'pt-files' });
    const inputs = {};
    const fieldEls = t.fields.map((f) => {
      let input;
      if (f.type === 'select') input = el('select', {}, f.options.map(([v, l]) => el('option', { value: v }, l)));
      else if (f.type === 'checkbox') input = el('input', { type: 'checkbox', checked: !!f.value });
      else input = el('input', { type: f.type === 'number' ? 'number' : 'text', value: f.value ?? '', placeholder: f.hint ?? '' });
      inputs[f.name] = input;
      return el('label', { class: 'pt-field' }, el('span', {}, f.label), input, f.hint && f.type !== 'number' ? null : f.hint ? el('small', { class: 'muted' }, f.hint) : null);
    });
    // 주석 종류 고르기
    let annotBox = null;
    if (t.id === 'annots') {
      annotBox = el('div', { class: 'pt-annots' }, ANNOT_TYPES.map(([k, l]) => el('label', {}, el('input', { type: 'checkbox', value: k, checked: k !== 'Link' && k !== 'Widget' }), ` ${l}`)));
    }
    const blankBox = el('div', { class: 'pt-blank' });
    const renderFiles = () => {
      list.replaceChildren(...files.map((f, i) => el('div', { class: 'pt-file' },
        el('span', { class: 'pt-fname' }, `${t.id === 'overlay' ? (i === 0 ? '문서: ' : '배경: ') : ''}${f.name}`),
        el('span', { class: 'muted' }, f.info ? `${f.info.pages.length}쪽 · ${fmtSize(f.bytes.length)}${f.password ? ' · 암호' : ''}${annotSummary(f)}` : fmtSize(f.bytes.length)),
        t.files === 'many' || t.files === 'images' ? el('span', {},
          el('button', { class: 'btn small', title: '위로', disabled: i === 0, onclick: () => { [files[i - 1], files[i]] = [files[i], files[i - 1]]; renderFiles(); } }, '▲'),
          el('button', { class: 'btn small', title: '아래로', disabled: i === files.length - 1, onclick: () => { [files[i + 1], files[i]] = [files[i], files[i + 1]]; renderFiles(); } }, '▼')) : null,
        el('button', { class: 'btn small', title: '빼기', onclick: () => { files.splice(i, 1); renderFiles(); } }, '✕'))));
      if (!files.length) list.append(el('div', { class: 'muted pt-drop' }, t.files === 'images' ? '그림 파일을 추가하거나 여기로 끌어 놓으세요' : 'PDF 파일을 추가하거나 여기로 끌어 놓으세요'));
    };
    const annotSummary = (f) => {
      if (t.id !== 'annots' || !f.info) return '';
      const a = Object.entries(f.info.annots ?? {}).map(([k, n]) => `${ANNOT_TYPES.find((x) => x[0] === k)?.[1] ?? k} ${n}`);
      return a.length ? ` · 주석: ${a.join(', ')}` : ' · 주석 없음';
    };
    const max = t.files === 'one' ? 1 : t.files === 'two' ? 2 : 999;
    const addFiles = async (fl) => {
      for (const f of fl) {
        if (files.length >= max) { if (max === 1) files = []; else break; }
        if (t.files === 'images') {
          if (!/^image\//.test(f.type) && !/\.(jpe?g|png|gif|webp|bmp|avif|heic)$/i.test(f.name)) continue;
          try { files.push({ name: f.name, bytes: await imageBytes(f) }); } catch { toast(`'${f.name}' 그림을 읽지 못했습니다`); }
        } else {
          const r = await loadPdfFile(f);
          if (r) files.push(r);
        }
      }
      renderFiles();
      blankBox.replaceChildren();
    };
    const add = el('button', { class: 'btn', onclick: async () => {
      const fl = await pickFile(t.files === 'images' ? 'image/*' : '.pdf,application/pdf', max > 1);
      if (fl) await addFiles(Array.isArray(fl) ? fl : [fl]);
    } }, t.files === 'images' ? '그림 추가…' : t.files === 'two' ? (files.length ? '배경 PDF 추가…' : 'PDF 추가…') : 'PDF 추가…');
    list.addEventListener('dragover', (e) => { e.preventDefault(); e.stopPropagation(); });
    list.addEventListener('drop', (e) => { e.preventDefault(); e.stopPropagation(); addFiles([...(e.dataTransfer?.files ?? [])]); });
    const go = el('button', { class: 'btn primary', onclick: () => execute(t, files, inputs, annotBox, blankBox) }, t.id === 'blank' ? '빈 페이지 찾기' : t.id === 'edit' ? '열기' : '실행하고 저장…');
    main.replaceChildren(...[el('h3', {}, t.label), el('p', { class: 'muted' }, t.desc),
      el('div', { class: 'pt-row' }, add, el('span', { class: 'muted' }, t.files === 'two' ? '첫 파일 = 문서, 둘째 파일 = 배경' : t.files === 'many' ? '위에서부터 차례대로 합칩니다' : '')),
      list, fieldEls.length ? el('div', { class: 'pt-fields' }, fieldEls) : null, annotBox, blankBox,
      el('div', { class: 'pt-row' }, go, el('span', { class: 'muted' }, '결과는 새 파일로 저장합니다.'))].filter(Boolean));
    renderFiles();
  };
  const val = (inputs, k) => { const i = inputs[k]; return !i ? undefined : i.type === 'checkbox' ? i.checked : i.value; };
  const execute = async (t, fl, inputs, annotBox, blankBox) => {
    const need = t.files === 'many' ? 1 : t.files === 'two' ? 2 : 1;
    if (fl.length < need) { toast(t.files === 'two' ? '문서 PDF 와 배경 PDF 를 차례로 추가하세요' : '파일을 먼저 추가하세요'); return; }
    const f = fl[0];
    const base = (f.name ?? '문서').replace(/\.[^.]+$/, '');
    const bar = progressOverlay(`${t.label} 하는 중…`);
    await new Promise((r) => setTimeout(r, 30));
    const pdf = (bytes, suffix) => ({ bytes, name: `${base}_${suffix}.pdf` });
    try {
      const opt = { password: f.password };
      let out = null;
      switch (t.id) {
        case 'merge': out = [{ bytes: await mergePdfs(fl.map((x) => ({ bytes: x.bytes, name: x.name.replace(/\.pdf$/i, ''), password: x.password })), { bookmarks: val(inputs, 'bookmarks') }), name: `${base}_병합.pdf` }]; break;
        case 'duplex': out = [{ bytes: await duplexMerge(fl.map((x) => ({ bytes: x.bytes, name: x.name.replace(/\.pdf$/i, ''), password: x.password })), { perSheet: Number(val(inputs, 'perSheet')) }), name: `${base}_양면인쇄.pdf` }]; break;
        case 'split': {
          const parts = await splitPdf(f.bytes, { ...opt, every: Number(val(inputs, 'every')) || 1, ranges: val(inputs, 'ranges') || null, base });
          out = parts; break;
        }
        case 'extract': out = [pdf(await extractPages(f.bytes, val(inputs, 'pages'), opt), '추출')]; break;
        case 'delete': if (!String(val(inputs, 'pages')).trim()) throw new Error('지울 페이지를 적으세요 (예: 2, 5-7)'); out = [pdf(await deletePages(f.bytes, val(inputs, 'pages'), opt), '삭제')]; break;
        case 'reorder': out = [pdf(await selectPages(f.bytes, { ...opt, pages: parsePages(val(inputs, 'pages'), f.info.pages.length) }), '순서')]; break;
        case 'reverse': out = [pdf(await reversePages(f.bytes, opt), '역순')]; break;
        case 'rotate': out = [pdf(await rotatePages(f.bytes, Number(val(inputs, 'deg')), val(inputs, 'pages'), opt), '회전')]; break;
        case 'nup': out = [pdf(await nUp(f.bytes, { ...opt, n: Number(val(inputs, 'n')), paper: val(inputs, 'paper'), order: val(inputs, 'order'), border: val(inputs, 'border') }), `모아찍기${val(inputs, 'n')}`)]; break;
        case 'overlay': out = [pdf(await overlayPdf(f.bytes, fl[1].bytes, { ...opt, bgPassword: fl[1].password, behind: val(inputs, 'behind') === '1', bgPage: Number(val(inputs, 'bgPage')) || 1, stretch: val(inputs, 'stretch'), pages: val(inputs, 'pages') }), '배경')]; break;
        case 'stamp': {
          const texts = [val(inputs, 'wm'), val(inputs, 'header'), val(inputs, 'footer'), val(inputs, 'numFormat')].join('');
          let font = null;
          if (/[^\x20-\x7e]/.test(texts)) {
            // 한글 글자는 이 컴퓨터의 글꼴을 넣음 (안 되면 보는 프로그램의 한글 글꼴)
            await askLocalFontsOnce().catch(() => false);
            for (const fam of ['맑은 고딕', 'Malgun Gothic', 'Apple SD Gothic Neo', 'Noto Sans KR', 'NanumGothic', '나눔고딕']) { const r = await fontFileFor(fam).catch(() => null); if (r?.bytes && canEmbedFont(r.bytes, r.ps)) { font = r; break; } }
          }
          out = [pdf(await stampPdf(f.bytes, {
            ...opt, file: f.name, font: font?.bytes, fontName: font?.ps, pages: val(inputs, 'pages'),
            numbers: val(inputs, 'numbers') ? { format: val(inputs, 'numFormat') || '{n}', pos: val(inputs, 'numPos'), start: Number(val(inputs, 'numStart')) || 1, from: Number(val(inputs, 'numFrom')) || 1 } : null,
            watermark: val(inputs, 'wm') ? { text: val(inputs, 'wm'), opacity: Math.max(0.03, Math.min(1, (Number(val(inputs, 'wmOpacity')) || 20) / 100)) } : null,
            header: val(inputs, 'header') ? { center: val(inputs, 'header') } : null,
            footer: val(inputs, 'footer') ? { left: val(inputs, 'footer') } : null,
          }), '도장')];
          break;
        }
        case 'annots': {
          const types = [...annotBox.querySelectorAll('input:checked')].map((i) => i.value);
          if (!types.length) throw new Error('지울 주석 종류를 하나 이상 고르세요');
          const r = await removeAnnotations(f.bytes, { ...opt, types: [...types, 'Popup'], pages: val(inputs, 'pages') });
          toast(`주석 ${r.removed}개를 지웠습니다`);
          out = [pdf(r, '주석정리')]; break;
        }
        case 'compact': {
          const r = await compactPdf(f.bytes, opt);
          toast(`${fmtSize(f.bytes.length)} → ${fmtSize(r.length)}${r.length >= f.bytes.length ? ' (더 줄일 것이 없었습니다)' : ''}`);
          out = [pdf(r.length < f.bytes.length ? r : r, '정리')]; break;
        }
        case 'decrypt': if (!f.info.encrypted) toast('암호가 없는 PDF 입니다. 그대로 다시 씁니다.'); out = [pdf(await decryptPdf(f.bytes, f.password), '암호해제')]; break;
        case 'info': out = [pdf(await setPdfInfo(f.bytes, { title: val(inputs, 'title'), author: val(inputs, 'author'), subject: val(inputs, 'subject'), keywords: val(inputs, 'keywords') }, opt), '정보')]; break;
        case 'images': {
          const paper = val(inputs, 'paper');
          out = [{ bytes: await imagesToPdf(fl, { paper, margin: (Number(val(inputs, 'margin')) || 0) * 72 / 25.4 }), name: `${fl.length === 1 ? base : '이미지'}.pdf` }]; break;
        }
        case 'text': {
          const r = readBytes(f.bytes, f.name, { password: f.password, encodeImage: () => null });
          const txt = new TextEncoder().encode(`${String.fromCharCode(0xfeff)}${docText(r.doc)}`);
          out = [{ bytes: txt, name: `${base}.txt`, mime: 'text/plain', desc: '텍스트' }]; break;
        }
        case 'edit': bar.close(); dlg.close(); await openFile(new File([f.bytes], f.name, { type: 'application/pdf' })); return;
        case 'blank': {
          const pages = analyzePages(f.bytes, opt);
          const white = await whitePages(f.bytes, f.password, pages.filter((p) => p.imageOnly).map((p) => p.index));
          const cands = pages.filter((p) => p.blank || white.has(p.index));
          bar.close();
          if (!cands.length) { blankBox.replaceChildren(el('p', {}, '빈 페이지를 찾지 못했습니다.')); return; }
          const boxes = cands.map((p) => el('label', {}, el('input', { type: 'checkbox', value: String(p.index), checked: true }), ` ${p.index + 1}쪽 ${p.blank ? '(내용 없음)' : '(흰 그림만 — 스캔한 빈 종이)'}`));
          blankBox.replaceChildren(el('p', {}, `빈 페이지 후보 ${cands.length}개 (전체 ${pages.length}쪽). 지울 것을 확인하세요.`), el('div', { class: 'pt-annots' }, boxes),
            el('button', { class: 'btn primary', onclick: async () => {
              const del = new Set(boxes.map((b) => b.querySelector('input')).filter((i) => i.checked).map((i) => Number(i.value)));
              if (!del.size) { toast('지울 페이지를 고르세요'); return; }
              const keep = pages.map((p) => p.index).filter((i) => !del.has(i));
              await saveResults([pdf(await selectPages(f.bytes, { ...opt, pages: keep }), '빈페이지삭제')], base);
            } }, '고른 페이지 지우고 저장…'));
          return;
        }
        default: break;
      }
      bar.close();
      if (out?.length) await saveResults(out, `${base}_${t.label}`);
    } catch (e) {
      alertDialog('PDF 도구', e.message || String(e));
    } finally { bar.close(); }
  };
  const start = toolId && PDF_TOOLS.find((t) => t.id === toolId);
  if (start) show(start); else home();
  return dlg;
}

register({ pdfTools: (id) => pdfToolsDialog(typeof id === 'string' ? id : null) }, { pdfTools: { label: 'PDF 도구', desc: '병합 · 분할 · 추출 · 회전 · 모아찍기 · 배경 · 쪽 번호 · 주석 지우기 · 이미지로 PDF …' } });
