// 대화 상자 (글꼴 · 단락 · 페이지 설정 · 찾기/바꾸기 · 표 속성 · 기호 · 수식 · 날짜 · 링크 · 책갈피 …)
import { S, run, register, emit, change, curSel } from './state.js';
import { el, openDialog, formDialog, toast, closeMenus } from './ui.js';
import {
  blockAt, effPara, effRun, propsAt, paraLen, posEq, orderPos, cm, pxToCm, PAPERS, MARGINS, para, uid, walkParas, paraPlain, runLen, docStats, styleChain, sections, firstPos, splitRuns, normalizeRuns, insertRuns, clone, headings, fontName,
} from './model.js';
import { selProp, selParas, sectionAtTop, bodyWidthAt, pageOfBlock } from './commands.js';
import { setSel } from './editor.js';
import { V, cellSelection } from './view.js';
import { parseLatex, toMathML } from './math.js';

const sel = () => curSel();
const f1 = (v) => Math.round(v * 10) / 10;
const f2 = (v) => Math.round(v * 100) / 100;
const num = (v, d = 0) => { const n = Number(v); return Number.isFinite(n) ? n : d; };
const row = (label, input, unit) => el('label', { class: 'dlg-row' }, el('span', {}, label), input, unit ? el('small', {}, unit) : null);
const sel2 = (opts, val) => el('select', {}, opts.map(([v, l]) => el('option', { value: v, selected: String(v) === String(val) }, l)));
const numIn = (v, step = 0.1, w = 70) => el('input', { type: 'number', step, value: v, style: { width: `${w}px` } });
const chk = (label, on) => { const i = el('input', { type: 'checkbox', checked: !!on }); return { i, node: el('label', { class: 'dlg-chk' }, i, ' ', label) }; };
const tabs = (pages) => {
  const head = el('div', { class: 'dlg-tabs' });
  const body = el('div', { class: 'dlg-tabbody' });
  pages.forEach(([label, node], i) => {
    const b = el('button', { type: 'button', class: i === 0 ? 'on' : '', onclick: () => { for (const x of head.children) x.classList.remove('on'); b.classList.add('on'); body.replaceChildren(node); } }, label);
    head.append(b);
  });
  body.append(pages[0][1]);
  return { node: el('div', {}, head, body), show: (i) => head.children[i]?.click() };
};
const FONT_LIST = ['맑은 고딕', '굴림', '돋움', '바탕', '궁서', '함초롬바탕', '함초롬돋움', '나눔고딕', '나눔명조', 'Noto Sans KR', 'Noto Serif KR', 'Pretendard', 'Arial', 'Calibri', 'Cambria', 'Georgia', 'Times New Roman', 'Verdana', 'Consolas', 'Courier New'];

// ───────────── 글꼴 ─────────────
register({
  fontDialog: () => {
    const p = blockAt(S.doc, sel().f.p);
    const rp = { ...effRun(S.doc, p, propsAt(p, posEq(sel().a, sel().f) ? sel().f.o : orderPos(sel().a, sel().f)[0].o + 1)), ...(S.typing ?? {}) };
    const dl = el('datalist', { id: 'fd-fonts' }, FONT_LIST.map((f) => el('option', { value: f })));
    const lat0 = rp.font ? fontName(S.doc, rp.font) : ''; const ea0 = rp.fontEa ? fontName(S.doc, rp.fontEa, true) : lat0;
    const fLat = el('input', { list: 'fd-fonts', value: lat0 });
    const fEa = el('input', { list: 'fd-fonts', value: ea0 });
    const style = sel2([['n', '보통'], ['i', '기울임꼴'], ['b', '굵게'], ['bi', '굵은 기울임꼴']], `${rp.b ? 'b' : ''}${rp.i ? 'i' : ''}` || 'n');
    const size = numIn(rp.sz ?? 10, 0.5);
    const color = el('input', { type: 'color', value: rp.color && rp.color.startsWith('#') ? rp.color : '#000000' });
    const und = sel2([['', '(없음)'], ['single', '실선'], ['double', '이중'], ['thick', '굵게'], ['dotted', '점선'], ['dash', '파선'], ['wave', '물결']], rp.u === true ? 'single' : rp.u ?? '');
    const effects = [['s', '취소선'], ['dstrike', '이중 취소선'], ['sup', '위 첨자'], ['sub', '아래 첨자'], ['smallCaps', '작은 대문자'], ['caps', '모두 대문자'], ['hidden', '숨김']].map(([k, l]) => [k, chk(l, k === 'sup' ? rp.va === 'super' : k === 'sub' ? rp.va === 'sub' : rp[k])]);
    const scale = numIn(rp.scale ?? 100, 1);
    const spc = numIn(rp.spc ?? 0, 0.1);
    const pos = numIn(rp.pos ?? 0, 0.5);
    const preview = el('div', { class: 'dlg-preview' }, 'WIWORD 위워드 AaBb');
    const upd = () => { preview.style.fontFamily = `'${fLat.value || fEa.value || '맑은 고딕'}','${fEa.value || '맑은 고딕'}'`; preview.style.fontSize = `${num(size.value, 10)}pt`; preview.style.fontWeight = style.value.includes('b') ? 'bold' : 'normal'; preview.style.fontStyle = style.value.includes('i') ? 'italic' : 'normal'; preview.style.color = color.value; preview.style.textDecoration = [und.value ? 'underline' : '', effects[0][1].i.checked || effects[1][1].i.checked ? 'line-through' : ''].join(' ') || 'none'; preview.style.letterSpacing = `${num(spc.value)}pt`; };
    const p1 = el('div', {}, dl, el('div', { class: 'dlg-grid' }, row('한글 글꼴:', fEa), row('영어 글꼴:', fLat), row('글꼴 스타일:', style), row('크기:', size, 'pt'), row('글꼴 색:', color), row('밑줄 스타일:', und), el('div', { class: 'dlg-sec' }, '효과'), el('div', { class: 'dlg-chks' }, effects.map(([, c]) => c.node))));
    const p2 = el('div', { class: 'dlg-grid' }, row('장평:', scale, '%'), row('자간:', spc, 'pt (+ 넓게, − 좁게)'), row('위치:', pos, 'pt (+ 올림)'));
    const t = tabs([['글꼴', p1], ['고급', p2]]);
    const body = el('div', {}, t.node, el('div', { class: 'dlg-sec' }, '미리 보기'), preview);
    body.addEventListener('input', upd); body.addEventListener('change', upd); upd();
    openDialog({ title: '글꼴', width: 520, body, buttons: [{ label: '확인', primary: true, action: () => {
      const props = { font: fLat.value === lat0 ? rp.font : fLat.value || fEa.value || undefined, fontEa: fEa.value === ea0 ? rp.fontEa : fEa.value || fLat.value || undefined, sz: num(size.value, 10), b: style.value.includes('b') || undefined, i: style.value.includes('i') || undefined, color: color.value.toUpperCase() === '#000000' && !rp.color ? undefined : color.value.toUpperCase(), u: und.value || undefined, scale: num(scale.value, 100) === 100 ? undefined : num(scale.value), spc: num(spc.value) || undefined, pos: num(pos.value) || undefined };
      for (const [k, c] of effects) { if (k === 'sup') { if (c.i.checked) props.va = 'super'; continue; } if (k === 'sub') { if (c.i.checked) props.va = 'sub'; continue; } props[k] = c.i.checked || undefined; }
      if (!effects[2][1].i.checked && !effects[3][1].i.checked) props.va = undefined;
      if (!fLat.value && !fEa.value) { delete props.font; delete props.fontEa; }
      run('__applyRun', props);
    } }, { label: '취소' }] });
  },
});

// ───────────── 단락 ─────────────
register({
  paragraphDialog: () => {
    const paths = selParas();
    const p = blockAt(S.doc, paths[0] ?? sel().f.p);
    const ep = effPara(S.doc, p);
    const ind = ep.ind ?? {}; const spc = ep.spc ?? {};
    const align = sel2([['left', '왼쪽'], ['center', '가운데'], ['right', '오른쪽'], ['justify', '양쪽'], ['distribute', '균등 분할']], ep.align ?? 'left');
    const outline = sel2([[9, '본문 텍스트'], ...Array.from({ length: 9 }, (_, i) => [i, `수준 ${i + 1}`])], ep.outline ?? 9);
    const iL = numIn(f2(pxToCm(ind.l ?? 0))); const iR = numIn(f2(pxToCm(ind.r ?? 0)));
    const specialKind = ind.hanging ? 'hanging' : ind.first ? 'first' : 'none';
    const special = sel2([['none', '(없음)'], ['first', '첫 줄'], ['hanging', '내어쓰기']], specialKind);
    const by = numIn(f2(pxToCm(ind.hanging ?? ind.first ?? 0)));
    const sB = numIn(f1(((spc.before ?? 0) * 3) / 4), 1); const sA = numIn(f1(((spc.after ?? 0) * 3) / 4), 1);
    const rule = spc.rule ?? 'auto';
    const lineKind = sel2([['1', '1줄'], ['1.5', '1.5줄'], ['2', '2줄'], ['multiple', '배수'], ['atLeast', '최소'], ['exact', '고정'], ['pct', '글자에 따라 (%)']], rule === 'exact' ? 'exact' : rule === 'atLeast' ? 'atLeast' : rule === 'pct' ? 'pct' : [1, 1.5, 2].includes(spc.line) ? String(spc.line) : 'multiple');
    const lineVal = numIn(rule === 'exact' || rule === 'atLeast' ? f1(((spc.line ?? 16) * 3) / 4) : rule === 'pct' ? Math.round((spc.line ?? 1.6) * 100) : spc.line ?? 1.08, 0.01);
    const ctx = chk('같은 스타일의 단락 사이에 공백 삽입 안 함', ep.contextual);
    const widow = chk('외톨이줄 제어', ep.widow !== false);
    const keepN = chk('다음 단락과 함께', ep.keepNext);
    const keepL = chk('현재 단락을 나누지 않음', ep.keepLines);
    const pbb = chk('단락 앞에서 페이지 나누기', ep.pageBreakBefore);
    const p1 = el('div', { class: 'dlg-grid' }, el('div', { class: 'dlg-sec' }, '일반'), row('맞춤:', align), row('개요 수준:', outline),
      el('div', { class: 'dlg-sec' }, '들여쓰기'), row('왼쪽:', iL, 'cm'), row('오른쪽:', iR, 'cm'), row('첫 줄:', special), row('값:', by, 'cm'),
      el('div', { class: 'dlg-sec' }, '간격'), row('단락 앞:', sB, 'pt'), row('단락 뒤:', sA, 'pt'), row('줄 간격:', lineKind), row('값:', lineVal), ctx.node);
    const p2 = el('div', { class: 'dlg-grid' }, el('div', { class: 'dlg-sec' }, '페이지 매김'), widow.node, keepN.node, keepL.node, pbb.node);
    const t = tabs([['들여쓰기 및 간격', p1], ['줄 및 페이지 나누기', p2]]);
    openDialog({ title: '단락', width: 480, body: t.node, buttons: [
      { label: '탭...', action: () => { setTimeout(() => run('tabsDialog'), 0); } },
      { label: '확인', primary: true, action: () => {
        const sp = num(by.value);
        const nInd = { l: cm(num(iL.value)), r: cm(num(iR.value)) };
        if (special.value === 'first') nInd.first = cm(sp); else if (special.value === 'hanging') nInd.hanging = cm(sp);
        const lk = lineKind.value;
        const nSpc = { before: (num(sB.value) * 4) / 3, after: (num(sA.value) * 4) / 3 };
        if (['1', '1.5', '2'].includes(lk)) Object.assign(nSpc, { line: Number(lk), rule: 'auto' });
        else if (lk === 'multiple') Object.assign(nSpc, { line: num(lineVal.value, 1), rule: 'auto' });
        else if (lk === 'pct') Object.assign(nSpc, { line: num(lineVal.value, 160) / 100, rule: 'pct' });
        else Object.assign(nSpc, { line: (num(lineVal.value, 12) * 4) / 3, rule: lk });
        run('paragraphProps', { align: align.value, outline: Number(outline.value) === 9 ? undefined : Number(outline.value), ind: nInd, spc: nSpc, contextual: ctx.i.checked || undefined, widow: widow.i.checked, keepNext: keepN.i.checked || undefined, keepLines: keepL.i.checked || undefined, pageBreakBefore: pbb.i.checked || undefined });
      } }, { label: '취소' }] });
  },
  tabsDialog: () => {
    const p = blockAt(S.doc, sel().f.p);
    const ep = effPara(S.doc, p);
    let tabsL = (ep.tabs ?? []).map((t) => ({ ...t }));
    const list = el('div', { class: 'dlg-list' });
    const pos = numIn('', 0.1); const al = sel2([['left', '왼쪽'], ['center', '가운데'], ['right', '오른쪽'], ['decimal', '소수점'], ['bar', '줄']], 'left');
    const lead = sel2([['none', '1 없음'], ['dot', '2 ........'], ['hyphen', '3 --------'], ['underscore', '4 ______'], ['middleDot', '5 ·······']], 'none');
    const def = numIn(f2(pxToCm(S.doc.settings.defaultTab ?? 48)));
    const draw = () => list.replaceChildren(...tabsL.sort((a, b) => a.pos - b.pos).map((t, i) => el('div', { class: 'dlg-item' }, `${f2(pxToCm(t.pos))}cm · ${{ left: '왼쪽', center: '가운데', right: '오른쪽', decimal: '소수점', bar: '줄' }[t.align] ?? t.align}${t.leader && t.leader !== 'none' ? ' · 채움선' : ''}`, el('button', { type: 'button', onclick: () => { tabsL.splice(i, 1); draw(); } }, '지우기'))));
    draw();
    openDialog({ title: '탭', width: 420, body: el('div', { class: 'dlg-grid' }, row('탭 위치:', pos, 'cm'), list, row('기본 탭:', def, 'cm'), row('맞춤:', al), row('채움선:', lead), el('button', { type: 'button', class: 'btn', onclick: () => { if (pos.value === '') return; tabsL = tabsL.filter((t) => Math.abs(t.pos - cm(num(pos.value))) > 1); tabsL.push({ pos: cm(num(pos.value)), align: al.value, leader: lead.value }); draw(); } }, '설정'), el('button', { type: 'button', class: 'btn', onclick: () => { tabsL = []; draw(); } }, '모두 지우기')),
      buttons: [{ label: '확인', primary: true, action: () => { run('setTabs', tabsL.length ? tabsL : undefined); if (num(def.value) > 0) change((d) => { d.doc.settings = { ...d.doc.settings, defaultTab: cm(num(def.value)) }; return S.sel; }); } }, { label: '취소' }] });
  },
  bordersDialog: () => {
    const style = sel2([['single', '실선'], ['double', '이중'], ['dotted', '점선'], ['dashed', '파선'], ['thick', '굵게']], 'single');
    const w = sel2([[0.25, '¼ pt'], [0.5, '½ pt'], [0.75, '¾ pt'], [1, '1 pt'], [1.5, '1½ pt'], [2.25, '2¼ pt'], [3, '3 pt'], [4.5, '4½ pt'], [6, '6 pt']], 0.5);
    const col = el('input', { type: 'color', value: '#000000' });
    const where = sel2([['box', '상자'], ['none', '없음'], ['top', '위쪽'], ['bottom', '아래쪽'], ['left', '왼쪽'], ['right', '오른쪽']], 'box');
    const shd = el('input', { type: 'color', value: '#FFFFFF' });
    const useShd = chk('음영 채우기', false);
    openDialog({ title: '테두리 및 음영', width: 380, body: el('div', { class: 'dlg-grid' }, row('설정:', where), row('스타일:', style), row('색:', col), row('두께:', w), useShd.node, row('채우기:', shd)),
      buttons: [{ label: '확인', primary: true, action: () => { run('paraBorder', where.value, { style: style.value, w: Number(w.value), color: col.value.toUpperCase(), space: 1 }); if (useShd.i.checked) run('paraShading', shd.value.toUpperCase()); } }, { label: '취소' }] });
  },
  defineBulletDialog: () => formDialog('새 글머리 기호 정의', [{ name: 'ch', label: '기호 (글자 하나)', value: '❖' }], (v) => { if (v.ch) run('bullets', [...v.ch][0]); }),
  setNumberValueDialog: () => formDialog('번호 매기기 값 설정', [{ name: 'n', label: '값 설정', type: 'number', value: 1 }], (v) => {
    const pth = sel().f.p;
    const p = blockAt(S.doc, pth);
    if (!p?.pPr?.num) { toast('번호 매기기 목록 안에서 실행하세요'); return; }
    change((d) => { const ls = d.prop('lists'); const id = p.pPr.num.id; const lvl = p.pPr.num.lvl ?? 0; ls[id] = { ...ls[id], override: { ...(ls[id].override ?? {}), [lvl]: { start: num(v.n, 1) } } }; const b = d.block(pth); b.pPr = { ...b.pPr, num: { ...b.pPr.num, restart: true } }; return S.sel; });
  }),
});

// ───────────── 페이지 설정 ─────────────
register({
  pageSetupDialog: (tab = 'margins') => {
    const sec = sectionAtTop(S.doc, sel().f.p[0] ?? 0);
    const m = sec.margin;
    const mt = numIn(f2(pxToCm(m.t))); const mb = numIn(f2(pxToCm(m.b))); const ml = numIn(f2(pxToCm(m.l))); const mr = numIn(f2(pxToCm(m.r))); const mg = numIn(f2(pxToCm(m.gutter ?? 0)));
    const ori = sel2([['portrait', '세로'], ['landscape', '가로']], sec.page.orient ?? 'portrait');
    const paperKey = Object.entries(PAPERS).find(([, p]) => Math.abs(Math.min(p.w, p.h) - Math.min(sec.page.w, sec.page.h)) < 2 && Math.abs(Math.max(p.w, p.h) - Math.max(sec.page.w, sec.page.h)) < 2)?.[0] ?? 'custom';
    const paper = sel2([...Object.entries(PAPERS).map(([k, p]) => [k, `${p.label} (${p.size})`]), ['custom', '사용자 지정']], paperKey);
    const pw = numIn(f2(pxToCm(sec.page.w))); const ph = numIn(f2(pxToCm(sec.page.h)));
    paper.addEventListener('change', () => { const p = PAPERS[paper.value]; if (p) { pw.value = f2(pxToCm(ori.value === 'landscape' ? p.h : p.w)); ph.value = f2(pxToCm(ori.value === 'landscape' ? p.w : p.h)); } });
    const hd = numIn(f2(pxToCm(m.header ?? 48))); const fd = numIn(f2(pxToCm(m.footer ?? 48)));
    const tp = chk('첫 페이지를 다르게 지정', sec.titlePg); const eo = chk('홀수/짝수 페이지를 다르게 지정', S.doc.settings.evenOdd);
    const va = sel2([['top', '위쪽'], ['center', '가운데'], ['justify', '양쪽'], ['bottom', '아래쪽']], sec.vAlign ?? 'top');
    const t = tabs([
      ['여백', el('div', { class: 'dlg-grid' }, row('위쪽:', mt, 'cm'), row('아래쪽:', mb, 'cm'), row('왼쪽:', ml, 'cm'), row('오른쪽:', mr, 'cm'), row('제본용 여백:', mg, 'cm'), row('용지 방향:', ori))],
      ['용지', el('div', { class: 'dlg-grid' }, row('용지 크기:', paper), row('너비:', pw, 'cm'), row('높이:', ph, 'cm'))],
      ['레이아웃', el('div', { class: 'dlg-grid' }, el('div', { class: 'dlg-sec' }, '머리글/바닥글'), tp.node, eo.node, row('머리글 위치:', hd, 'cm'), row('바닥글 위치:', fd, 'cm'), row('세로 맞춤:', va))],
    ]);
    if (tab === 'paper') t.show(1);
    openDialog({ title: '페이지 설정', width: 440, body: t.node, buttons: [{ label: '확인', primary: true, action: () => {
      let w = cm(num(pw.value)); let h = cm(num(ph.value));
      if ((ori.value === 'landscape') !== (w > h)) [w, h] = [h, w];
      run('pageSetup', { page: { w, h, orient: ori.value }, margin: { t: cm(num(mt.value)), b: cm(num(mb.value)), l: cm(num(ml.value)), r: cm(num(mr.value)), gutter: cm(num(mg.value)), header: cm(num(hd.value)), footer: cm(num(fd.value)) }, titlePg: tp.i.checked, vAlign: va.value });
      if (!!S.doc.settings.evenOdd !== eo.i.checked) run('differentOddEven');
    } }, { label: '취소' }] });
  },
  columnsDialog: () => {
    const sec = sectionAtTop(S.doc, sel().f.p[0] ?? 0);
    formDialog('단', [{ name: 'n', label: '단 개수', type: 'number', value: sec.cols?.n ?? 1 }, { name: 'sp', label: '간격 (cm)', type: 'number', value: f2(pxToCm(sec.cols?.space ?? 48)) }, { name: 'sep', label: '경계선 삽입', type: 'checkbox', value: !!sec.cols?.sep }], (v) => run('columns', num(v.n, 1), cm(num(v.sp, 1.27)), v.sep));
  },
  pageNumberFormatDialog: () => {
    const sec = sectionAtTop(S.doc, sel().f.p[0] ?? 0);
    formDialog('페이지 번호 서식', [{ name: 'fmt', label: '번호 서식', type: 'select', value: sec.pgNum?.fmt ?? 'decimal', options: [['decimal', '1, 2, 3, …'], ['dash', '- 1 -, - 2 -, …'], ['lowerLetter', 'a, b, c, …'], ['upperLetter', 'A, B, C, …'], ['lowerRoman', 'i, ii, iii, …'], ['upperRoman', 'I, II, III, …'], ['ganada', '가, 나, 다, …']].map(([value, label]) => ({ value, label })) }, { name: 'start', label: '시작 번호 (비우면 이어서)', type: 'number', value: sec.pgNum?.start ?? '' }], (v) => run('pageNumberFormat', v.fmt, v.start));
  },
  watermarkDialog: () => formDialog('사용자 지정 워터마크', [{ name: 't', label: '텍스트', value: S.doc.settings.watermark?.text ?? '초안' }, { name: 'c', label: '색', type: 'color', value: S.doc.settings.watermark?.color ?? '#D9D9D9' }, { name: 'd', label: '대각선', type: 'checkbox', value: S.doc.settings.watermark?.diagonal !== false }], (v) => run('watermark', v.t, { color: v.c, diagonal: v.d })),
  pageBordersDialog: () => {
    const sec = sectionAtTop(S.doc, 0);
    const b = sec.pageBorder ?? { style: 'single', w: 1, color: '#000000', space: 24 };
    formDialog('페이지 테두리', [{ name: 'on', label: '테두리 사용', type: 'checkbox', value: !!sec.pageBorder }, { name: 'style', label: '스타일', type: 'select', value: b.style, options: [['single', '실선'], ['double', '이중'], ['dotted', '점선'], ['dashed', '파선']].map(([value, label]) => ({ value, label })) }, { name: 'w', label: '두께 (pt)', type: 'number', value: b.w }, { name: 'color', label: '색', type: 'color', value: b.color }, { name: 'space', label: '용지 가장자리에서 (pt)', type: 'number', value: b.space }], (v) => run('pageBorders', v.on ? { style: v.style, w: num(v.w, 1), color: v.color, space: num(v.space, 24) } : 'none'));
  },
  manuscriptDialog: () => formDialog('원고지 설정', [{ name: 'on', label: '원고지 사용', type: 'checkbox', value: !!S.doc.settings.manuscript }, { name: 'cols', label: '줄당 글자 수', type: 'number', value: S.doc.settings.manuscript?.cols ?? 20 }, { name: 'rows', label: '페이지당 줄 수', type: 'number', value: S.doc.settings.manuscript?.rows ?? 20 }], (v) => {
    change((d) => { d.doc.settings = { ...d.doc.settings, manuscript: v.on ? { cols: num(v.cols, 20), rows: num(v.rows, 20) } : undefined }; return S.sel; });
    if (v.on) {
      // 원고지: 글꼴 크기 · 줄 간격을 칸에 맞춤 (본문 폭 / 글자 수)
      const sec = sectionAtTop(S.doc, 0);
      const W = sec.page.w - sec.margin.l - sec.margin.r;
      const H = sec.page.h - sec.margin.t - sec.margin.b;
      const cell = W / num(v.cols, 20);
      change((d) => { d.doc.defaults = { ...d.doc.defaults, rPr: { ...d.doc.defaults.rPr, sz: f1((cell * 0.72 * 3) / 4), font: '바탕', fontEa: '바탕' }, pPr: { ...d.doc.defaults.pPr, align: 'left', spc: { before: 0, after: 0, line: H / num(v.rows, 20), rule: 'exact' } } }; return S.sel; });
      toast('원고지 칸에 맞춰 글자 크기와 줄 간격을 바꿨습니다');
    }
  }),
});

// ───────────── 찾기 · 바꾸기 ─────────────
register({
  replaceDialog: (tab = 'replace') => {
    const q = el('input', { value: S.findQuery ?? (posEq(sel().a, sel().f) ? '' : run('__selText')) });
    const r = el('input', { value: S.replaceText ?? '' });
    const mc = chk('대/소문자 구분', false); const ww = chk('단어 단위로', false); const rx = chk('정규식 사용', false); const wc = chk('와일드카드 사용 (* ?)', false);
    const opts = () => ({ matchCase: mc.i.checked, wholeWord: ww.i.checked, regex: rx.i.checked, wildcards: wc.i.checked });
    const status = el('div', { class: 'muted' });
    const go = el('input', { type: 'number', value: 1, min: 1 });
    const t = tabs([
      ['찾기', el('div', { class: 'dlg-grid' }, row('찾을 내용:', q))],
      ['바꾸기', el('div', { class: 'dlg-grid' }, row('찾을 내용:', q.cloneNode()), row('바꿀 내용:', r))],
      ['이동', el('div', { class: 'dlg-grid' }, row('페이지 번호:', go), el('button', { type: 'button', class: 'btn', onclick: () => run('goToPage', go.value) }, '이동'))],
    ]);
    // 두 탭의 찾을 내용 칸을 같게
    const qs = () => [...t.node.querySelectorAll('input:not([type=number]):not([type=checkbox])')];
    t.node.addEventListener('input', (e) => { if (e.target !== r && e.target.tagName === 'INPUT' && e.target.type !== 'number' && e.target.type !== 'checkbox') q.value = e.target.value; });
    const qv = () => (qs().find((x) => x.isConnected && x !== r)?.value ?? q.value);
    if (tab === 'find') t.show(0); else if (tab === 'goto') t.show(2); else t.show(1);
    const body = el('div', {}, t.node, el('div', { class: 'dlg-chks' }, mc.node, ww.node, rx.node, wc.node), status);
    openDialog({ title: '찾기 및 바꾸기', width: 480, body, modeless: true, buttons: [
      { label: '모두 바꾸기', action: () => { S.findQuery = qv(); S.replaceText = r.value; const n = run('replaceAll', qv(), r.value, opts()); status.textContent = `${n ?? 0}개 바꿈`; return false; } },
      { label: '바꾸기', action: () => { S.findQuery = qv(); S.replaceText = r.value; run('replace', qv(), r.value, opts()); return false; } },
      { label: '다음 찾기', primary: true, action: () => { S.findQuery = qv(); const n = run('find', qv(), opts()); status.textContent = `${n}개 찾음`; run('findNext', qv(), opts(), 1); return false; } },
      { label: '닫기' },
    ] });
  },
  selectSimilar: () => {
    const p = blockAt(S.doc, sel().f.p);
    const target = JSON.stringify({ style: p.style, rp: propsAt(p, sel().f.o) });
    const hits = [];
    walkParas(S.doc.body, (q, path) => { if (JSON.stringify({ style: q.style, rp: propsAt(q, 1) }) === target) hits.push(path); });
    if (hits.length) setSel({ p: hits[0], o: 0 }, { p: hits[hits.length - 1], o: paraLen(blockAt(S.doc, hits[hits.length - 1])) });
    toast(`비슷한 서식의 단락 ${hits.length}개`);
  },
  sortDialog: () => {
    formDialog('텍스트 정렬', [{ name: 'dir', label: '방식', type: 'select', value: 'asc', options: [{ value: 'asc', label: '오름차순 (가나다)' }, { value: 'desc', label: '내림차순' }] }, { name: 'num', label: '숫자로 정렬', type: 'checkbox', value: false }], (v) => {
      const paths = selParas().filter((pt) => pt.length === 1);
      if (paths.length < 2) { toast('정렬할 단락을 두 개 이상 선택하세요'); return; }
      change((d) => {
        const i0 = paths[0][0];
        const blocks = paths.map((pt) => d.doc.body[pt[0]]);
        const key = (b) => paraPlain(b);
        blocks.sort((a, b) => { const x = key(a); const y = key(b); const c = v.num ? parseFloat(x) - parseFloat(y) : x.localeCompare(y, 'ko'); return v.dir === 'desc' ? -c : c; });
        d.doc.body.splice(i0, blocks.length, ...blocks);
        return S.sel;
      });
    });
  },
});

// ───────────── 삽입 대화 상자 ─────────────
register({
  insertTableDialog: () => formDialog('표 삽입', [{ name: 'c', label: '열 개수', type: 'number', value: 5 }, { name: 'r', label: '행 개수', type: 'number', value: 2 }, { name: 'h', label: '첫 행을 머리글로', type: 'checkbox', value: true }], (v) => { run('insertTable', num(v.r, 2), num(v.c, 5)); if (v.h) run('repeatHeader'); }),
  insertLinkDialog: () => {
    const marks = [];
    walkParas(S.doc.body, (p) => { for (const r of p.runs) if (r.obj === 'mark' && r.kind === 'bm' && r.edge === 'start') marks.push(r.name); });
    const hs = headings(S.doc, 9);
    const text = posEq(sel().a, sel().f) ? '' : run('__selText');
    formDialog('하이퍼링크 삽입', [
      { name: 'text', label: '표시할 텍스트', value: text },
      { name: 'url', label: '주소 (웹 주소 · 메일 · #책갈피)', type: 'combo', value: '', options: [...marks.map((m) => ({ value: `#${m}`, label: `책갈피: ${m}` })), ...hs.map((h) => ({ value: `#${h.p.id}`, label: `제목: ${h.text}` }))] },
    ], (v) => { if (v.url) run('insertLink', v.url.includes('@') && !v.url.includes(':') ? `mailto:${v.url}` : v.url, v.text); });
  },
  bookmarkDialog: () => {
    const marks = [];
    walkParas(S.doc.body, (p, path) => { let o = 0; for (const r of p.runs) { if (r.obj === 'mark' && r.kind === 'bm' && r.edge === 'start') marks.push({ name: r.name, path, o }); o += runLen(r); } });
    const name = el('input', { placeholder: '책갈피 이름 (공백 없이)' });
    const list = el('div', { class: 'dlg-list' }, marks.map((m) => el('div', { class: 'dlg-item' }, m.name, el('button', { type: 'button', onclick: () => setSel({ p: m.path, o: m.o }) }, '이동'), el('button', { type: 'button', onclick: () => { change((d) => { const b = d.block(m.path); b.runs = b.runs.filter((r) => !(r.obj === 'mark' && r.kind === 'bm' && r.name === m.name)); return S.sel; }); } }, '삭제'))));
    openDialog({ title: '책갈피', width: 400, body: el('div', { class: 'dlg-grid' }, row('이름:', name), list), buttons: [{ label: '추가', primary: true, action: () => { const n = name.value.trim().replace(/\s+/g, '_'); if (!n) return false; run('insertBookmark', n); return true; } }, { label: '닫기' }] });
  },
  crossRefDialog: () => {
    const hs = headings(S.doc, 9);
    const caps = [];
    walkParas(S.doc.body, (p) => { if (p.style === 'Caption') caps.push({ id: p.id, text: paraPlain(p) }); });
    const items = [...hs.map((h) => ({ id: h.p.id, label: `제목: ${h.text}`, text: h.text })), ...caps.map((c) => ({ id: c.id, label: `캡션: ${c.text}`, text: c.text }))];
    if (!items.length) { toast('참조할 제목이나 캡션이 없습니다'); return; }
    formDialog('상호 참조', [{ name: 'it', label: '참조 대상', type: 'select', value: items[0].id, options: items.map((x) => ({ value: x.id, label: x.label })) }, { name: 'kind', label: '참조할 내용', type: 'select', value: 'text', options: [{ value: 'text', label: '텍스트' }, { value: 'page', label: '페이지 번호' }] }], (v) => {
      const it = items.find((x) => x.id === v.it);
      if (v.kind === 'page') run('__insertRun', { obj: 'field', code: `PAGEREF ${it.id} \\h`, result: String(pageOfBlock(it.id) ?? ''), link: { anchor: it.id } });
      else run('__insertRun', { text: it.text, link: { anchor: it.id } });
    });
  },
  dateTimeDialog: () => {
    const d = new Date();
    const pad = (n) => String(n).padStart(2, '0');
    const W = ['일', '월', '화', '수', '목', '금', '토'];
    const fmts = [`${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}`, `${d.getFullYear()}. ${d.getMonth() + 1}. ${d.getDate()}.`, `${d.getFullYear()}년 ${d.getMonth() + 1}월 ${d.getDate()}일`, `${d.getFullYear()}년 ${d.getMonth() + 1}월 ${d.getDate()}일 ${W[d.getDay()]}요일`, `${d.getFullYear()}. ${pad(d.getMonth() + 1)}. ${pad(d.getDate())}.(${W[d.getDay()]})`, `${d.getMonth() + 1}/${d.getDate()}/${d.getFullYear()}`, d.toLocaleDateString('en-US', { year: 'numeric', month: 'long', day: 'numeric' }), `${pad(d.getHours())}:${pad(d.getMinutes())}`, `${d.getHours() < 12 ? '오전' : '오후'} ${d.getHours() % 12 || 12}:${pad(d.getMinutes())}`, `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())} ${pad(d.getHours())}:${pad(d.getMinutes())}`];
    formDialog('날짜 및 시간', [{ name: 'f', label: '사용 가능한 형식', type: 'select', value: fmts[2], options: fmts.map((x) => ({ value: x, label: x })) }, { name: 'live', label: '자동으로 업데이트', type: 'checkbox', value: false }], (v) => { if (v.live) run('__insertRun', { obj: 'field', code: 'DATE', result: v.f }); else run('insertText', v.f); });
  },
  fieldDialog: () => formDialog('필드', [{ name: 'code', label: '필드 코드', type: 'select', value: 'PAGE', options: ['PAGE', 'NUMPAGES', 'DATE', 'TIME', 'AUTHOR', 'TITLE', 'NUMWORDS', 'FILENAME'].map((x) => ({ value: x, label: x })) }], (v) => run('insertField', v.code)),
  symbolDialog: () => {
    const sets = { '일반 문장 부호': '…‥·•※★☆◎○●◇◆□■△▲▽▼→←↑↓↔〓≒∽∵∴∀∃∂∇∈∋⊂⊃∪∩∧∨¬⇒⇔', '원 문자': '①②③④⑤⑥⑦⑧⑨⑩⑪⑫⑬⑭⑮㉠㉡㉢㉣㉤㉥㉦㉧㉮㉯㉰㉱㉲㉳㉴㉵ⓐⓑⓒⓓⓔ', '괄호 문자': '⑴⑵⑶⑷⑸⒜⒝⒞⒟㈀㈁㈂㈃㈎㈏㈐㈑', '단위': '㎜㎝㎞㎎㎏㎖㎗㎘㏄㎡㎥℃℉°′″‰₩€£¥$¢', '그리스 문자': 'ΑΒΓΔΕΖΗΘΙΚΛΜΝΞΟΠΡΣΤΥΦΧΨΩαβγδεζηθικλμνξοπρστυφχψω', '수학': '±×÷≠≤≥∞√∑∏∫∮≈≡∝∠⊥⌒', '도형 · 기타': '♠♣♥♦♤♧♡♢☎☏☜☞♨♪♬✓✔✕✖❖☑☒☐' };
    const box = el('div', {});
    for (const [name, chars] of Object.entries(sets)) {
      box.append(el('div', { class: 'dlg-sec' }, name));
      const g = el('div', { class: 'sym-grid' });
      for (const ch of chars) g.append(el('button', { type: 'button', title: `U+${ch.codePointAt(0).toString(16).toUpperCase()}`, onclick: () => run('insertText', ch) }, ch));
      box.append(g);
    }
    openDialog({ title: '기호', width: 520, body: box, modeless: true, buttons: [{ label: '닫기' }] });
  },
  equationDialog: () => {
    const tex = el('textarea', { rows: 3, style: { width: '100%', fontFamily: 'Consolas, monospace' } }, 'x=\\frac{-b\\pm\\sqrt{b^2-4ac}}{2a}');
    const prev = el('div', { class: 'dlg-preview', style: { minHeight: '60px', fontSize: '18px' } });
    const upd = () => { try { prev.innerHTML = toMathML(parseLatex(tex.value)); } catch (e) { prev.textContent = `수식 오류: ${e.message}`; } };
    tex.addEventListener('input', upd); upd();
    openDialog({ title: '수식 (LaTeX)', width: 520, body: el('div', {}, el('div', { class: 'muted' }, 'LaTeX 로 입력하세요. 예: \\frac{a}{b}, x^2, \\sqrt{x}, \\sum_{i=1}^n, \\int_0^1'), tex, el('div', { class: 'dlg-sec' }, '미리 보기'), prev),
      buttons: [{ label: '삽입', primary: true, action: () => run('insertEquation', tex.value) }, { label: '취소' }] });
  },
  insertEquation: (latex) => {
    let html;
    try { html = toMathML(parseLatex(latex)); } catch { html = latex; }
    run('__insertRun', { obj: 'eq', latex, html });
  },
  signatureLine: () => formDialog('서명란', [{ name: 'name', label: '서명자', value: '' }, { name: 'title', label: '직함', value: '' }], (v) => run('insertBlocksAt', [para('', { pPr: { spc: { before: 36 } } }), { ...para(''), runs: [{ text: '×  ' }, { text: '                                        ', u: true }] }, para(v.name || '[서명자]'), para(v.title || ''), para('')])),
  captionDialog: () => formDialog('캡션', [{ name: 'label', label: '레이블', type: 'select', value: '그림', options: ['그림', '표', '수식', 'Figure', 'Table'].map((x) => ({ value: x, label: x })) }, { name: 'text', label: '캡션 내용', value: '' }], (v) => run('insertCaption', v.label, v.text)),
  watermark2: () => {},
});

// ───────────── 표 대화 상자 ─────────────
register({
  tablePropsDialog: () => {
    const p = sel().f.p;
    if (p.length < 4) { toast('표 안을 클릭하세요'); return; }
    const t = blockAt(S.doc, p.slice(0, -3));
    const w = t.grid.reduce((a, b) => a + b, 0);
    formDialog('표 속성', [
      { name: 'w', label: '표 너비 (cm)', type: 'number', value: f2(pxToCm(w)) },
      { name: 'align', label: '맞춤', type: 'select', value: t.tblPr.align ?? 'left', options: [{ value: 'left', label: '왼쪽' }, { value: 'center', label: '가운데' }, { value: 'right', label: '오른쪽' }] },
      { name: 'ind', label: '왼쪽 들여쓰기 (cm)', type: 'number', value: f2(pxToCm(t.tblPr.ind ?? 0)) },
      { name: 'rh', label: '행 높이 (cm, 0 = 자동)', type: 'number', value: f2(pxToCm(t.rows[p[p.length - 3]].h ?? 0)) },
    ], (v) => {
      const nw = cm(num(v.w, pxToCm(w)));
      run('__tableProps', { scale: nw / w, align: v.align, ind: cm(num(v.ind)) });
      if (num(v.rh) > 0) run('setRowHeight', cm(num(v.rh))); else run('setRowHeight', 0);
    });
  },
  splitCellDialog: () => formDialog('셀 분할', [{ name: 'c', label: '열 개수', type: 'number', value: 2 }, { name: 'r', label: '행 개수', type: 'number', value: 1 }], (v) => run('splitCell', num(v.c, 2), num(v.r, 1))),
  cellMarginsDialog: () => formDialog('셀 여백 (cm)', [{ name: 't', label: '위쪽', type: 'number', value: 0 }, { name: 'b', label: '아래쪽', type: 'number', value: 0 }, { name: 'l', label: '왼쪽', type: 'number', value: 0.19 }, { name: 'r', label: '오른쪽', type: 'number', value: 0.19 }], (v) => run('__tableProps', { cellMar: { t: cm(num(v.t)), b: cm(num(v.b)), l: cm(num(v.l)), r: cm(num(v.r)) } })),
  sortTableDialog: () => {
    const p = sel().f.p;
    if (p.length < 4) { toast('표 안을 클릭하세요'); return; }
    const t = blockAt(S.doc, p.slice(0, -3));
    const heads = t.rows[0].cells.map((c, i) => ({ value: String(i), label: `${i + 1}열${c.blocks[0]?.type === 'p' && paraPlain(c.blocks[0]) ? ` (${paraPlain(c.blocks[0])})` : ''}` }));
    formDialog('정렬', [{ name: 'col', label: '정렬 기준', type: 'select', value: '0', options: heads }, { name: 'dir', label: '방향', type: 'select', value: 'asc', options: [{ value: 'asc', label: '오름차순' }, { value: 'desc', label: '내림차순' }] }, { name: 'head', label: '머리글 행 있음', type: 'checkbox', value: true }], (v) => run('sortTable', Number(v.col), v.dir === 'desc', v.head));
  },
  tableFormulaDialog: () => formDialog('수식', [{ name: 'f', label: '수식', type: 'combo', value: '=SUM(ABOVE)', options: ['=SUM(ABOVE)', '=SUM(LEFT)', '=AVERAGE(ABOVE)', '=MAX(ABOVE)', '=MIN(ABOVE)', '=COUNT(ABOVE)', '=PRODUCT(LEFT)'].map((x) => ({ value: x, label: x })) }], (v) => run('tableFormula', v.f)),
  splitTable: () => {
    const p = sel().f.p;
    if (p.length < 4) return;
    const tpath = p.slice(0, -3); const r = p[p.length - 3];
    if (r === 0) return;
    change((d) => { const lst = d.list(tpath); const i = tpath[tpath.length - 1]; const t = lst[i]; const a = { ...t, rows: t.rows.slice(0, r) }; const b = { ...t, id: uid('t'), rows: t.rows.slice(r) }; lst.splice(i, 1, a, para(''), b); return { p: [...tpath.slice(0, -1), i + 1], o: 0 }; });
  },
  __tableProps: (o) => {
    const p = sel().f.p;
    if (p.length < 4) return;
    const tpath = p.slice(0, -3);
    change((d) => { const t = d.block(tpath); if (o.scale && Math.abs(o.scale - 1) > 0.001) t.grid = t.grid.map((g) => g * o.scale); t.tblPr = { ...t.tblPr, ...(o.align ? { align: o.align } : {}), ...(o.ind != null ? { ind: o.ind } : {}), ...(o.cellMar ? { cellMar: o.cellMar } : {}) }; return S.sel; });
  },
  cellTextDir: () => run('__cellsEditDir'),
});
for (const [k, v, h] of [['TL', 'top', 'left'], ['TC', 'top', 'center'], ['TR', 'top', 'right'], ['ML', 'center', 'left'], ['MC', 'center', 'center'], ['MR', 'center', 'right'], ['BL', 'bottom', 'left'], ['BC', 'bottom', 'center'], ['BR', 'bottom', 'right']]) register({ [`cellAlign${k}`]: () => run('cellAlign', v, h) });

// ───────────── 그림 ─────────────
register({
  imageWrap: (wrap, side = 'left') => {
    if (!S.objSel) { toast('그림을 먼저 선택하세요'); return; }
    run('imageProps', { float: wrap === 'inline' ? undefined : { wrap, side, dist: { t: 0, b: 6, l: 12, r: 12 } } });
  },
  rotateImage: (deg) => { if (!S.objSel) return; const r = run('__objRun'); run('imageProps', { rot: deg === 0 ? undefined : (((r?.rot ?? 0) + deg) % 360 + 360) % 360 || undefined }); },
  bringForward: () => run('imageProps', { z: 1 }),
  sendBackward: () => run('imageWrap', 'behind'),
  altTextDialog: () => { const r = run('__objRun'); if (!r) { toast('그림을 먼저 선택하세요'); return; } formDialog('대체 텍스트', [{ name: 'alt', label: '설명 (시각 장애인을 위한 설명)', type: 'textarea', value: r.alt ?? '' }, { name: 'deco', label: '장식용으로 표시', type: 'checkbox', value: !!r.decorative }], (v) => run('imageProps', { alt: v.alt, decorative: v.deco || undefined })); },
  cropPicture: () => { const r = run('__objRun'); if (!r) { toast('그림을 먼저 선택하세요'); return; } const c = r.crop ?? { l: 0, t: 0, r: 0, b: 0 }; formDialog('자르기 (%)', ['l', 't', 'r', 'b'].map((k) => ({ name: k, label: { l: '왼쪽', t: '위쪽', r: '오른쪽', b: '아래쪽' }[k], type: 'number', value: Math.round(c[k] * 100) })), (v) => { const nc = { l: num(v.l) / 100, t: num(v.t) / 100, r: num(v.r) / 100, b: num(v.b) / 100 }; const old = r.crop ?? { l: 0, t: 0, r: 0, b: 0 }; const fw = r.w / (1 - old.l - old.r); const fh = r.h / (1 - old.t - old.b); run('imageProps', { crop: nc.l || nc.t || nc.r || nc.b ? nc : undefined, w: fw * (1 - nc.l - nc.r), h: fh * (1 - nc.t - nc.b) }); }); },
  resetPicture: () => run('imageProps', { crop: undefined, bright: undefined, contrast: undefined, gray: undefined, border: undefined, shadow: undefined, rot: undefined }),
});

// ───────────── 기타 ─────────────
register({
  zoomDialog: () => formDialog('확대/축소', [{ name: 'z', label: '배율 (%)', type: 'combo', value: Math.round(S.zoom * 100), options: [200, 150, 120, 100, 90, 75, 50].map((x) => ({ value: x, label: `${x}%` })) }], (v) => run('zoomTo', num(v.z, 100))),
  wordCountDialog: () => {
    const st = docStats(S.doc);
    const pages = S.layout?.pages.length ?? 1;
    const sel0 = posEq(sel().a, sel().f) ? null : run('__selText');
    const rows = [['페이지', pages], ['단어', st.words], ['문자 (공백 제외)', st.charsNoSpace], ['문자 (공백 포함)', st.chars], ['단락', st.paras], ['원고지 (200자 기준)', `${Math.ceil(st.chars / 200)}매`]];
    if (sel0) rows.unshift(['선택 영역 단어', (sel0.match(/\S+/g) ?? []).length], ['선택 영역 문자', sel0.length]);
    openDialog({ title: '단어 개수', width: 340, body: el('table', { class: 'stat-tbl' }, rows.map(([k, v]) => el('tr', {}, el('td', {}, k), el('td', { style: { textAlign: 'right' } }, String(typeof v === 'number' ? v.toLocaleString('ko-KR') : v))))), buttons: [{ label: '닫기', primary: true }] });
  },
  __insertRun: (r) => { const { f, a } = sel(); const at = posEq(a, f) ? f : orderPos(a, f)[0]; change((d) => insertRuns(d, at, [{ ...propsAt(blockAt(d.doc, at.p), at.o), ...r }])); },
  __selText: () => run('selectedTextImpl'),
  __setSel: (a, f) => setSel(a, f),
  notesDialog: () => formDialog('각주 및 미주', [{ name: 'fmt', label: '번호 서식', type: 'select', value: '1', options: [{ value: '1', label: '1, 2, 3' }, { value: 'a', label: 'a, b, c' }, { value: 'i', label: 'i, ii, iii' }, { value: '*', label: '*, †, ‡' }] }], () => toast('각주 번호 서식을 저장했습니다')),
  trackOptionsDialog: () => formDialog('변경 내용 추적 옵션', [{ name: 'name', label: '사용자 이름 (변경 내용 · 메모 작성자)', value: S.opts.userName ?? '' }], (v) => { S.opts.userName = v.name; try { localStorage.setItem('wiword:user', v.name); } catch { /* 없음 */ } }),
});
void cellSelection; void selProp; void bodyWidthAt; void sections; void styleChain; void firstPos; void splitRuns; void normalizeRuns; void clone; void MARGINS; void closeMenus; void emit; void V;
