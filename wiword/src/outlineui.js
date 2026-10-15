// 개요로 새 문서 만들기 창 — 마크다운 개요를 쓰거나 다른 문서의 제목들을 가져와, 표지 · 차례 · 번호가 있는 문서로
import { register, run, setDoc, S } from './state.js';
import { OUTLINE_KINDS, buildFromOutline, parseOutline, outlineFromDoc } from './outline.js';
import { pickFile, loadFile } from './fileio.js';
import { OPEN_ACCEPT } from './formats.js';
import { el, openDialog, toast, alertDialog } from './ui.js';

const SAMPLE = `# 2026년 하반기 사업 계획
## 개요
- 추진 배경과 목적
## 현황 분석
### 시장 동향
### 경쟁사 비교
| 구분 | 우리 | 경쟁사 |
|---|---|---|
| 점유율 | 12% | 30% |
## 추진 계획
### 일정
### 예산
## 기대 효과
`;
const today = () => { const d = new Date(); return `${d.getFullYear()}. ${d.getMonth() + 1}. ${d.getDate()}.`; };

export function outlineDialog(initial = null) {
  const ta = el('textarea', { class: 'ol-text', spellcheck: 'false', placeholder: '# 문서 제목\n## 장 제목\n### 절 제목\n- 항목 (들여쓰면 아래 단계)\n본문 글, 표(| 가 | 나 |)도 그대로 씁니다' });
  ta.value = initial ?? SAMPLE;
  const kindSel = el('select', {}, OUTLINE_KINDS.map((k) => el('option', { value: k.id }, `${k.name} — ${k.desc}`)));
  const f = (label, input) => el('label', { class: 'ol-field' }, el('span', {}, label), input);
  const title = el('input', { type: 'text', placeholder: '비우면 개요의 맨 위 # 제목' });
  const subtitle = el('input', { type: 'text' });
  const author = el('input', { type: 'text', value: S.opts?.userName ?? '' });
  const date = el('input', { type: 'text', value: today() });
  const chk = (label) => { const i = el('input', { type: 'checkbox' }); return [i, el('label', { class: 'ol-chk' }, i, ` ${label}`)]; };
  const [cover, coverL] = chk('표지');
  const [toc, tocL] = chk('차례');
  const [brk, brkL] = chk('장마다 새 쪽');
  const [place, placeL] = chk('빈 절에 안내 문구');
  const numSel = el('select', {}, [['decimal', '1. / 1.1 / 1.1.1'], ['plan', 'Ⅰ. / 1. / 가.'], ['gongmun', '1. / 가. / 1) / 가)'], ['paper', '1 / 1.1 / 1.1.1'], ['none', '번호 없음']].map(([v, l]) => el('option', { value: v }, l)));
  const preview = el('div', { class: 'ol-preview' });
  const applyKind = () => {
    const k = OUTLINE_KINDS.find((x) => x.id === kindSel.value);
    cover.checked = k.cover; toc.checked = k.toc; brk.checked = k.chapterBreak; numSel.value = k.number; place.checked = true;
    refresh();
  };
  const refresh = () => {
    const items = parseOutline(ta.value).filter((x) => x.kind === 'h');
    if (!items.length) { preview.replaceChildren(el('div', { class: 'muted' }, '제목(# · ## · ###)이 없으면 글 그대로 문서가 됩니다.')); return; }
    const doc = buildFromOutline(ta.value, opts());
    const hs = doc.body.filter((b) => b.type === 'p' && /^Heading[1-6]$/.test(b.style ?? ''));
    preview.replaceChildren(el('div', { class: 'ol-ptitle' }, doc.props.title),
      ...hs.slice(0, 60).map((h) => el('div', { class: `ol-h ol-h${h.style.slice(-1)}` }, h.runs.map((r) => r.text ?? '').join(''))),
      el('div', { class: 'muted' }, `제목 ${hs.length}개${cover.checked ? ' · 표지' : ''}${toc.checked ? ' · 차례' : ''}`));
  };
  const opts = () => ({ kind: kindSel.value, title: title.value.trim() || undefined, subtitle: subtitle.value.trim() || undefined, author: author.value.trim() || undefined, date: date.value.trim() || undefined, cover: cover.checked, toc: toc.checked, chapterBreak: brk.checked, placeholders: place.checked, number: numSel.value });
  for (const x of [ta, title, subtitle, author, date]) x.addEventListener('input', () => { clearTimeout(x._t); x._t = setTimeout(refresh, 250); });
  for (const x of [cover, toc, brk, place, numSel]) x.addEventListener('change', refresh);
  kindSel.addEventListener('change', applyKind);
  const fromFile = el('button', { class: 'btn', onclick: async () => {
    const file = await pickFile(OPEN_ACCEPT);
    if (!file) return;
    try {
      const r = await loadFile(file);
      ta.value = outlineFromDoc(r.doc);
      if (!title.value && r.doc.props?.title) title.value = r.doc.props.title;
      refresh();
      toast(`'${file.name}' 의 제목들로 개요를 만들었습니다. 고쳐서 [만들기]를 누르세요.`);
    } catch (e) { alertDialog('개요 가져오기', e.message); }
  } }, '파일에서 개요 가져오기…');
  const fromCurrent = el('button', { class: 'btn', onclick: () => { ta.value = outlineFromDoc(S.doc); refresh(); } }, '지금 문서의 개요');
  const body = el('div', { class: 'ol-wrap' },
    el('div', { class: 'ol-left' }, el('div', { class: 'ol-bar' }, fromFile, fromCurrent), ta),
    el('div', { class: 'ol-right' },
      f('종류', kindSel), f('제목', title), f('부제', subtitle), f('만든 이', author), f('날짜', date), f('번호', numSel),
      el('div', { class: 'ol-chks' }, coverL, tocL, brkL, placeL),
      el('div', { class: 'ol-plabel' }, '미리 보기'), preview));
  const dlg = openDialog({
    title: '개요로 새 문서 만들기', width: Math.min(1000, window.innerWidth - 24), body,
    buttons: [{ label: '만들기', primary: true, action: () => {
      const doc = buildFromOutline(ta.value, opts());
      setDoc(doc, doc.props.title || '문서1');
      S.fileHandle = null; S.fileKind = null; S.docId = null;
      setTimeout(() => { try { run('updateToc'); } catch { /* 차례 없음 */ } }, 300);
      toast('새 문서를 만들었습니다. 차례의 쪽 번호는 [참조 › 목차 업데이트]로 다시 맞출 수 있습니다.');
    } }, { label: '취소' }],
  });
  applyKind();
  setTimeout(() => ta.focus(), 30);
  return dlg;
}

register({ outlineDoc: (text) => outlineDialog(typeof text === 'string' ? text : null) }, { outlineDoc: { label: '개요로 새 문서 만들기', args: ['markdown?'] } });
