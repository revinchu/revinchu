// 시작 문서 · 서식 파일 (DOM 없음)
import { newDoc, para, table, uid, bulletList, numberList, multiList, cm, newSection, PAPERS, MARGINS } from './model.js';

const R = (text, o = {}) => ({ text, ...o });
const P = (runs, style = 'Normal', pPr = {}) => ({ ...para(typeof runs === 'string' ? runs : '', { style, pPr }), runs: typeof runs === 'string' ? (runs ? [R(runs)] : []) : runs });
const pageNum = () => ({ obj: 'field', code: 'PAGE', result: '1' });
const numPages = () => ({ obj: 'field', code: 'NUMPAGES', result: '1' });

/** 둘러보기 문서 (처음 실행) */
export function sampleDoc() {
  const d = newDoc();
  d.props.title = 'WIWORD 둘러보기';
  d.lists = { 1: bulletList('●'), 2: numberList(), 3: multiList('korean') };
  const fn1 = uid('n');
  d.notes.foot[fn1] = [P('HWP 5.0 · HWPX(OWPML) · HWPML 형식을 읽습니다. 저장은 HWPX · DOCX · PDF · 마크다운으로 할 수 있습니다.', 'FootnoteText')];
  const t = table(4, 4, { style: 'GridTable4-Accent1', width: cm(14.6) });
  const cells = [['기능', 'Word', 'WIWORD', '비고'], ['문서 형식', 'DOCX', 'DOCX · HWP · HWPX · PDF · MD', '한글 문서 열기'], ['AI 연동', 'Copilot', 'MCP (Claude 등)', '도구 22개'], ['설치', '필요', '필요 없음 (브라우저)', '단일 HTML']];
  t.rows.forEach((row, r) => row.cells.forEach((c, k) => { c.blocks = [P(cells[r][k], 'Normal', { spc: { after: 0, line: 1, rule: 'auto' } })]; }));
  t.rows[0].header = true;
  d.body = [
    P('WIWORD 둘러보기', 'Title'),
    P([R('Word 처럼 쓰고, 한글(HWP) 문서와 PDF 도 여는 브라우저 문서 편집기', { color: '#595959' })], 'Subtitle'),
    P('시작하기', 'Heading1'),
    P([R('이 문서에서 바로 글을 고쳐 보세요. 위쪽 '), R('리본', { b: true }), R('에서 글꼴 · 단락 · 스타일을 바꾸고, '), R('Ctrl+Z', { font: 'Consolas' }), R(' 로 되돌릴 수 있습니다. 쪽은 실제 용지(A4)처럼 나뉘며, 머리글 · 바닥글 · 각주'), { obj: 'note', id: fn1, kind: 'foot' }, R(' 가 쪽마다 그려집니다.')]),
    P('이런 것을 할 수 있습니다', 'Heading2'),
    P('Word 문서(.docx), 한글 문서(.hwp · .hwpx), PDF, 마크다운(.md), HTML, 텍스트 열기 · 편집 · 저장', 'ListParagraph', { num: { id: '1', lvl: 0 } }),
    P('스타일 · 목차 · 머리글/바닥글 · 각주/미주 · 쪽 번호 · 다단 · 구역', 'ListParagraph', { num: { id: '1', lvl: 0 } }),
    P('표 (병합 · 스타일 · 머리글 행 반복) · 그림 · 하이퍼링크 · 책갈피', 'ListParagraph', { num: { id: '1', lvl: 0 } }),
    P('변경 내용 추적 · 메모 · 맞춤법 · 단어 수 · 찾기/바꾸기(정규식)', 'ListParagraph', { num: { id: '1', lvl: 0 } }),
    P('AI 연동: MCP 서버로 Claude 같은 AI 가 문서를 읽고 고침', 'ListParagraph', { num: { id: '1', lvl: 0 } }),
    P('비교', 'Heading2'),
    t,
    P([R('표 1. Word 와 WIWORD', { i: true })], 'Caption'),
    P('한국 공문서 항목 기호', 'Heading2'),
    P('개요', 'ListParagraph', { num: { id: '3', lvl: 0 } }),
    P('추진 배경', 'ListParagraph', { num: { id: '3', lvl: 1 } }),
    P('세부 내용', 'ListParagraph', { num: { id: '3', lvl: 2 } }),
    P('세부 내용', 'ListParagraph', { num: { id: '3', lvl: 2 } }),
    P('기대 효과', 'ListParagraph', { num: { id: '3', lvl: 1 } }),
    P('단축키', 'Heading1'),
    P([R('Ctrl+B', { b: true }), { obj: 'tab' }, R('굵게'), { obj: 'br' }, R('Ctrl+Alt+1', { b: true }), { obj: 'tab' }, R('제목 1'), { obj: 'br' }, R('Ctrl+Enter', { b: true }), { obj: 'tab' }, R('페이지 나누기'), { obj: 'br' }, R('Alt', { b: true }), { obj: 'tab' }, R('리본 키 팁')], 'Normal', { tabs: [{ pos: cm(4), align: 'left', leader: 'dot' }] }),
    P([R('“글쓰기는 생각을 정리하는 가장 좋은 방법이다.”')], 'Quote'),
    P(''),
  ];
  d.section = newSection({
    headers: { default: [P([R('WIWORD 둘러보기'), { obj: 'tab' }, { obj: 'tab' }, R('위워드', { color: '#2B579A' })], 'Header')] },
    footers: { default: [P([{ obj: 'tab' }, R('- '), pageNum(), R(' / '), numPages(), R(' -')], 'Footer')] },
  });
  return d;
}

// ───────────── 서식 파일 ─────────────
export const TEMPLATES = [
  { id: 'blank', name: '새 문서', desc: '빈 문서', build: () => newDoc() },
  { id: 'tour', name: 'WIWORD 둘러보기', desc: '기능 소개', build: sampleDoc },
  { id: 'report', name: '보고서', desc: '표지 · 목차 · 제목', build: reportDoc },
  { id: 'gongmun', name: '공문서 (기안문)', desc: '행정업무운영 편람 서식', build: gongmunDoc },
  { id: 'letter', name: '편지', desc: '보내는 사람 · 받는 사람', build: letterDoc },
  { id: 'resume', name: '이력서', desc: '표로 정리한 이력서', build: resumeDoc },
  { id: 'minutes', name: '회의록', desc: '일시 · 참석자 · 안건', build: minutesDoc },
  { id: 'press', name: '보도자료', desc: '제목 · 부제 · 본문', build: pressDoc },
  { id: 'essay', name: '논문 · 과제', desc: '각주 · 참고 문헌', build: essayDoc },
  { id: 'hwp', name: '한글 기본 문서', desc: '함초롬바탕 10pt · 줄 간격 160%', build: hwpDoc },
];

function reportDoc() {
  const d = newDoc();
  d.props.title = '보고서';
  d.body = [
    P('', 'Normal', { spc: { before: cm(6) } }),
    P('[보고서 제목]', 'Title', { align: 'center' }),
    P('[부제목]', 'Subtitle', { align: 'center' }),
    P([R('[작성자] · [날짜]')], 'Normal', { align: 'center', spc: { before: cm(8) } }),
    P([{ obj: 'br', kind: 'page' }]),
    P('목차', 'TOCHeading'),
    P([R('[참조 › 목차 업데이트를 누르면 목차가 만들어집니다]', { color: '#7F7F7F' })]),
    P([{ obj: 'br', kind: 'page' }]),
    P('1. 개요', 'Heading1'), P('보고서의 목적과 범위를 적습니다.'),
    P('2. 현황', 'Heading1'), P('2.1 배경', 'Heading2'), P('현재 상황을 설명합니다.'), P('2.2 문제점', 'Heading2'), P('개선이 필요한 점을 적습니다.'),
    P('3. 개선 방안', 'Heading1'), P('방안과 일정, 예산을 적습니다.'),
    P('4. 결론', 'Heading1'), P('요약과 건의 사항을 적습니다.'),
  ];
  d.section = newSection({ titlePg: true, footers: { default: [P([pageNum()], 'Footer', { align: 'center' })] } });
  return d;
}
function gongmunDoc() {
  const d = newDoc();
  d.props.title = '기안문';
  d.themeFonts = { major: { latin: '함초롬바탕', ea: '함초롬바탕' }, minor: { latin: '함초롬바탕', ea: '함초롬바탕' } };
  d.defaults = { rPr: { font: '+mn', fontEa: '+mn', sz: 13, lang: 'ko-KR' }, pPr: { align: 'justify', spc: { before: 0, after: 0, line: 1.6, rule: 'pct' }, widow: true } };
  d.lists = { 1: multiList('korean') };
  const L = (t, lvl) => P(t, 'Normal', { num: { id: '1', lvl } });
  d.body = [
    P([R('행정기관명', { sz: 18, b: true })], 'Normal', { align: 'center', spc: { after: 13 } }),
    P([R('수신  '), R('수신자 참조')]),
    P([R('(경유)')]),
    P([R('제목  '), R('○○○ 추진 계획 알림', { b: true })], 'Normal', { border: { bottom: { style: 'single', w: 1.5, color: '#000000', space: 2 } } }),
    L('관련: ○○과-1234(2026. 1. 1.)', 0),
    L('위 호와 관련하여 ○○○ 추진 계획을 다음과 같이 알려드립니다.', 0),
    L('추진 일정: 2026. 3. 2.(월) ~ 3. 31.(화)', 1),
    L('추진 내용', 1),
    L('세부 사항 1', 2),
    L('세부 사항 2', 2),
    P(''),
    P([R('붙임  '), R('추진 계획서 1부.  끝.')]),
    P(''),
    P([R('행 정 기 관 장', { sz: 20, b: true })], 'Normal', { align: 'center', spc: { before: 26 } }),
  ];
  d.section = newSection({ margin: { t: cm(3), b: cm(1.5), l: cm(2), r: cm(2), header: cm(1), footer: cm(1), gutter: 0 } });
  return d;
}
function letterDoc() {
  const d = newDoc();
  d.body = [
    P('[보내는 사람 이름]', 'Normal', { spc: { after: 0 } }), P('[주소]', 'Normal', { spc: { after: 0 } }), P('[전화 · 전자 메일]'),
    P([R(new Date().toLocaleDateString('ko-KR', { year: 'numeric', month: 'long', day: 'numeric' }))], 'Normal', { spc: { before: 24 } }),
    P('[받는 사람 이름]', 'Normal', { spc: { before: 24, after: 0 } }), P('[주소]'),
    P('[받는 사람] 님께,', 'Normal', { spc: { before: 24 } }),
    P('여기에 편지 내용을 적습니다. 첫 문단에서는 편지를 쓰는 이유를 밝히고, 다음 문단에서 자세한 내용을 적습니다.'),
    P('감사합니다.'),
    P('[이름] 드림', 'Normal', { spc: { before: 36 } }),
  ];
  return d;
}
function resumeDoc() {
  const d = newDoc();
  const t = table(6, 2, { style: 'TableGrid', width: cm(16) });
  t.grid = [cm(4), cm(12)];
  const rows = [['이름', ''], ['연락처', ''], ['학력', '2016 ~ 2020  ○○대학교 ○○학과 졸업'], ['경력', '2020 ~ 현재  ○○주식회사 ○○팀'], ['자격', '정보처리기사 (2019)'], ['기술', '문서 작성 · 데이터 분석 · 기획']];
  t.rows.forEach((row, r) => { row.cells[0].blocks = [P([R(rows[r][0], { b: true })], 'Normal', { spc: { after: 0, line: 1, rule: 'auto' } })]; row.cells[0].shd = '#DEEAF6'; row.cells[1].blocks = [P(rows[r][1], 'Normal', { spc: { after: 0, line: 1, rule: 'auto' } })]; row.h = 36; row.cells[0].vAlign = 'center'; row.cells[1].vAlign = 'center'; });
  d.body = [P('이 력 서', 'Title', { align: 'center' }), t, P('자기소개', 'Heading1'), P('지원 동기와 강점을 적습니다.')];
  return d;
}
function minutesDoc() {
  const d = newDoc();
  const t = table(4, 2, { style: 'TableGrid', width: cm(16) });
  t.grid = [cm(3), cm(13)];
  const rows = [['일시', '2026년  월  일 ( )  :  ~  :'], ['장소', ''], ['참석자', ''], ['작성자', '']];
  t.rows.forEach((row, r) => { row.cells[0].blocks = [P([R(rows[r][0], { b: true })], 'Normal', { align: 'center', spc: { after: 0 } })]; row.cells[0].shd = '#F2F2F2'; row.cells[1].blocks = [P(rows[r][1], 'Normal', { spc: { after: 0 } })]; });
  d.lists = { 1: numberList() };
  d.body = [P('회 의 록', 'Title', { align: 'center' }), t, P('안건', 'Heading1'), P('안건 1', 'ListParagraph', { num: { id: '1', lvl: 0 } }), P('안건 2', 'ListParagraph', { num: { id: '1', lvl: 0 } }), P('논의 내용', 'Heading1'), P(''), P('결정 사항', 'Heading1'), P(''), P('향후 일정', 'Heading1'), P('')];
  return d;
}
function pressDoc() {
  const d = newDoc();
  d.body = [
    P([R('보 도 자 료', { sz: 20, b: true, color: '#2B579A' })], 'Normal', { align: 'center' }),
    P([R('배포 일시: 2026.  .  . (  )   |   담당: ○○팀 ○○○ (02-000-0000)', { sz: 9, color: '#595959' })], 'Normal', { align: 'right', border: { bottom: { style: 'single', w: 1, color: '#2B579A', space: 2 } } }),
    P('[제목: 핵심 내용을 한 줄로]', 'Heading1', { align: 'center' }),
    P([R('- 부제: 보충 설명을 한두 줄로 -', { color: '#595959' })], 'Normal', { align: 'center' }),
    P('□ 첫 문단에 육하원칙(누가, 언제, 어디서, 무엇을, 어떻게, 왜)을 담습니다.'),
    P('□ 다음 문단에서 배경과 세부 내용을 설명합니다.'),
    P('□ 인용문과 향후 계획으로 마무리합니다.'),
  ];
  return d;
}
function essayDoc() {
  const d = newDoc();
  const n1 = uid('n');
  d.notes.foot[n1] = [P('저자, 「제목」, 출판사, 연도, 쪽.', 'FootnoteText')];
  d.body = [
    P('[논문 제목]', 'Title', { align: 'center' }),
    P('[이름] · [소속]', 'Normal', { align: 'center' }),
    P('초록', 'Heading1'), P('연구의 목적, 방법, 결과, 결론을 200자 안팎으로 적습니다.'),
    P('Ⅰ. 서론', 'Heading1'), P([R('선행 연구를 인용할 때 각주를 씁니다.'), { obj: 'note', id: n1, kind: 'foot' }]),
    P('Ⅱ. 본론', 'Heading1'), P(''),
    P('Ⅲ. 결론', 'Heading1'), P(''),
    P('참고 문헌', 'Heading1'), P('저자 (연도). 제목. 출판사.', 'Bibliography', { ind: { l: 32, hanging: 32 } }),
  ];
  d.section = newSection({ footers: { default: [P([pageNum()], 'Footer', { align: 'center' })] } });
  return d;
}
function hwpDoc() {
  const d = newDoc();
  d.themeFonts = { major: { latin: '함초롬돋움', ea: '함초롬돋움' }, minor: { latin: '함초롬바탕', ea: '함초롬바탕' } };
  d.defaults = { rPr: { font: '+mn', fontEa: '+mn', sz: 10, lang: 'ko-KR' }, pPr: { align: 'justify', spc: { before: 0, after: 0, line: 1.6, rule: 'pct' }, widow: true } };
  d.section = newSection({ page: { w: PAPERS.A4.w, h: PAPERS.A4.h, orient: 'portrait' }, margin: { t: MARGINS.hwp.t, b: MARGINS.hwp.b, l: MARGINS.hwp.l, r: MARGINS.hwp.r, header: MARGINS.hwp.header, footer: MARGINS.hwp.footer, gutter: 0 } });
  d.body = [P('')];
  return d;
}
