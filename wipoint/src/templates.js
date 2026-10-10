// 서식 파일 (새로 만들기 갤러리) — 직접 만든 원본 구성, DOM 없음
import { newPresentation, newSlide, newShape, newTextBox, newTable, newChart, para, run, scaleRect } from './model.js';
import { smartArt } from './smartart.js';

/** 개체 틀 채우기 도우미 */
function fill(slide, ...texts) {
  const phs = slide.objects.filter((o) => o.ph && o.type === 'shape');
  texts.forEach((t, i) => {
    const o = phs[i];
    if (!o || t == null) return;
    const tpl = o.text.paras[0];
    const lines = Array.isArray(t) ? t : [t];
    o.text.paras = lines.map((line) => {
      const lvl = typeof line === 'string' ? (line.match(/^\s*-+/)?.[0].trim().length ?? 0) : 0;
      const txt = String(line).replace(/^\s*-+\s*/, '');
      return { ...JSON.parse(JSON.stringify({ ...tpl, runs: [] })), lvl, runs: txt ? [run(txt)] : [] };
    });
  });
  return slide;
}
const add = (pres, layout, ...texts) => { const s = fill(newSlide(pres, layout), ...texts); pres.slides.push(s); return s; };

function kpiCards(pres, slide, cards, top = 210) {
  const n = cards.length;
  const W = pres.size.w;
  const margin = 88 * (W / 1280);
  const gap = 24;
  const cw = (W - margin * 2 - gap * (n - 1)) / n;
  cards.forEach(([label, value, delta], i) => {
    const card = newShape('roundRect', { x: margin + i * (cw + gap), y: top, w: cw, h: 190 });
    card.adj = { adj: 8000 };
    card.fill = { type: 'solid', color: '@accent1:lm20:lo80' };
    card.line = null;
    card.text.paras = [
      para(label, { align: 'ctr' }, { size: 16, color: '@tx1:lm65:lo35' }),
      para(value, { align: 'ctr', spcBef: 6 }, { size: 36, b: true, color: '@accent1:lm75' }),
      para(delta ?? '', { align: 'ctr', spcBef: 4 }, { size: 14, color: String(delta ?? '').startsWith('-') || String(delta ?? '').startsWith('▼') ? '#C00000' : '#00A050' }),
    ];
    card.text.anchor = 'ctr';
    slide.objects.push(card);
  });
}

function blank(theme, size = 'wide') {
  const p = newPresentation({ theme, size, firstLayout: null });
  return p;
}

const BUILDERS = {
  blank: () => newPresentation({ theme: 'office' }),

  business: () => {
    const p = blank('navy');
    add(p, 'title', '2027 사업 계획서', '회사명 · 발표자 · 2026년 10월');
    add(p, 'titleContent', '목차', ['사업 개요', '시장 분석', '경쟁 우위', '실행 계획', '재무 계획']);
    add(p, 'section', '사업 개요', '우리가 해결하려는 문제와 방법');
    add(p, 'twoContent', '문제와 해결', ['고객의 문제', '-정보가 흩어져 있어 비교가 어렵다', '-결정에 시간이 오래 걸린다'], ['우리의 해결 방법', '-한 화면에서 비교·추천', '-3분 안에 결정']);
    const m = add(p, 'titleOnly', '시장 규모');
    m.objects.push(newChart(scaleRect(p.size, 140, 190, 1000, 470), { kind: 'col', title: '연도별 시장 규모 (억 원)', legend: null, labels: true, cats: ['2023', '2024', '2025', '2026', '2027(F)'], series: [{ name: '시장 규모', vals: [820, 960, 1150, 1380, 1650] }] }));
    const c = add(p, 'titleOnly', '경쟁 분석');
    const t = newTable(4, 4, scaleRect(p.size, 140, 200, 1000, 300));
    const rows = [['구분', '우리 회사', '경쟁사 A', '경쟁사 B'], ['가격', '월 9,900원', '월 15,000원', '월 12,000원'], ['추천 정확도', '92%', '78%', '81%'], ['고객 지원', '24시간', '평일', '평일']];
    rows.forEach((r, ri) => r.forEach((v, ci) => { t.rows[ri].cells[ci].text.paras[0].runs = [run(v)]; t.rows[ri].cells[ci].text.paras[0].align = ci ? 'ctr' : 'l'; }));
    c.objects.push(t);
    const e = add(p, 'titleOnly', '실행 계획');
    e.objects.push(...smartArt('chevron', ['1분기\n시장 검증', '2분기\n제품 출시', '3분기\n마케팅 확대', '4분기\n해외 진출'], scaleRect(p.size, 90, 250, 1100, 220)));
    add(p, 'titleContent', '재무 계획', ['2027년 매출 목표 120억 원', '-영업이익률 18%', '투자 유치 계획', '-시리즈 A 40억 원 (2027년 2분기)']);
    add(p, 'title', '감사합니다', '문의: contact@example.com');
    return p;
  },

  marketing: () => {
    const p = blank('naver');
    add(p, 'title', '월간 마케팅 성과 보고', '2026년 9월 · 퍼포먼스 마케팅팀');
    const k = add(p, 'titleOnly', '핵심 지표 요약');
    kpiCards(p, k, [['노출수', '1,284만', '▲ 12.4%'], ['클릭수', '41.2만', '▲ 8.1%'], ['전환수', '6,920', '▲ 15.3%'], ['ROAS', '428%', '▼ -3.2%']]);
    const tb = newTextBox(scaleRect(p.size, 88, 450, 1104, 120));
    tb.text.paras = [para('· 검색광고 전환이 크게 늘었고, 디스플레이 광고의 ROAS 는 소폭 하락했습니다.', {}, { size: 18 }), para('· 다음 달에는 성과가 낮은 키워드 120개의 입찰가를 조정합니다.', {}, { size: 18 })];
    k.objects.push(tb);
    const ch = add(p, 'titleOnly', '채널별 성과');
    const t = newTable(5, 6, scaleRect(p.size, 88, 190, 1104, 300));
    const rows = [['채널', '비용', '클릭', 'CPC', '전환', 'ROAS'], ['네이버 검색', '1,820만', '18.2만', '100원', '3,410', '512%'], ['카카오 비즈보드', '940만', '9.8만', '96원', '1,280', '361%'], ['구글 디스플레이', '760만', '8.6만', '88원', '1,420', '345%'], ['메타', '610만', '4.6만', '133원', '810', '298%']];
    rows.forEach((r, ri) => r.forEach((v, ci) => { t.rows[ri].cells[ci].text.paras[0].runs = [run(v)]; t.rows[ri].cells[ci].text.paras[0].align = ci ? 'r' : 'l'; }));
    ch.objects.push(t);
    const tr = add(p, 'titleOnly', '월별 추이');
    tr.objects.push(newChart(scaleRect(p.size, 88, 180, 1104, 480), { kind: 'line', title: '월별 전환수 · 비용', legend: 'b', labels: false, markers: true, cats: ['4월', '5월', '6월', '7월', '8월', '9월'], series: [{ name: '전환수', vals: [4210, 4630, 5120, 5480, 6000, 6920] }, { name: '비용 (만 원)', vals: [3100, 3300, 3550, 3720, 3900, 4130] }] }));
    const r = add(p, 'twoContent', '인사이트와 다음 단계', ['잘된 점', '-브랜드 키워드 전환율 6.8% → 8.1%', '-모바일 랜딩 개선 효과'], ['다음 달 계획', '-저효율 키워드 120개 입찰 조정', '-리타기팅 소재 3종 A/B 테스트', '-주간 이상치 알림 적용']);
    r.notes = '이상치 알림은 WIXEL 의 이상치 감지 도구로 만든 보고서를 사용합니다.';
    add(p, 'title', 'Q&A', '감사합니다');
    return p;
  },

  proposal: () => {
    const p = blank('sunset');
    add(p, 'title', '협업 제안서', '귀사와 함께 만드는 새로운 고객 경험');
    add(p, 'titleContent', '제안 배경', ['고객 행동이 온라인 중심으로 이동', '-모바일 구매 비중 68%', '개인화된 경험에 대한 기대 증가']);
    const s = add(p, 'titleOnly', '제안 내용');
    s.objects.push(...smartArt('blocks', ['데이터 통합', '맞춤 추천', '자동화 캠페인', '성과 분석'], scaleRect(p.size, 140, 200, 1000, 420)));
    const t = add(p, 'titleOnly', '추진 일정');
    t.objects.push(...smartArt('timeline', ['착수 회의', '요구사항 정의', '개발 · 연동', '시범 운영', '정식 오픈'], scaleRect(p.size, 88, 190, 1104, 440)));
    add(p, 'comparison', '기대 효과', '정량적 효과', ['매출 15% 증가', '운영 시간 30% 절감'], '정성적 효과', ['고객 만족도 향상', '브랜드 신뢰 강화']);
    add(p, 'title', '함께 성장하겠습니다', '담당자 · 연락처');
    return p;
  },

  lecture: () => {
    const p = blank('mint');
    add(p, 'title', '데이터 분석 입문', '1주차 · 엑셀과 피벗 테이블');
    add(p, 'titleContent', '학습 목표', ['데이터를 표로 정리할 수 있다', '피벗 테이블로 요약할 수 있다', '차트로 결과를 설명할 수 있다']);
    add(p, 'section', '1. 데이터 정리', '좋은 표의 조건');
    add(p, 'titleContent', '좋은 표의 조건', ['한 열에는 한 가지 정보', '-날짜 · 숫자 · 글자를 섞지 않기', '병합된 셀 쓰지 않기', '머리글은 한 줄로']);
    const c = add(p, 'titleOnly', '분석 과정');
    c.objects.push(...smartArt('process', ['질문 정하기', '데이터 모으기', '정리 · 분석', '결과 공유'], scaleRect(p.size, 88, 260, 1104, 200)));
    add(p, 'titleContent', '실습 과제', ['예제 파일의 월별 매출을 피벗 테이블로 요약하기', '지역별 막대 차트 만들기', '발견한 점 세 가지 적기']);
    add(p, 'title', '수고하셨습니다', '다음 주: 함수와 조건부 서식');
    return p;
  },

  meeting: () => {
    const p = blank('mono');
    add(p, 'title', '주간 회의', '2026년 10월 둘째 주');
    add(p, 'titleContent', '안건', ['지난주 결정 사항 확인', '프로젝트 진행 현황', '이슈와 도움 요청', '다음 주 일정']);
    const s = add(p, 'titleOnly', '프로젝트 진행 현황');
    const t = newTable(5, 4, scaleRect(p.size, 88, 190, 1104, 300));
    [['프로젝트', '담당', '진행률', '상태'], ['웹사이트 개편', '김OO', '80%', '정상'], ['앱 2.0', '이OO', '45%', '주의'], ['데이터 이전', '박OO', '100%', '완료'], ['고객 설문', '최OO', '20%', '정상']].forEach((r, ri) => r.forEach((v, ci) => { t.rows[ri].cells[ci].text.paras[0].runs = [run(v)]; }));
    s.objects.push(t);
    add(p, 'twoContent', '이슈와 결정', ['이슈', '-디자인 시안 확정 지연', '-외부 API 응답 속도'], ['결정', '-금요일까지 시안 확정', '-캐시 적용 후 재측정']);
    add(p, 'titleContent', '다음 주 일정', ['월: 디자인 리뷰', '수: 고객사 미팅', '금: 배포 점검']);
    return p;
  },

  project: () => {
    const p = blank('forest');
    add(p, 'title', '프로젝트 현황 보고', '스마트 물류 시스템 구축');
    const k = add(p, 'titleOnly', '한눈에 보기');
    kpiCards(p, k, [['전체 진행률', '62%', '계획 대비 +4%'], ['예산 집행', '5.4억', '계획의 48%'], ['남은 기간', '14주', '12월 말 완료'], ['위험 요소', '2건', '관리 중']]);
    const o = add(p, 'titleOnly', '조직');
    o.objects.push(...smartArt('hierarchy', ['프로젝트 책임자', '기획', '개발', '품질', '운영'], scaleRect(p.size, 140, 190, 1000, 420)));
    const ch = add(p, 'titleOnly', '단계별 진행률');
    ch.objects.push(newChart(scaleRect(p.size, 140, 180, 1000, 480), { kind: 'bar', title: '', legend: 'b', labels: true, cats: ['분석', '설계', '개발', '테스트', '이행'], series: [{ name: '계획', vals: [100, 100, 70, 30, 0] }, { name: '실적', vals: [100, 100, 76, 28, 0] }] }));
    add(p, 'titleContent', '위험 요소와 대응', ['장비 납품 지연 (중)', '-대체 공급처 확보', '현장 인력 부족 (하)', '-협력사 인력 추가 투입']);
    return p;
  },

  photo: () => {
    const p = blank('slate');
    add(p, 'title', '포토 앨범', '우리의 여행 기록');
    for (const cap of ['첫째 날', '둘째 날', '셋째 날']) add(p, 'pictureCaption', cap, '사진 설명을 입력하십시오');
    add(p, 'title', '다음 여행에서 만나요', '');
    return p;
  },

  pitch: () => {
    const p = blank('berry');
    add(p, 'title', '서비스 소개', '한 줄로 말하는 우리 서비스');
    const s = add(p, 'titleOnly', '숫자로 보는 우리');
    kpiCards(p, s, [['월간 사용자', '32만', '▲ 41%'], ['재방문율', '64%', '▲ 9%p'], ['고객 만족도', '4.8', '5점 만점']]);
    const v = add(p, 'titleOnly', '핵심 가치');
    v.objects.push(...smartArt('venn', ['쉬움', '빠름', '정확함'], scaleRect(p.size, 240, 190, 800, 440)));
    const pr = add(p, 'titleOnly', '성장 단계');
    pr.objects.push(...smartArt('pyramid', ['비전', '플랫폼', '핵심 서비스', '고객 기반'], scaleRect(p.size, 340, 180, 600, 470)));
    add(p, 'title', '함께해 주세요', 'hello@example.com');
    return p;
  },
};

export const TEMPLATES = [
  { id: 'blank', name: '새 프레젠테이션', desc: '빈 제목 슬라이드', theme: 'office' },
  { id: 'business', name: '사업 계획서', desc: '목차 · 시장 · 경쟁 · 실행 계획 (9장)', theme: 'navy' },
  { id: 'marketing', name: '마케팅 성과 보고', desc: 'KPI 카드 · 채널 표 · 추이 차트 (6장)', theme: 'naver' },
  { id: 'proposal', name: '제안서', desc: '배경 · 내용 · 일정 · 기대 효과 (6장)', theme: 'sunset' },
  { id: 'lecture', name: '강의 · 교육 자료', desc: '학습 목표 · 과정 · 실습 (7장)', theme: 'mint' },
  { id: 'meeting', name: '주간 회의', desc: '안건 · 현황 표 · 결정 사항 (5장)', theme: 'mono' },
  { id: 'project', name: '프로젝트 현황 보고', desc: '지표 · 조직도 · 진행률 차트 (5장)', theme: 'forest' },
  { id: 'pitch', name: '서비스 소개', desc: '숫자 · 핵심 가치 · 성장 단계 (5장)', theme: 'berry' },
  { id: 'photo', name: '포토 앨범', desc: '그림 개체 틀 (5장)', theme: 'slate' },
];

export function buildTemplate(id) {
  const b = BUILDERS[id] ?? BUILDERS.blank;
  return b();
}
