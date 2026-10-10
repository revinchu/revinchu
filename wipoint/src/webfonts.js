// 인터넷에 공개된 (무료 · 웹 삽입 가능) 한글 글꼴 목록 (DOM 없음)
//   google: Google Fonts (fonts.googleapis.com) — 쓰는 순간 내려받음, 화면 · 인쇄 · PNG 에 그대로
//   cdn: 저작권자가 웹 글꼴로 배포하는 CSS (jsDelivr 등). pptx 로 저장하면 글꼴 이름만 들어가므로 받는 PC 에도 그 글꼴이 있어야 같게 보임
const G = (family, ko, cat, w = '400;700') => ({ family, ko, cat, src: 'google', css: `https://fonts.googleapis.com/css2?family=${family.replace(/ /g, '+')}${w ? `:wght@${w}` : ''}&display=swap` });
const C = (family, ko, cat, css) => ({ family, ko, cat, src: 'cdn', css });

export const WEB_FONTS = [
  // 고딕 (본문 · 제목)
  G('Noto Sans KR', '본고딕 (Noto Sans KR)', '고딕', '100..900'),
  G('Nanum Gothic', '나눔고딕', '고딕', '400;700;800'),
  G('Gothic A1', '고딕 A1', '고딕', '100;300;400;500;700;900'),
  G('IBM Plex Sans KR', 'IBM Plex Sans KR', '고딕', '100;300;400;500;700'),
  G('Gowun Dodum', '고운돋움', '고딕', ''),
  G('Sunflower', '해바라기', '고딕', '300;500;700'),
  G('Orbit', '오르빗', '고딕', ''),
  C('Pretendard', '프리텐다드', '고딕', 'https://cdn.jsdelivr.net/gh/orioncactus/pretendard@v1.3.9/dist/web/static/pretendard.min.css'),
  C('SUIT', 'SUIT', '고딕', 'https://cdn.jsdelivr.net/gh/sun-typeface/SUIT@2/fonts/static/woff2/SUIT.css'),
  C('Spoqa Han Sans Neo', '스포카 한 산스 Neo', '고딕', 'https://spoqa.github.io/spoqa-han-sans/css/SpoqaHanSansNeo.css'),
  C('Wanted Sans', '원티드 산스', '고딕', 'https://cdn.jsdelivr.net/gh/wanteddev/wanted-sans@v1.0.3/packages/wanted-sans/fonts/webfonts/static/split/WantedSans.min.css'),
  C('NanumSquare', '나눔스퀘어', '고딕', 'https://cdn.jsdelivr.net/gh/moonspam/NanumSquare@2.0/nanumsquare.css'),
  C('GmarketSans', 'G마켓 산스', '고딕', 'https://cdn.jsdelivr.net/gh/fonts-archive/GmarketSans/GmarketSans.css'),
  C('S-CoreDream', '에스코어 드림', '고딕', 'https://cdn.jsdelivr.net/gh/fonts-archive/S-CoreDream/S-CoreDream.css'),
  C('Paperlogy', '페이퍼로지', '고딕', 'https://cdn.jsdelivr.net/gh/fonts-archive/Paperlogy/Paperlogy.css'),
  // 명조 (바탕)
  G('Noto Serif KR', '본명조 (Noto Serif KR)', '명조', '200..900'),
  G('Nanum Myeongjo', '나눔명조', '명조', '400;700;800'),
  G('Gowun Batang', '고운바탕', '명조', '400;700'),
  G('Song Myung', '송명', '명조', ''),
  G('Hahmlet', '함렛', '명조', '100..900'),
  G('Diphylleia', '디필레이아', '명조', ''),
  // 제목 · 디자인
  G('Black Han Sans', '검은고딕', '디자인', ''),
  G('Do Hyeon', '도현 (배달의민족)', '디자인', ''),
  G('Jua', '주아 (배달의민족)', '디자인', ''),
  G('Dongle', '동글', '디자인', '300;400;700'),
  G('Gugi', '구기', '디자인', ''),
  G('Black And White Picture', '흑백사진', '디자인', ''),
  G('Stylish', '스타일리시', '디자인', ''),
  G('Yeon Sung', '연성', '디자인', ''),
  G('Bagel Fat One', '베이글 팻 원', '디자인', ''),
  G('Gasoek One', '가석 원', '디자인', ''),
  G('Moirai One', '모이라이 원', '디자인', ''),
  G('Grandiflora One', '그랜디플로라 원', '디자인', ''),
  // 손글씨
  G('Nanum Pen Script', '나눔손글씨 펜', '손글씨', ''),
  G('Nanum Brush Script', '나눔손글씨 붓', '손글씨', ''),
  G('Gaegu', '개구', '손글씨', '300;400;700'),
  G('Gamja Flower', '감자꽃', '손글씨', ''),
  G('Hi Melody', '하이멜로디', '손글씨', ''),
  G('Poor Story', '푸어스토리', '손글씨', ''),
  G('Cute Font', '귀여운 글꼴', '손글씨', ''),
  G('Single Day', '싱글데이', '손글씨', ''),
  G('Dokdo', '독도', '손글씨', ''),
  G('East Sea Dokdo', '동해독도', '손글씨', ''),
  G('Kirang Haerang', '기랑해랑', '손글씨', ''),
  // 코딩
  G('Nanum Gothic Coding', '나눔고딕코딩', '코딩', '400;700'),
  C('D2Coding', 'D2Coding', '코딩', 'https://cdn.jsdelivr.net/gh/joungkyun/font-d2coding/d2coding.css'),
];
const byName = new Map();
for (const f of WEB_FONTS) { byName.set(f.family.toLowerCase(), f); byName.set(f.ko.toLowerCase(), f); }
// 같은 글꼴의 한글 이름 (PowerPoint 파일에 적힌 이름)
for (const [k, fam] of [['나눔고딕', 'Nanum Gothic'], ['나눔명조', 'Nanum Myeongjo'], ['나눔스퀘어', 'NanumSquare'], ['본고딕', 'Noto Sans KR'], ['본명조', 'Noto Serif KR'], ['프리텐다드', 'Pretendard'], ['나눔손글씨 펜', 'Nanum Pen Script'], ['나눔손글씨 붓', 'Nanum Brush Script'], ['나눔고딕코딩', 'Nanum Gothic Coding'], ['NanumGothic', 'Nanum Gothic'], ['NanumMyeongjo', 'Nanum Myeongjo'], ['Gmarket Sans', 'GmarketSans'], ['G마켓 산스', 'GmarketSans'], ['에스코어 드림', 'S-CoreDream'], ['S-Core Dream', 'S-CoreDream']]) byName.set(k.toLowerCase(), byName.get(fam.toLowerCase()));

/** 글꼴 이름 → 웹 글꼴 정보 (없으면 null) */
export const webFont = (name) => byName.get(String(name ?? '').trim().toLowerCase()) ?? null;
export const WEB_FONT_CATS = ['고딕', '명조', '디자인', '손글씨', '코딩'];

/** 문서에서 쓰는 글꼴 이름들 (테마 + 글자 + 표 + 마스터) */
export function fontsUsed(pres) {
  const out = new Set([pres.theme?.fonts?.major, pres.theme?.fonts?.minor].filter(Boolean));
  const body = (t) => { for (const p of t?.paras ?? []) { for (const r of p.runs) if (r.font && !r.font.startsWith('+')) out.add(r.font); if (p.end?.font && !p.end.font.startsWith('+')) out.add(p.end.font); } };
  const objs = (list) => { for (const o of list ?? []) { body(o.text); for (const row of o.rows ?? []) for (const c of row.cells ?? []) body(c.text); } };
  for (const s of pres.slides ?? []) { objs(s.objects); objs(s.bgObjects); }
  for (const m of Object.values(pres.masters ?? {})) objs(m.objects);
  return [...out];
}
