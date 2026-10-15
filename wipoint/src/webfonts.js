// 인터넷에 공개된 (무료 · 웹 삽입 가능) 한글 글꼴 목록 (DOM 없음)
//   google: Google Fonts (fonts.googleapis.com) — 쓰는 순간 내려받음, 화면 · 인쇄 · PNG 에 그대로
//   cdn: 저작권자가 웹 글꼴로 배포하는 CSS (jsDelivr 등). pptx 로 저장하면 글꼴 이름만 들어가므로 받는 PC 에도 그 글꼴이 있어야 같게 보임
//   archive: 무료 한글 글꼴 모음 (fontarchive.js, 565종 — 눈누에 올라온 기관 · 기업 무료 글꼴 대부분)
import { ARCHIVE_FONTS_TEXT } from './fontarchive.js';
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
// 무료 한글 글꼴 모음 (이미 위에 있는 글꼴은 위의 것을 씀)
const ARCHIVE_CSS = (n) => `https://cdn.jsdelivr.net/gh/fonts-archive/${encodeURIComponent(n)}/${encodeURIComponent(n)}.css`;
{
  const have = new Set(WEB_FONTS.map((f) => f.family.toLowerCase().replace(/[\s-]/g, '')));
  for (const line of ARCHIVE_FONTS_TEXT.split('\n')) {
    const [family, ko, cat] = line.split('|');
    if (!family || !cat || have.has(family.toLowerCase().replace(/[\s-]/g, ''))) continue;
    WEB_FONTS.push({ family, ko, cat, src: 'archive', css: ARCHIVE_CSS(family) });
  }
}
const byName = new Map();
for (const f of WEB_FONTS) { byName.set(f.family.toLowerCase(), f); byName.set(f.ko.toLowerCase(), f); byName.set(f.family.toLowerCase().replace(/[\s_-]/g, ''), f); }
// 같은 글꼴의 한글 이름 (PowerPoint 파일에 적힌 이름)
for (const [k, fam] of [['나눔고딕', 'Nanum Gothic'], ['나눔명조', 'Nanum Myeongjo'], ['나눔스퀘어', 'NanumSquare'], ['본고딕', 'Noto Sans KR'], ['본명조', 'Noto Serif KR'], ['프리텐다드', 'Pretendard'], ['나눔손글씨 펜', 'Nanum Pen Script'], ['나눔손글씨 붓', 'Nanum Brush Script'], ['나눔고딕코딩', 'Nanum Gothic Coding'], ['NanumGothic', 'Nanum Gothic'], ['NanumMyeongjo', 'Nanum Myeongjo'], ['Gmarket Sans', 'GmarketSans'], ['G마켓 산스', 'GmarketSans'], ['에스코어 드림', 'S-CoreDream'], ['S-Core Dream', 'S-CoreDream']]) byName.set(k.toLowerCase(), byName.get(fam.toLowerCase()));

// PowerPoint 파일에 흔히 적히는 설치 글꼴 이름 → 같은 무료 웹 글꼴
const ALIASES = [
  ['배달의민족 주아', 'BMJUA'], ['BM JUA_TTF', 'BMJUA'], ['BM JUA_OTF', 'BMJUA'], ['배달의민족 도현', 'BMDOHYEON'], ['BM DoHyeon', 'BMDOHYEON'], ['BM DoHyeon OTF', 'BMDOHYEON'],
  ['배달의민족 한나는 열한살', 'BMHANNA11yrs'], ['BM HANNA 11yrs old', 'BMHANNA11yrs'], ['배달의민족 한나체 Pro', 'BMHANNAPro'], ['BM HANNA Pro', 'BMHANNAPro'], ['배달의민족 한나체 Air', 'BMHANNAAir'],
  ['배달의민족 연성', 'BMYEONSUNG'], ['BM YEONSUNG', 'BMYEONSUNG'], ['배달의민족 기랑해랑', 'BMKIRANGHAERANG'], ['배달의민족 을지로체', 'BMEULJIRO'], ['배달의민족 을지로10년후체', 'BMEuljiro10yearslater'], ['배달의민족 을지로오래오래체', 'BMEuljiroOraeorae'], ['배달의민족 꾸불림', 'BMKkubulim'],
  ['나눔바른고딕', 'NanumBarunGothic'], ['나눔바른펜', 'NanumBarunpen'], ['나눔스퀘어라운드', 'NanumSquareRound'], ['나눔스퀘어 라운드', 'NanumSquareRound'], ['나눔스퀘어 네오', 'NanumSquareNeo'], ['나눔스퀘어네오', 'NanumSquareNeo'], ['나눔휴먼', 'NanumHuman'],
  ['넥슨Lv1고딕', 'NEXONLv1Gothic'], ['넥슨Lv2고딕', 'NEXONLv2Gothic'], ['NEXON Lv1 Gothic', 'NEXONLv1Gothic'], ['NEXON Lv2 Gothic', 'NEXONLv2Gothic'], ['넥슨 풋볼고딕', 'NEXONFootballGothic'], ['메이플스토리', 'Maplestory'],
  ['Cafe24 Ssurround', 'Cafe24Ssurround'], ['카페24 써라운드', 'Cafe24Ssurround'], ['Cafe24 Ohsquare', 'Cafe24Ohsquare'], ['Cafe24 Simplehae', 'Cafe24Simplehae'], ['Cafe24 Danjunghae', 'Cafe24Danjunghae'], ['Cafe24 Dangdanghae', 'Cafe24Dangdanghae'], ['Cafe24 Shiningstar', 'Cafe24Shiningstar'],
  ['여기어때 잘난체', 'Jalnan'], ['잘난체', 'Jalnan'], ['여기어때 잘난체 고딕', 'JalnanGothic'], ['티몬 몬소리', 'TmonMonsori'], ['티몬 몬소리 Black', 'TmonMonsori'], ['티머니 둥근바람', 'TmoneyRoundWind'],
  ['KoPub돋움체', 'KoPubDotum'], ['KoPub바탕체', 'KoPubBatang'], ['KoPubWorld돋움체', 'KoPubDotum'], ['KoPubWorld바탕체', 'KoPubBatang'], ['KoPubWorld돋움체 Medium', 'KoPubDotum'], ['KoPubWorld바탕체 Medium', 'KoPubBatang'],
  ['마루 부리', 'MaruBuri'], ['마루부리', 'MaruBuri'], ['아리따 돋움', 'AritaDotum'], ['아리따돋움', 'AritaDotum'], ['아리따 부리', 'AritaBuri'], ['아리따부리', 'AritaBuri'], ['어그로체', 'Aggro'], ['이사만루', 'esamanru'],
  ['LINE Seed Sans KR', 'LINESeedKR'], ['LINE Seed KR', 'LINESeedKR'], ['스포카 한 산스', 'SpoqaHanSans'], ['수트', 'SUIT'], ['스위트', 'SUITE'], ['프리젠테이션', 'Freesentation'], ['리디바탕', 'RIDIBatang'],
  ['조선일보명조', 'ChosunSm'], ['쿠키런', 'CookieRun'], ['CookieRun Regular', 'CookieRun'], ['CookieRun Bold', 'CookieRun'], ['한글누리', 'HangeulNuri'], ['제주고딕', 'JejuGothic'], ['제주명조', 'JejuMyeongjo'], ['제주한라산', 'JejuHallasan'],
  ['경기천년제목', 'GyeonggiTitle'], ['경기천년바탕', 'GyeonggiBatang'], ['서울남산체', 'SeoulNamsan'], ['서울한강체', 'SeoulHangang'], ['던파 비트비트체', 'DNFBitBit'], ['둥근모꼴', 'DungGeunMo'], ['갈무리11', 'Galmuri11'],
  ['카카오 큰글씨', 'KakaoBigSans'], ['카카오 작은글씨', 'KakaoSmallSans'], ['KBO 다이아고딕', 'KBODiaGothic'], ['롯데마트 드림', 'LotteMartDream'], ['롯데마트 행복', 'LotteMartHappy'], ['G마켓 산스 TTF', 'GmarketSans'], ['Gmarket Sans TTF', 'GmarketSans'],
];
for (const [k, fam] of ALIASES) { const f = byName.get(fam.toLowerCase()); if (f && !byName.has(k.toLowerCase())) byName.set(k.toLowerCase(), f); }

/** 글꼴 이름 → 웹 글꼴 정보 (없으면 null). 굵기 · 꼴 이름이 붙은 이름(… Bold, … ExtraBold)도 찾음 */
export const webFont = (name) => {
  const k = String(name ?? '').trim().toLowerCase();
  if (!k) return null;
  return byName.get(k) ?? byName.get(k.replace(/[\s_-]/g, '')) ?? byName.get(k.replace(/[\s_-]*(thin|extralight|ultralight|light|regular|medium|semibold|demibold|bold|extrabold|ultrabold|heavy|black|ttf|otf|b|l|m|r|eb|el)$/i, '').replace(/[\s_-]/g, '')) ?? null;
};
export const WEB_FONT_CATS = ['고딕', '명조', '디자인', '손글씨', '코딩', '픽셀'];

/** 문서에서 쓰는 글꼴 이름들 (테마 + 글자 + 표 + 마스터) */
export function fontsUsed(pres) {
  const out = new Set([pres.theme?.fonts?.major, pres.theme?.fonts?.minor].filter(Boolean));
  const body = (t) => { for (const p of t?.paras ?? []) { for (const r of p.runs) if (r.font && !r.font.startsWith('+')) out.add(r.font); if (p.end?.font && !p.end.font.startsWith('+')) out.add(p.end.font); } };
  const objs = (list) => { for (const o of list ?? []) { body(o.text); for (const row of o.rows ?? []) for (const c of row.cells ?? []) body(c.text); } };
  for (const s of pres.slides ?? []) { objs(s.objects); objs(s.bgObjects); }
  for (const m of Object.values(pres.masters ?? {})) objs(m.objects);
  return [...out];
}
