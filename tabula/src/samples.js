// 새로 만들기 → 샘플 통합 문서
const HEAD = { fill: '#217346', color: '#ffffff', bold: true, align: 'center', bb: true };
const TOTAL = { bold: true, bt: true, bb: true, fill: '#e2efda' };

function sheet(name, rows, colWidths = {}, extra = {}) {
  const cells = {};
  rows.forEach((row, r) => row.forEach((v, c) => {
    if (v === null || v === undefined) return;
    const cell = typeof v === 'object' ? v : { raw: String(v) };
    cells[`${r},${c}`] = cell;
  }));
  return { name, cells, colWidths, rowHeights: {}, merges: [], cond: [], ...extra };
}

const h = (raw) => ({ raw, style: HEAD });
const n = (raw) => ({ raw: String(raw), style: { numFmt: 'comma' } });
const t = (raw, style = {}) => ({ raw, style: { ...TOTAL, ...style } });

function budget() {
  const items = [
    ['식비', 620000, 580000, 655000],
    ['교통비', 98000, 105000, 92000],
    ['통신비', 65000, 65000, 65000],
    ['주거비', 850000, 850000, 850000],
    ['문화생활', 120000, 45000, 210000],
    ['기타', 73000, 131000, 64000],
  ];
  const rows = [
    [{ raw: '2026년 1분기 가계부', style: { bold: true, size: 16, color: '#1e6f42' } }],
    [],
    ['항목', '1월', '2월', '3월', '합계', '월평균', '비중'].map(h),
    ...items.map((it, i) => {
      const r = i + 4;
      return [it[0], n(it[1]), n(it[2]), n(it[3]),
        n(`=SUM(B${r}:D${r})`), n(`=AVERAGE(B${r}:D${r})`), { raw: `=E${r}/$E$10`, style: { numFmt: 'percent', decimals: 1 } }];
    }),
    [t('합계', { align: 'center' }), ...['B', 'C', 'D', 'E', 'F'].map((c) => t(`=SUM(${c}4:${c}9)`, { numFmt: 'comma' })), t('=SUM(G4:G9)', { numFmt: 'percent', decimals: 1 })],
    [],
    ['최대 지출 항목', { raw: '=INDEX(A4:A9,MATCH(MAX(E4:E9),E4:E9,0))', style: { bold: true, color: '#c00000' } }],
    ['예산(월)', n(2000000)],
    ['예산 초과 월 수', { raw: '=COUNTIF(B10:D10,">"&B13)' }],
  ];
  return sheet('가계부', rows, { 0: 110, 1: 90, 2: 90, 3: 90, 4: 100, 5: 90, 6: 70 }, {
    merges: [{ r1: 0, c1: 0, r2: 0, c2: 6 }],
    rowHeights: { 0: 30 },
    freeze: { rows: 3, cols: 0 },
    charts: [{ id: 'sample-chart', type: 'column', title: '항목별 1분기 지출', range: { r1: 2, c1: 0, r2: 8, c2: 3 }, x: 700, y: 80, w: 470, h: 290 }],
    cond: [{ r1: 3, c1: 4, r2: 8, c2: 4, type: 'bar', color: '#8fd19e' }],
  });
}

function grades() {
  const students = [
    ['김하늘', 92, 88, 95], ['이서준', 78, 85, 80], ['박지우', 65, 72, 58],
    ['최민서', 88, 94, 91], ['정도윤', 55, 61, 70], ['한예린', 97, 90, 99],
  ];
  const rows = [
    ['이름', '국어', '영어', '수학', '평균', '결과', '등급'].map(h),
    ...students.map((s, i) => {
      const r = i + 2;
      return [s[0], s[1], s[2], s[3],
        { raw: `=ROUND(AVERAGE(B${r}:D${r}),1)` },
        { raw: `=IF(E${r}>=70,"합격","불합격")`, style: { align: 'center' } },
        { raw: `=IFS(E${r}>=90,"A",E${r}>=80,"B",E${r}>=70,"C",TRUE,"D")`, style: { align: 'center' } }];
    }),
    [],
    ['과목 평균', ...['B', 'C', 'D', 'E'].map((c) => ({ raw: `=ROUND(AVERAGE(${c}2:${c}7),1)`, style: { bold: true } }))],
    ['합격자 수', { raw: '=COUNTIF(F2:F7,"합격")' }],
  ];
  return sheet('성적표', rows, { 0: 90, 1: 70, 2: 70, 3: 70, 4: 70, 5: 70, 6: 70 }, {
    filter: { r1: 0, c1: 0, r2: 6, c2: 6, criteria: {}, hidden: {} },
    cond: [
      { r1: 1, c1: 5, r2: 6, c2: 5, type: 'eq', v1: '불합격', style: { fill: '#ffc7ce', color: '#9c0006' } },
      { r1: 1, c1: 1, r2: 6, c2: 3, type: 'scale', colors: ['#f8696b', '#ffeb84', '#63be7b'] },
    ],
  });
}

/**
 * 빅데이터 예제: 광고 성과 데이터 n 행을 열 블록(형식화 배열)으로 바로 만듦 + 표 · 피벗 · 슬라이서 · 피벗 차트
 */
function bigData(n) {
  const camps = ['검색_브랜드', '검색_일반', '쇼핑_브랜드', '쇼핑_일반', '디스플레이', '동영상', '리타게팅', '앱설치'];
  const devices = ['모바일', 'PC', '태블릿'];
  const groups = Array.from({ length: 120 }, (_, i) => `그룹_${String(i + 1).padStart(3, '0')}`);
  const regions = ['서울', '경기', '부산', '대구', '인천', '광주', '대전', '울산', '세종', '강원', '충북', '충남', '전북', '전남', '경북', '경남', '제주'];
  const days = 365;
  const d0 = 46023; // 2026-01-01
  const f64 = () => new Float64Array(n);
  const i32 = () => new Int32Array(n);
  const date = f64(); const camp = i32(); const grp = i32(); const dev = i32(); const reg = i32();
  const imp = f64(); const clk = f64(); const cost = f64(); const conv = f64(); const rev = f64();
  let x = 2463534242;
  const rnd = () => { x ^= x << 13; x ^= x >>> 17; x ^= x << 5; return (x >>> 0) / 4294967296; };
  for (let i = 0; i < n; i++) {
    const c = Math.floor(rnd() * camps.length);
    date[i] = d0 + Math.floor(rnd() * days);
    camp[i] = c;
    grp[i] = c * 15 + Math.floor(rnd() * 15);
    dev[i] = rnd() < 0.62 ? 0 : rnd() < 0.85 ? 1 : 2;
    reg[i] = Math.floor(rnd() * regions.length);
    const im = 50 + Math.floor(rnd() * 5000);
    const ck = Math.floor(im * (0.005 + rnd() * 0.05));
    imp[i] = im;
    clk[i] = ck;
    cost[i] = ck * (200 + Math.floor(rnd() * 900));
    const cv = Math.floor(ck * rnd() * 0.12);
    conv[i] = cv;
    rev[i] = cv * (15000 + Math.floor(rnd() * 60000));
  }
  const numCol = (num, fmt = { numFmt: 'comma' }) => ({ num, str: null, dict: [], fmt });
  const strCol = (str, dict) => ({ num: null, str, dict, fmt: null });
  const header = ['날짜', '캠페인', '광고그룹', '기기', '지역', '노출수', '클릭수', '비용', '전환수', '매출'];
  const cells = {};
  header.forEach((h, j) => { cells[`0,${j}`] = { raw: h, style: { bold: true } }; });
  const block = {
    r0: 1, c0: 0, n, ver: 0,
    cols: [numCol(date, { numFmt: 'date' }), strCol(camp, camps), strCol(grp, groups), strCol(dev, devices), strCol(reg, regions),
      numCol(imp), numCol(clk), numCol(cost), numCol(conv), numCol(rev)],
  };
  const calcFields = [
    { name: 'CTR', formula: '클릭수/노출수' }, { name: 'CPC', formula: '비용/클릭수' },
    { name: 'CVR', formula: '전환수/클릭수' }, { name: 'ROAS', formula: '매출/비용' },
  ];
  const val = (field, extra = {}) => ({ field, agg: 'sum', ...extra });
  const pct = { numFmt: 'percent', decimals: 2 };
  const values = [val('비용'), val('클릭수'), val('전환수'), val('매출'), val('CTR', { numFmt: pct }), val('CPC', { numFmt: { numFmt: 'comma' } }), val('ROAS', { numFmt: pct })];
  const pv1 = { name: '캠페인별', table: '광고성과', rows: ['캠페인'], cols: [], values, calcFields, top: 3, left: 1, style: 'PivotStyleMedium9', needsRender: true };
  const pv2 = { name: '월별', table: '광고성과', rows: ['날짜'], cols: [], values: [val('비용'), val('ROAS', { numFmt: pct })], calcFields, top: 16, left: 1, style: 'PivotStyleMedium2', needsRender: true, groups: { 날짜: { by: 'months' } } };
  const pv3 = { name: '기기x지역', table: '광고성과', rows: ['지역'], cols: ['기기'], values: [val('비용')], calcFields, top: 33, left: 1, style: 'PivotStyleLight16', needsRender: true };
  const pivots = [{ sheet: '대시보드', name: '캠페인별' }, { sheet: '대시보드', name: '월별' }, { sheet: '대시보드', name: '기기x지역' }];
  const slicer = (id, field, x, y, h, columns = 1) => ({ id, caption: field, source: { kind: 'pivot', field, pivots }, columns, style: 'SlicerStyleLight1', multi: false, x, y, w: 180, h, z: 10 });
  return {
    sheets: [
      {
        name: '대시보드',
        cells: { '0,1': { raw: `광고 성과 대시보드 — ${n.toLocaleString()}행`, style: { bold: true, size: 16 } } },
        colWidths: { 0: 16, 1: 130 }, rowHeights: { 0: 32 },
        pivot: pv1, pivotsExtra: [pv2, pv3], noGrid: true,
        slicers: [slicer('slb1', '기기', 1010, 60, 130), slicer('slb2', '캠페인', 1010, 200, 290), slicer('slb3', '지역', 1200, 60, 430, 2)],
        charts: [{
          id: 'chb1', type: 'combo', title: '월별 비용 · ROAS', pivot: { sheet: '대시보드', name: '월별' }, x: 430, y: 330, w: 570, h: 300, z: 5,
          seriesFmt: [{ type: 'column', labels: false }, { type: 'line', axis: 1, marker: 'circle', numFmt: '0%' }], legend: 'top',
        }],
      },
      {
        // 셀은 Map 으로 넘겨 불러올 때 블록을 복사하지 않게 (통합 문서가 그대로 소유)
        name: '데이터', cells: new Map(Object.entries(cells)), blocks: [block], freeze: { rows: 1, cols: 0 },
        colWidths: { 0: 88, 1: 96, 2: 88 },
        tables: [{ id: 'tb1', name: '광고성과', r1: 0, c1: 0, r2: n, c2: 9, header: true, totals: false, style: 'TableStyleMedium2', banded: true, filter: null }],
      },
    ],
  };
}

export const SAMPLES = [
  { name: '빅데이터 100만 행', desc: '열 블록 · 표 · 피벗 3개 · 슬라이서 · 피벗 차트', big: true, build: () => bigData(1000000) },
  { name: '빅데이터 1,000만 행', desc: '천만 행 피벗 · 슬라이서 (메모리 약 0.7GB)', big: true, build: () => bigData(10000000) },
  { name: '가계부', desc: '수식 · 차트 · 필터 · 틀 고정', build: () => ({ sheets: [budget(), grades()] }) },
];
