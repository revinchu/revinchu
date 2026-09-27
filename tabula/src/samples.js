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

export const SAMPLES = [
  { name: '가계부', desc: '수식 · 차트 · 필터 · 틀 고정', build: () => ({ sheets: [budget(), grades()] }) },
];
