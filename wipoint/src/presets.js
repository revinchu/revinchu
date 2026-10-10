// 빠른 스타일 · WordArt · 그림 스타일 · 표 스타일 미리 정의 (DOM 없음)

const ACC = ['dk1', 'accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6'];

/** 도형 빠른 스타일: 7색 × 4종류 (진한 채우기 · 밝은 채우기 · 윤곽선만 · 그라데이션) */
export const SHAPE_STYLES = [];
for (const v of [0, 1, 2, 3]) {
  for (const a of ACC) {
    const c = `@${a}`;
    if (v === 0) SHAPE_STYLES.push({ fill: { type: 'solid', color: c }, line: { color: `${c}:s75`, width: 1.33 }, text: '@lt1', label: '색 채우기' });
    if (v === 1) SHAPE_STYLES.push({ fill: { type: 'solid', color: a === 'dk1' ? '@dk1:lm20:lo80' : `${c}:lm20:lo80` }, line: { color: c, width: 1.33 }, text: a === 'dk1' ? '@dk1' : `${c}:lm75`, label: '밝은 채우기' });
    if (v === 2) SHAPE_STYLES.push({ fill: { type: 'solid', color: '@lt1' }, line: { color: c, width: 1.33 }, text: a === 'dk1' ? '@dk1' : c, label: '윤곽선만' });
    if (v === 3) SHAPE_STYLES.push({ fill: { type: 'gradient', angle: 90, stops: [[0, a === 'dk1' ? '@dk1:lm50:lo50' : `${c}:lm60:lo40`], [1, c]] }, line: null, text: '@lt1', label: '그라데이션' });
  }
}

/** WordArt (글자 효과) 미리 정의 — 글자 서식으로 적용 */
export const WORDART = [
  { label: '채우기: 검정, 텍스트 색 1, 그림자', run: { color: '@tx1', shadow: true, b: true } },
  { label: '채우기: 강조 1, 그림자', run: { color: '@accent1', shadow: true, b: true } },
  { label: '채우기: 흰색, 윤곽선: 강조 2', run: { color: '@lt1', outline: { color: '@accent2', width: 1.5 }, b: true } },
  { label: '채우기: 강조 2, 윤곽선: 흰색', run: { color: '@accent2', outline: { color: '@lt1', width: 1 }, shadow: true, b: true } },
  { label: '윤곽선: 강조 1', run: { color: '@lt1', outline: { color: '@accent1', width: 1.5 } } },
  { label: '채우기: 강조 4, 그림자', run: { color: '@accent4', shadow: true, b: true } },
  { label: '채우기: 회색, 그림자', run: { color: '@tx1:lm50:lo50', shadow: true } },
];

/** 그림 스타일 */
export const PICTURE_STYLES = [
  { label: '단순형 프레임, 흰색', set: { line: { color: '#FFFFFF', width: 8 }, shadow: true, shape: undefined } },
  { label: '단순형 프레임, 검정', set: { line: { color: '#000000', width: 6 }, shadow: undefined, shape: undefined } },
  { label: '모서리가 둥근 사각형', set: { shape: 'roundRect', line: undefined, shadow: true } },
  { label: '그림자 사각형', set: { shape: undefined, line: undefined, shadow: true } },
  { label: '타원', set: { shape: 'ellipse', line: { color: '#FFFFFF', width: 4 }, shadow: true } },
  { label: '얇은 테두리', set: { shape: undefined, line: { color: '@tx1:lm50:lo50', width: 1 }, shadow: undefined } },
];

/** 표 스타일 목록: [강조 색, 종류] */
export const TABLE_STYLES = [];
for (const kind of ['medium', 'light', 'none']) {
  for (const a of ['accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6']) {
    if (kind === 'none' && a !== 'accent1') continue;
    TABLE_STYLES.push({ accent: a, kind, label: kind === 'none' ? '스타일 없음, 표 눈금 없음' : `${kind === 'light' ? '밝은 스타일' : '보통 스타일 2'} - 강조 ${a.slice(-1)}` });
  }
}
export function tableStyleProps(ts) {
  if (ts.kind === 'none') return { none: true };
  return { accent: ts.accent, light: ts.kind === 'light' || undefined, none: undefined };
}

/** 디자인 아이디어: 슬라이드 배치 몇 가지 (제목 + 내용을 다시 배치) */
export const DESIGN_IDEAS = [
  { label: '왼쪽 색 띠', bg: null, band: { x: 0, y: 0, w: 0.36, h: 1, color: '@accent1' }, title: { x: 0.04, y: 0.12, w: 0.28, h: 0.5, color: '@lt1', align: 'l' }, body: { x: 0.42, y: 0.12, w: 0.52, h: 0.76 } },
  { label: '위쪽 제목 띠', band: { x: 0, y: 0, w: 1, h: 0.24, color: '@accent1' }, title: { x: 0.06, y: 0.03, w: 0.88, h: 0.18, color: '@lt1', align: 'l' }, body: { x: 0.06, y: 0.3, w: 0.88, h: 0.62 } },
  { label: '어두운 배경', bg: { type: 'solid', color: '@dk2' }, title: { x: 0.07, y: 0.08, w: 0.86, h: 0.2, color: '@lt1', align: 'l' }, body: { x: 0.07, y: 0.32, w: 0.86, h: 0.6, color: '@lt1' } },
  { label: '가운데 카드', bg: { type: 'solid', color: '@accent1:lm20:lo80' }, card: { x: 0.08, y: 0.1, w: 0.84, h: 0.8 }, title: { x: 0.12, y: 0.14, w: 0.76, h: 0.18, align: 'ctr' }, body: { x: 0.12, y: 0.36, w: 0.76, h: 0.5 } },
  { label: '오른쪽 강조 원', band: { x: 0.72, y: -0.2, w: 0.5, h: 0.9, color: '@accent2:a35', shape: 'ellipse' }, title: { x: 0.06, y: 0.08, w: 0.66, h: 0.2, align: 'l' }, body: { x: 0.06, y: 0.32, w: 0.6, h: 0.6 } },
];
