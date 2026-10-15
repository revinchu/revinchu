// 개요 (행 · 열 그룹, DOM 없음): 데이터 → 그룹 / 그룹 해제 / 부분합, 수준 단추 1 2 3, 그룹마다 +/− 단추.
// 시트 속성 outline = {
//   rows: { 행: 수준(1~7) }, cols: { 열: 수준 },        그룹에 들어 있는 행 · 열의 수준 (0 은 적지 않음)
//   rowsColl: { 요약 행: true }, colsColl: { 요약 열: true },   접힌 그룹 (요약 행 · 열에 표시, 엑셀과 같음)
//   below: true (요약 행이 그룹 아래), right: true (요약 열이 그룹 오른쪽)
// }
// 접으면 그룹의 행 · 열을 숨기고(hiddenRows / hiddenCols), 펼치면 다시 보이게 함 (안쪽의 접힌 그룹은 그대로)
export const MAX_OUTLINE = 7;

export function normOutline(o) {
  return { rows: {}, cols: {}, rowsColl: {}, colsColl: {}, below: true, right: true, ...(o ?? {}) };
}

/** 개요가 비었는지 */
export const outlineEmpty = (o) => !o || (!Object.keys(o.rows ?? {}).length && !Object.keys(o.cols ?? {}).length);

/** 가장 깊은 수준 */
export function maxLevel(levels) {
  let m = 0;
  for (const v of Object.values(levels ?? {})) if (v > m) m = v;
  return m;
}

/** a~b 의 수준을 delta 만큼 (1~7 사이) 바꾼 새 객체 */
export function changeLevels(levels, a, b, delta) {
  const out = { ...(levels ?? {}) };
  for (let i = a; i <= b; i++) {
    const v = Math.max(0, Math.min(MAX_OUTLINE, (out[i] ?? 0) + delta));
    if (v) out[i] = v; else delete out[i];
  }
  return out;
}

const groupMemo = new WeakMap();
/**
 * 그룹 목록 [{ a, b, level }]: 수준 L 이상인 행이 이어진 구간마다 수준 L 그룹 하나 (바깥 그룹부터)
 */
export function groupsOf(levels) {
  if (!levels) return [];
  const hit = groupMemo.get(levels);
  if (hit) return hit;
  const idx = Object.keys(levels).map(Number).sort((x, y) => x - y);
  const out = [];
  const open = []; // 수준별 열린 그룹
  let prev = -2;
  const closeFrom = (L) => { while (open.length > L) out.push(open.pop()); };
  for (const i of idx) {
    const lv = levels[i];
    if (i !== prev + 1) closeFrom(0); // 끊기면 모두 닫음
    else closeFrom(lv);
    while (open.length < lv) open.push({ a: i, b: i, level: open.length + 1 });
    for (const g of open) g.b = i;
    prev = i;
  }
  closeFrom(0);
  out.sort((x, y) => x.level - y.level || x.a - y.a);
  groupMemo.set(levels, out);
  return out;
}

/** 그룹의 요약 행(열): 아래(오른쪽)면 b+1, 위(왼쪽)면 a-1 */
export const summaryOf = (g, after) => (after ? g.b + 1 : g.a - 1);

/** 인덱스 i 가 들어 있는 가장 안쪽 그룹 (없으면 null). 요약 행 · 열도 그 그룹으로 봄 */
export function groupAt(levels, i, after) {
  let best = null;
  for (const g of groupsOf(levels)) {
    const s = summaryOf(g, after);
    if ((i >= g.a && i <= g.b) || i === s) { if (!best || g.level > best.level) best = g; }
  }
  return best;
}

/**
 * 그룹 접기 · 펼치기 → { hidden: 새 숨김 객체, coll: 새 접힘 객체 }
 * hidden: hiddenRows / hiddenCols, coll: rowsColl / colsColl
 */
export function toggleGroup(levels, hidden, coll, g, after, collapse) {
  const h = { ...(hidden ?? {}) };
  const cl = { ...(coll ?? {}) };
  const s = summaryOf(g, after);
  if (collapse) {
    for (let i = g.a; i <= g.b; i++) h[i] = true;
    cl[s] = true;
  } else {
    delete cl[s];
    for (let i = g.a; i <= g.b; i++) delete h[i];
    // 안쪽의 접힌 그룹은 계속 숨김
    for (const x of groupsOf(levels)) {
      if (x.level <= g.level || x.a < g.a || x.b > g.b || !cl[summaryOf(x, after)]) continue;
      for (let i = x.a; i <= x.b; i++) h[i] = true;
    }
  }
  return { hidden: h, coll: cl };
}

/**
 * 수준 단추 n (1 = 가장 요약된 모양): 수준 n 이상인 행 · 열 숨김, 나머지 보임
 */
export function showLevel(levels, hidden, n, after) {
  const h = { ...(hidden ?? {}) };
  const cl = {};
  for (const [k, lv] of Object.entries(levels ?? {})) {
    if (lv >= n) h[k] = true; else delete h[k];
  }
  for (const g of groupsOf(levels)) if (g.level >= n) cl[summaryOf(g, after)] = true;
  return { hidden: h, coll: cl };
}

// ───────────── 부분합 (데이터 → 부분합) ─────────────
export const SUBTOTAL_FNS = [
  { id: 9, label: '합계' }, { id: 3, label: '개수' }, { id: 1, label: '평균' }, { id: 4, label: '최대' },
  { id: 5, label: '최소' }, { id: 6, label: '곱' }, { id: 2, label: '숫자 개수' }, { id: 7, label: '표준 편차' },
  { id: 8, label: '표준 편차(전체)' }, { id: 10, label: '분산' }, { id: 11, label: '분산(전체)' },
];

/**
 * 부분합 계획: 데이터 행(r1+1~r2, r1 은 머리글)의 keyOf(r) 가 바뀌는 곳마다 요약 행을 넣음.
 * → { breaks: [원래 행 번호 (이 행 앞에 요약 행 삽입, 마지막은 r2+1)], groups: [{ key, a, b }] (원래 행 기준) }
 */
export function planSubtotals(r1, r2, keyOf) {
  const groups = [];
  let cur = null;
  for (let r = r1 + 1; r <= r2; r++) {
    const k = keyOf(r);
    if (!cur || k !== cur.key) { cur = { key: k, a: r, b: r }; groups.push(cur); } else cur.b = r;
  }
  return { groups, breaks: groups.map((g) => g.b + 1) };
}
