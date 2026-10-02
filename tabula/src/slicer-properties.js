// 슬라이서 크기 입력은 cm, 저장 좌표는 기존 CSS px(96dpi)입니다.
export const SLICER_COLUMNS_MAX = 20000;
export const SLICER_DEFAULT_BUTTON_HEIGHT = 24;
// Excel 16.0 합성 검증: 양옆/스크롤 영역 18px, 열 사이 기본 간격 3px.
export function slicerDimensions(slicer) {
  const positive = (value, fallback) => Number.isFinite(value) && value > 0 ? value : fallback;
  const columns = Math.max(1, Math.min(SLICER_COLUMNS_MAX, Math.floor(positive(slicer.columns, 1))));
  const gap = Number.isFinite(slicer.gap) && slicer.gap >= 0 ? slicer.gap : 3;
  const w = positive(slicer.w, 180), h = positive(slicer.h, 200);
  // 매우 많은 열이 좁은 틀에 들어오면 음수 트랙 대신 0px을 사용한다.
  const buttonWidth = positive(slicer.buttonWidth, Math.max(0, (w - 18 - gap * (columns - 1)) / columns));
  return { columns, gap, buttonHeight: positive(slicer.buttonHeight, SLICER_DEFAULT_BUTTON_HEIGHT), buttonWidth, w, h };
}

function dimensionPx(value, label, min, current) {
  const n = Number(value);
  if (!String(value ?? '').trim() || !Number.isFinite(n) || n < 0) throw new Error(`${label}에 0 이상의 숫자를 입력하세요.`);
  const px = n * 96 / 2.54;
  // 리본의 cm 소수 둘째 자리 표시값을 그대로 확정해도 기존 정밀 치수는 유지한다.
  if (Number.isFinite(current) && n === Number((current * 2.54 / 96).toFixed(2))) return current;
  if (px < min - .02 || px > 20000) throw new Error(`${label}는 ${(min * 2.54 / 96).toFixed(2)}~529.16cm 범위로 입력하세요.`);
  return Math.round(px * 1000) / 1000;
}

/** columns는 정수, 그 외 치수는 cm. 전체 폭 변경/열 변경은 단추 자동 폭으로 돌아간다. */
export function slicerDimensionPatch(slicer, kind, value) {
  const d = slicerDimensions(slicer), patch = {};
  if (kind === 'columns') {
    const n = Number(value);
    if (!String(value ?? '').trim() || !Number.isInteger(n) || n < 1 || n > SLICER_COLUMNS_MAX) throw new Error(`단추 열 수는 1~${SLICER_COLUMNS_MAX}의 정수로 입력하세요.`);
    if (n === d.columns) return patch;
    patch.columns = n; patch.buttonWidth = undefined;
  } else if (kind === 'buttonWidth') {
    const width = dimensionPx(value, '단추 너비', 1, d.buttonWidth);
    if (width === d.buttonWidth) return patch;
    const w = Math.round((width * d.columns + d.gap * (d.columns - 1) + 18) * 1000) / 1000;
    if (w < 80 || w > 20000) throw new Error('단추 너비와 열 수로 정한 전체 너비는 2.12~529.16cm 범위여야 합니다.');
    patch.w = w; patch.buttonWidth = undefined;
  } else if (['buttonHeight', 'w', 'h'].includes(kind)) {
    const label = { buttonHeight: '단추 높이', w: '전체 너비', h: '전체 높이' }[kind];
    const px = dimensionPx(value, label, kind === 'w' ? 80 : kind === 'h' ? 56 : 1, d[kind]);
    if (px === d[kind]) return patch;
    patch[kind] = px;
    if (kind === 'w') patch.buttonWidth = undefined;
  } else throw new Error('슬라이서 크기 항목을 확인하세요.');
  if (slicer.noMove) throw new Error('크기·위치를 바꾸려면 [크기 조정 및 이동 사용 안 함]을 해제하세요.');
  return patch;
}

export function slicerSizePatch(slicer, values) {
  const patch = {};
  for (const [key, label, min] of [['x', '가로 위치', 0], ['y', '세로 위치', 0], ['w', '너비', 80], ['h', '높이', 56]]) {
    const n = Number(values[key]);
    if (!String(values[key] ?? '').trim() || !Number.isFinite(n) || n < 0) throw new Error(`${label}에 0 이상의 숫자를 입력하세요.`);
    const px = n * 96 / 2.54;
    if (px < min - .02 || px > 20000) throw new Error(`${label}는 ${min ? (min * 2.54 / 96).toFixed(2) : '0'}~529.16cm 범위로 입력하세요.`);
    patch[key] = Math.abs(px - slicer[key]) < .025 ? slicer[key] : Math.round(px * 1000) / 1000;
  }
  if (!['twoCell', 'oneCell', 'absolute'].includes(values.placement)) throw new Error('개체 위치 속성을 선택하세요.');
  if (values.noMove && ['x', 'y', 'w', 'h'].some(key => patch[key] !== slicer[key])) throw new Error('크기·위치를 바꾸려면 [크기 조정 및 이동 사용 안 함]을 해제하세요.');
  return { ...patch, ...(patch.w !== slicer.w && slicer.buttonWidth !== undefined ? { buttonWidth: undefined } : {}), placement: values.placement, locked: values.locked ? undefined : false, noMove: values.noMove || undefined };
}
export function slicerSourceKey(source) {
  if (!source || !Number.isInteger(source.si) || source.si < 0 || !source.ref) return null;
  const { r1, c1, r2, c2 } = source.ref;
  if (![r1, c1, r2, c2].every(Number.isFinite)) return null;
  return JSON.stringify([source.si, source.table ?? null, r1, c1, r2, c2]);
}
