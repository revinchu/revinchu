// 슬라이서 크기 입력은 cm, 저장 좌표는 기존 CSS px(96dpi)입니다.
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
  return { ...patch, placement: values.placement, locked: values.locked ? undefined : false, noMove: values.noMove || undefined };
}
export function slicerSourceKey(source) {
  if (!source || !Number.isInteger(source.si) || source.si < 0 || !source.ref) return null;
  const { r1, c1, r2, c2 } = source.ref;
  if (![r1, c1, r2, c2].every(Number.isFinite)) return null;
  return JSON.stringify([source.si, source.table ?? null, r1, c1, r2, c2]);
}
