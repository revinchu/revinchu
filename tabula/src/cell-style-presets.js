// WIXEL 보고서용 스타일. Microsoft 기본 스타일 이름/테마는 app.js에서 그대로 유지한다.
// 색·글꼴·테두리만 적용: 통화/백분율/날짜, 맞춤, 셀 잠금, 수식 숨김은 바꾸지 않는다.
const INCLUDE = { number: false, alignment: false, font: true, border: true, fill: true, protection: false };
const preset = (name, fill, color, extra = {}) => ({ name, style: { fill, color, ...extra }, include: { ...INCLUDE } });
const edge = (color, style = 'thin') => ({ bb: true, bbc: color, bbs: style });

export const REPORT_CELL_STYLE_SECTIONS = [
  ['보고서 · 제목과 머리글', [
    preset('보고서 네이비', '#183153', '#ffffff', { bold: true, size: 18 }),
    preset('보고서 인디고', '#37306b', '#ffffff', { bold: true, size: 18 }),
    preset('보고서 청록', '#115e59', '#ffffff', { bold: true, size: 18 }),
    preset('머리글 블루', '#e8f0fe', '#183153', { bold: true, ...edge('#3b6ea8', 'medium') }),
    preset('머리글 슬레이트', '#e8edf2', '#27364b', { bold: true, ...edge('#64748b', 'medium') }),
    preset('머리글 플럼', '#f2eafa', '#5b2169', { bold: true, ...edge('#9460a3', 'medium') }),
  ]],
  ['보고서 · KPI와 합계', [
    preset('KPI 블루', '#eff6ff', '#1e40af', { bold: true, size: 15 }),
    preset('KPI 그린', '#ecfdf5', '#166534', { bold: true, size: 15 }),
    preset('KPI 퍼플', '#f5f3ff', '#5b21b6', { bold: true, size: 15 }),
    preset('합계 네이비', '#183153', '#ffffff', { bold: true, ...edge('#ffffff', 'double') }),
    preset('합계 청록', '#115e59', '#ffffff', { bold: true, ...edge('#ffffff', 'double') }),
    preset('소계 스톤', '#f5f3ef', '#44403c', { bold: true, bt: true, btc: '#a8a29e' }),
  ]],
  ['보고서 · 입력과 계산', [
    preset('입력 블루', '#eff6ff', '#1e3a8a', { ...edge('#60a5fa') }),
    preset('입력 민트', '#ecfdf5', '#14532d', { ...edge('#34d399') }),
    preset('입력 크림', '#fffbeb', '#78350f', { ...edge('#fbbf24') }),
    preset('계산 슬레이트', '#f1f5f9', '#334155', { italic: true, ...edge('#94a3b8') }),
    preset('계산 라벤더', '#f5f3ff', '#5b21b6', { italic: true, ...edge('#a78bfa') }),
    preset('참고 블루그레이', '#f8fafc', '#475569', { italic: true }),
  ]],
  ['보고서 · 검토 상태', [
    preset('검토 완료', '#dcfce7', '#166534', { bold: true, ...edge('#22c55e') }),
    preset('검토 진행', '#dbeafe', '#1e40af', { bold: true, ...edge('#3b82f6') }),
    preset('검토 대기', '#fef3c7', '#78350f', { bold: true, ...edge('#d97706') }),
    preset('수정 필요', '#fee2e2', '#991b1b', { bold: true, ...edge('#ef4444') }),
    preset('검토 보류', '#ede9fe', '#5b21b6', { bold: true, ...edge('#8b5cf6') }),
    preset('참고 사항', '#f1f5f9', '#334155', { ...edge('#94a3b8') }),
  ]],
];
