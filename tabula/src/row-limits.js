import { MAX_ROWS, EXCEL_MAX_ROWS } from './formula.js';

/** 확장 행은 명시적으로 켠 경우에만 사용한다. 계산 엔진의 한도는 바꾸지 않는다. */
export function worksheetRowLimit(options) {
  return options?.extendedRows === true ? MAX_ROWS : EXCEL_MAX_ROWS;
}

/** 화면 상태에는 두 지원 한도만 허용하며 잘못된 값은 Excel 기본값으로 돌아간다. */
export function normalizeWorksheetRowLimit(value) {
  return value === MAX_ROWS ? MAX_ROWS : EXCEL_MAX_ROWS;
}
