// Excel Workbook has no RepairMode property. A coerced false value from older
// PowerShell logs is not evidence; use the Open policy and completed reads.
export function nativeOpenEvidence(record) {
  const result = { passed: false, status: 'normal-open-not-passed', policy: null, repairModePropertyIgnored: true };
  if (!record) return { ...result, reason: 'not-run' };
  const corruptLoad = record.corruptLoad ?? record.loadPolicy?.corruptLoad;
  if (record.mode === 'repair-diagnostic-only' || record.repairDiagnostic === true ||
      record.mode === 'extract-diagnostic-only' || (corruptLoad != null && corruptLoad !== 0) ||
      record.repairMode === true || record.repairLogs?.length) return { ...result, reason: 'diagnostic-or-repair-load' };
  if (record.opened !== true || record.readOnly !== true || record.errors?.length ||
      !(record.sourceUnchanged === true || record.unchanged === true)) return { ...result, reason: 'open-or-preservation-incomplete' };
  const sheets = record.sheets;
  const inspected = Array.isArray(sheets) && sheets.length > 0 && sheets.every(sheet =>
    typeof sheet.name === 'string' && (sheet.sample?.length > 0 ||
      (Array.isArray(sheet.pivots) && Array.isArray(sheet.shapes))));
  if (!inspected) return { ...result, reason: 'structure-read-incomplete' };
  if (corruptLoad === 0) {
    if (record.readinessVerified !== true || record.fullNameMatches !== true ||
        record.readyWorksheetCount !== sheets.length) return { ...result, reason: 'ready-state-unverified' };
    return { ...result, passed: true, status: 'normal-open-completed', policy: 'explicit-xlNormalLoad' };
  }
  // The earlier checked-in helper recorded mode:'normal' and called Open with
  // CorruptLoad omitted (documented xlNormalLoad default). Keep those actual
  // completed inspections, explicitly separated from the unsupported property.
  if (record.mode === 'normal') return { ...result, passed: true, status: 'normal-open-completed', policy: 'legacy-default-xlNormalLoad' };
  return { ...result, reason: 'load-policy-unrecorded' };
}
