// Run only in an already loaded, isolated local audit page. No file I/O or export.
// Throws with error.audit if an assertion or restoration fails; callers must stop export.
export async function auditStructure(page, { id } = {}) {
  if (!['127.0.0.1', 'localhost', '[::1]'].includes(new URL(page.url()).hostname)) {
    throw new Error('구조 검사는 격리된 로컬 검사 페이지에서만 실행합니다.');
  }
  const report = await page.evaluate(({ id }) => {
    const app = window.tabula, w = app?.wb();
    const out = { id, scope: '앱 명령으로 A1 앞 행·열 삽입, 삽입한 빈 행·열 삭제, Undo/Redo 및 표본 복원. 전체 셀 계산 검사가 아님.', ok: false, restored: false, checks: [], observations: [] };
    const require = (ok, label) => { out.checks.push({ name: label, ok: !!ok }); if (!ok) throw new Error(label); };
    if (!w || !app.selectRange || !app.commands) throw new Error('검사 가능한 앱 API가 없습니다.');
    const commands = app.commands();
    for (const cmd of ['insertRows', 'insertCols', 'deleteRows', 'deleteCols', 'undo', 'redo']) require(commands.includes(cmd), `명령 연결: ${cmd}`);
    const props = ['charts', 'images', 'shapes', 'slicers'];
    const pivots = s => [s.pivot, ...(s.pivotsExtra ?? [])].filter(Boolean);
    const blocked = w.props?.markedFinal || w.props?.lockStructure || ['1', 1, true].includes(w.props?.workbookProtection?.lockStructure);
    const candidates = w.sheets.map((s, si) => ({ s, si, objects: props.reduce((n, p) => n + (s[p]?.length ?? 0), 0), cells: s.cells?.size ?? 0 }))
      .filter(x => !blocked && !x.s.external && !['hidden', 'veryHidden'].includes(x.s.state) && !x.s.protect?.on && x.cells <= 200000 && !(x.s.blocks ?? []).some(b => b.n * b.cols.length > 200000))
      .sort((a, b) => Number(!!(b.objects || pivots(b.s).length)) - Number(!!(a.objects || pivots(a.s).length)) || a.cells - b.cells);
    if (!candidates.length) return { ...out, ok: true, restored: true, skipped: '안전하게 편집할 수 있는 작은 표시 시트가 없습니다.' };
    const si = candidates[0].si;
    out.sheet = si;
    const stable = v => {
      if (v === undefined) return null;
      if (v === null || typeof v !== 'object') return v;
      if (Array.isArray(v)) return v.map(stable);
      return Object.fromEntries(Object.keys(v).sort().filter(k => v[k] !== undefined).map(k => [k, stable(v[k])]));
    };
    const json = v => JSON.stringify(stable(v));
    const digest = text => { let h = 2166136261; for (let i = 0; i < text.length; i++) h = Math.imul(h ^ text.charCodeAt(i), 16777619); return (h >>> 0).toString(16).padStart(8, '0'); };
    const pick = (o, keys) => Object.fromEntries(keys.filter(k => o?.[k] !== undefined).map(k => [k, o[k]]));
    const originalUi = { si: app.si, sel: { ...app.sel }, active: { ...app.active }, sx: app.gv().sx, sy: app.gv().sy };
    const oldUndo = w.undoStack, oldRedo = w.redoStack;
    let baseline, baselineText, pointSets;
    const pointSet = s => {
      const points = new Map(), add = (r, c) => { if (Number.isInteger(r) && Number.isInteger(c) && r >= 0 && c >= 0 && r < 1048575 && c < 16383 && points.size < 120) points.set(`${r},${c}`, { r, c }); };
      for (let r = 0; r < 8; r++) for (let c = 0; c < 8; c++) add(r, c);
      for (const a of [...pivots(s).map(d => d.area), ...(s.tables ?? [])]) if (a) { add(a.r1, a.c1); add(a.r1 + 1, a.c1); add(a.r2, a.c2); }
      return [...points.values()];
    };
    const cell = (sheet, r, c) => { const d = w.getCell(sheet, r, c); return { raw: w.getRaw(sheet, r, c), formula: !!d?.formula, style: w.styleAt(sheet, r, c), comment: d?.comment ?? null, link: d?.link ?? null, image: d?.image ? digest(json(d.image)) : null }; };
    const snapshot = () => ({
      names: w.names,
      sheets: w.sheets.slice(0, w.ownSheetCount()).map((s, sheet) => ({
        cells: pointSets[sheet].map(({ r, c }) => ({ r, c, ...cell(sheet, r, c) })),
        rowSizes: Array.from({ length: 16 }, (_, r) => w.rowHeight(sheet, r)),
        colSizes: Array.from({ length: 16 }, (_, c) => w.colWidth(sheet, c)),
        pivots: pivots(s).map(d => pick(d, ['name', 'source', 'range', 'top', 'left', 'area'])),
        tables: (s.tables ?? []).map(t => pick(t, ['name', 'r1', 'c1', 'r2', 'c2'])),
        objects: props.flatMap(p => (s[p] ?? []).map(o => ({ prop: p, ...pick(o, ['id', 'placement', 'x', 'y', 'w', 'h', 'sheet', 'range', 'linked']) }))),
      })),
    });
    const run = (cmd, expectedUndo) => {
      require(app.wb() === w, `${cmd}: 문서 동일성`);
      app.run(cmd);
      require(w.undoStack.length === expectedUndo, `${cmd}: 실행 취소 기록`);
    };
    // A1 may be merged: the opposite-axis selection kind forces exactly one
    // inserted/deleted row or column in the real app commands despite expansion.
    const atOrigin = axis => app.selectRange({ r1: 0, c1: 0, r2: 0, c2: 0 }, axis === 'row' ? 'cols' : axis === 'col' ? 'rows' : 'cells', { r: 0, c: 0 });
    const compare = (state, label) => require(json(snapshot()) === json(state), label);
    try {
      app.switchSheet(si); atOrigin();
      pointSets = w.sheets.slice(0, w.ownSheetCount()).map(pointSet);
      // Freeze only the bounded audit fingerprint, not the workbook or its cells.
      baseline = JSON.parse(json(snapshot())); baselineText = json(baseline);
      out.beforeFingerprint = digest(baselineText);
      out.sampleCells = pointSets.reduce((n, p) => n + p.length, 0);
      out.sampleFormulas = baseline.sheets.reduce((n, s) => n + s.cells.filter(c => c.formula).length, 0);
      out.objects = baseline.sheets[si].objects.length;
      out.pivots = baseline.sheets.reduce((n, s) => n + s.pivots.length, 0);
      // Keep the caller's undo/redo history intact and keep probe transactions separate.
      w.undoStack = []; w.redoStack = [];
      for (const axis of ['row', 'col']) {
        const insert = axis === 'row' ? 'insertRows' : 'insertCols', remove = axis === 'row' ? 'deleteRows' : 'deleteCols';
        atOrigin(axis); run(insert, 1);
        const dr = axis === 'row' ? 1 : 0, dc = axis === 'col' ? 1 : 0;
        let literals = 0, formulas = 0;
        for (const before of baseline.sheets[si].cells) {
          const after = cell(si, before.r + dr, before.c + dc);
          if (before.formula) { require(after.formula, `${axis}: 수식 셀 이동`); formulas++; }
          else { require(json(after.raw) === json(before.raw), `${axis}: 값 셀 이동`); literals++; }
          require(json(after.style) === json(before.style), `${axis}: 셀 서식 이동`);
        }
        const sizeKey = axis === 'row' ? 'rowSizes' : 'colSizes';
        const sizeAt = n => axis === 'row' ? w.rowHeight(si, n) : w.colWidth(si, n);
        require(baseline.sheets[si][sizeKey].every((n, i) => sizeAt(i + 1) === n), `${axis}: 행·열 크기 이동`);
        const inserted = JSON.parse(json(snapshot()));
        const shifted = inserted.sheets[si].objects, delta = sizeAt(0), coord = axis === 'row' ? 'y' : 'x';
        require(baseline.sheets[si].objects.every((o, i) => {
          const a = shifted[i]; if (!a || a.id !== o.id) return false;
          if (!Number.isFinite(o[coord]) || !Number.isFinite(a[coord])) return true;
          const placement = o.placement ?? (o.prop === 'slicers' ? 'oneCell' : 'twoCell');
          return Math.abs(a[coord] - o[coord] - (placement === 'absolute' ? 0 : delta)) <= 1;
        }), `${axis}: 배치 개체 앵커 이동`);
        // A separate observation, rather than a claim of complete pivot layout parity.
        const pivotCoordinate = axis === 'row' ? 'top' : 'left';
        out.observations.push({ axis, pivotPositionMoved: baseline.sheets[si].pivots.map((p, i) => (inserted.sheets[si].pivots[i]?.[pivotCoordinate] ?? 0) === (p[pivotCoordinate] ?? 0) + 1), movedLiterals: literals, movedFormulas: formulas });
        atOrigin(axis); run(remove, 2); // Delete only the newly inserted empty row/column.
        const deleted = JSON.parse(json(snapshot()));
        run('undo', 1); compare(inserted, `${axis}: 삭제 Undo 복원`);
        run('undo', 0); compare(baseline, `${axis}: 삽입 Undo 원상복원`);
        run('redo', 1); compare(inserted, `${axis}: 삽입 Redo 복원`);
        run('redo', 2); compare(deleted, `${axis}: 삭제 Redo 복원`);
        run('undo', 1); compare(inserted, `${axis}: 삭제 재실행 Undo 복원`);
        run('undo', 0); compare(baseline, `${axis}: 최종 원상복원`);
        w.redoStack = [];
      }
      out.ok = true;
    } catch (e) {
      // Labels contain no source cell values or private file names.
      out.error = out.checks.at(-1)?.ok === false ? out.checks.at(-1).name : '구조 검사 도중 예외';
      out.errorType = e?.name ?? 'Error';
    } finally {
      if (baseline && app.wb() === w) {
        for (let attempts = 0; w.undoStack.length && w.undoStack !== oldUndo && attempts < 12; attempts++) {
          const n = w.undoStack.length; app.run('undo'); if (w.undoStack.length >= n) break;
        }
        out.restored = json(snapshot()) === baselineText;
        out.afterFingerprint = digest(json(snapshot()));
        if (out.restored) { w.undoStack = oldUndo; w.redoStack = oldRedo; }
      }
      if (app.wb() === w && out.restored) {
        app.switchSheet(originalUi.si); app.selectRange(originalUi.sel, 'cells', originalUi.active);
        app.gv().setScroll(originalUi.sx, originalUi.sy);
      }
      out.ok = out.ok && out.restored;
    }
    return out;
  }, { id });
  if (!report.ok || !report.restored) {
    const error = new Error('구조 검사 또는 원상복원 검증 실패: 내보내기를 중단합니다.');
    error.audit = report;
    throw error;
  }
  return report;
}
