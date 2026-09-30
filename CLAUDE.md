# CLAUDE.md

## Project Overview

**네이버 파워링크 광고 순위 체크 프로그램** (Naver Power Link Ad Rank Checker)

A Python CLI tool that checks the ranking position of a specific business in Naver's paid ad search results (파워링크) for a list of keywords. It checks both PC and mobile ad search pages separately.

## Repository Structure

```
revinchu/
├── CLAUDE.md                    # This file - AI assistant guide
├── requirements.txt             # Python dependencies
├── keywords.txt                 # Sample keyword list (one per line)
├── naver_ad_rank_checker.py     # Main rank checking program
├── debug_html_structure.py      # HTML structure analysis utility
├── .github/workflows/tabula-pages.yml  # GitHub Pages deploy for Tabula
└── tabula/                      # Excel-style web spreadsheet (separate app, see below)
```

## Tech Stack

- **Language**: Python 3.11+
- **Dependencies**: requests, beautifulsoup4, lxml
- **No framework** - standalone CLI scripts

## Setup

```bash
pip install -r requirements.txt
```

## Key Commands

### Run rank checker
```bash
# With keyword file
python naver_ad_rank_checker.py -k keywords.txt -t 하늘마음

# Single keyword
python naver_ad_rank_checker.py -q 부산가려움증 -t 하늘마음

# Export to CSV
python naver_ad_rank_checker.py -k keywords.txt -t 하늘마음 -o result.csv

# Debug mode (saves raw HTML to debug_html/ folder)
python naver_ad_rank_checker.py -q 부산가려움증 -t 하늘마음 --debug
```

### Analyze HTML structure (debugging)
```bash
python debug_html_structure.py 부산가려움증
```

## Architecture

### URL Patterns

| Platform | URL Template |
|----------|-------------|
| PC | `https://ad.search.naver.com/search.naver?where=ad&query={keyword}` |
| Mobile | `https://m.ad.search.naver.com/search.naver?where=m_expd&query={keyword}` |

### HTML Parsing Strategy

The parser uses multiple fallback strategies because Naver may change HTML structure:

1. **CSS selector-based**: tries known selectors (`ul.lst_type li`, `a.lnk_tit`, `.tit_area a`, etc.)
2. **data-attribute based**: looks for `[data-cr-area]` containers
3. **Heading-based fallback**: extracts business names from `h2/h3/h4` tags

If parsing fails, use `--debug` flag or `debug_html_structure.py` to inspect actual HTML and update selectors.

### Business Name Matching

Uses **substring matching** - searching for "하늘마음" will match titles like "하늘마음한의원 덕천" or "하늘마음한의원 서면". The match is case-insensitive and bidirectional (target in title OR title in target).

## Conventions

- All user-facing output is in Korean
- CSV files use `utf-8-sig` encoding for Excel compatibility
- Request delay between PC and mobile calls defaults to 1 second (configurable via `--delay`)
- Debug HTML files go into `debug_html/` directory (gitignored)
- Keywords file uses `#` for comments and blank lines are skipped

## Important Notes for AI Assistants

- **Naver HTML changes frequently**: If parsing stops working, run `debug_html_structure.py` to analyze current HTML structure and update selectors in `parse_pc_ads()` / `parse_mobile_ads()`
- **Rate limiting**: Keep delay between requests ≥ 1s to avoid being blocked by Naver
- **User-Agent strings**: Must be kept up-to-date; Naver may block outdated UA strings
- **Proxy**: Do not set `http_proxy`/`https_proxy` unless explicitly required; the script accesses Naver directly
- **Encoding**: Naver pages use UTF-8; CSV export uses UTF-8 with BOM for Korean Excel compatibility
- **No tests yet**: This is a scraping tool whose correctness depends on live page structure

## Tabula (Excel-style spreadsheet, `tabula/`)

A standalone, dependency-free browser spreadsheet that mimics Microsoft Excel's UI (Korean). Unrelated to the rank checker.

- Run: `cd tabula && npm start` → http://localhost:5178 ; test: `npm test` (Node built-in test runner)
- Web build: `npm run build` → `tabula/dist/index.html` (single file; `build.mjs` inlines CSS and wraps each ES module, so only `import {…} from './x.js'` and `export const|let|function|class` forms are allowed — no default/`*`/dynamic imports, no import cycles). `.github/workflows/tabula-pages.yml` deploys it to GitHub Pages.
- Pure ES modules, no build step, no npm dependencies. Keep it that way (zip/inflate/XML for .xlsx are hand-written in `src/zip.js`, `src/xml.js`).
- DOM-free modules (unit-tested in Node): `cellmap.js`, `depgraph.js`, `outline.js`, `audit.js`, `protect.js`, `sparkline.js`, `page.js`, `cube.js`, `block.js`, `formula.js`, `fxcore.js`, `fx-*.js`, `xlfn.js`, `flashfill.js`, `workbook.js`, `format.js`, `axis.js`, `xlsx.js`, `zip.js`, `xml.js`, `chart.js`, `pivot.js`, `series.js`, `csv.js`, `validation.js`, `shapes.js`, `vba.js`, `tables.js`, `textsplit.js`, `fmtpresets.js`, `condfmt.js`, `slicerstyle.js`, `stylepresets.js`, `xlsb.js`, `xls.js`, `emf.js`, `ods.js`, `find.js`, `analysis.js`, `ets.js`, `library.js` (IndexedDB).
- Formula engine: `formula.js` = tokenizer/parser/evaluator; functions live in `fx-math/stat/text/date/lookup/logic/fin.js` (import only `fxcore.js`, never `formula.js` — no cycles). Values: scalars, `Range` (2-D array, `.ref` when from cells), `RefValue` (unresolved reference; `evalAny` keeps refs, `evaluate` derefs), `Lambda`. Function flags: `lazy(fn)` gets AST args + `(ctx, ev)`; `refFn(fn)` gets RefValue args; `lift(fn, positions)` maps scalar functions over arrays and sets `fn.liftPos` (used by `mayReturnArray` to find spill candidates). Normal functions receive error values as arguments (not thrown).
- Dynamic arrays: `evaluateArray` returns a Range for multi-cell results; `Workbook.placeSpill` spills into empty cells (`#SPILL!` if blocked), `ensureSpills()` evaluates all `cell.maybeArray` formulas before reading empty cells. xlsx writes them as `t="array"` + `cm="1"` + `xl/metadata.xml`; `xlfn.js` converts app formulas ↔ file form (`_xlfn.`, `_xlpm.`, `_xleta.`, `SINGLE`, `ANCHORARRAY`, legacy implicit intersection `@`).
- Defined names: `wb.names = [{ name, ref: '=…', sheet|null, comment, hidden }]` (change via `wb.setNames`); evaluated by `wb.nameValue` (range → RefValue, constant, or LAMBDA). Cells may carry `cached` (file value for unsupported functions) and `link` (hyperlink URL or `#Sheet!A1`).
- Pivot defs use field names (also `styleDef` = custom style parts from the file's `<tableStyles>`, `errorCaption` = Excel's showError/errorCaption): `{ name, rows, cols, values: [{ field, agg, name, showAs, numFmt… }], pages, filters, calcFields: [{name, formula}], sort: {field: {dir, by}}, order: {field: [items]}, fieldFilters: {field: top|label|value}, style, styleOpts, rowCaption, cellFmt, captureFmt, layout, subtotals, grandRows, grandCols, top, left, area }` (old `rowField/colField/valueField` defs are still read via `normalizeDef`). Calculated fields are evaluated on field sums (`makeMeasures`). `computePivot` cells carry a `role` (rowHead, valueHead:i, rowItem:d, data:i, grandData…) used for pivot styles (`pivotStyleParts`/`roleStyle`) and for re-applying formats captured from the file (`cellFmt`). A sheet has `pivot` and `pivotsExtra`; in app.js use `allPivots`/`findPivotEntry`/`putPivotDef` and refer to pivots by `{sheet, name}`. Source reads are cached per sheet version (`wb.sheetVersion(si)`).
- Sheet size is `MAX_ROWS` = 20,000,000 (more than Excel's `EXCEL_MAX_ROWS` = 1,048,576); xlsx export drops rows beyond Excel's limit and `xlsxOverflow()` reports how many.
- Charts: `{ type ('combo' allowed), series: [{name, cat, val}], seriesFmt: [{type, axis: 'secondary', color, labels, marker, numFmt}], labels, legend, grouping, axes, pivot: {sheet, name} }`; `chartModelData` (chart.js) resolves ranges or pivot data (`pivotChartData`) — use it for rendering and export.
- In-cell pictures: cell `image: {src, alt}` with `raw: ''` (value is a `CellImage` from fxcore; `IMAGE()` returns one too). Keep `image` when rewriting a cell's style/comment. xlsx: richData (`metadata.xml` valueMetadata XLRICHVALUE + `xl/richData/*`, cells `t="e" vm="N"`).
- Recalc model: formula results are cached per sheet (`wb.caches[si]`). Imported sheets have `fileValues: true` and formula cells keep Excel's saved result in `cell.cached`; `getValue` returns it until that sheet (or a sheet it references) changes. `wb.invalidate(si)` clears only `si` and its dependents (`wb.affected(si)`, from the static `sheetDeps()` map built from formula ASTs: sheet refs, table refs, names → all, INDIRECT → all); `wb.invalidate()` with no argument is a full recalc (F9) and drops all file values — avoid it. Structural changes use `invalidateStructure()` (keeps file values). Formula ASTs are shared per formula text (`parseMemo`) and must never be mutated.
- Undo: cell edits record `cell` entries; sheet props use `propSnap` (per property); `CALC_NEUTRAL` props don't recalc. Row/column insert uses `snapshotSheet(si)` (that sheet only) + recorded cells in other sheets; add/delete/move sheet uses `snapshotList()` (sheet objects are kept, so `putSheet` restores contents in place); rename records a `rename` entry; names use `snapshotNames()`. `snapshotAll()` serializes the whole workbook — don't use it in new code (too slow for big files).
- Big files: `readXlsxAsync` / `writeXlsxAsync` / `wb.loadAsync` run step generators and yield to the browser (progress via `progressOverlay` in app.js); the sync `readXlsx` / `writeXlsx` drive the same generators (tests use them). Sheet XML is read by the DOM-free `scanRows` scanner (not `parseXml`); `readXlsx` returns `sheet.cells` as a `Map`. `zipAsync` compresses with the browser's `CompressionStream`. Workbooks over 50k cells autosave to IndexedDB per sheet, only changed sheets (`sheet._ev` edit counter, `_sid` id), as gzip Blob chunks (`saveBigToIdb` / `loadBigFromIdb`); over 300k cells the server is saved only on explicit [저장].
- Per-sheet `noGrid` (gridlines hidden, xlsx `showGridLines="0"`) is a sheet prop.
- Cell storage: `sheet.cells` is a `CellMap` (`cellmap.js`): same API as `Map('r,c' → cell)` but stored per column as `Map(row → cell)`. Hot paths must use `getRC/setRC/deleteRC/forEachRC/col(c)` (string-key lookups on 1M-entry maps are ~8× slower). New sheets/readers create `new CellMap()`; plain Maps passed to `restore` are converted.
- Calc engine: formulas with the same R1C1 shape share one AST (`parseAt`/`shareKey`); a cell may carry `dr`/`dc` (offset from the AST's anchor) and `evalAny` shifts relative refs by `ctx.dr/ctx.dc` — never read `cell.ast` refs without applying them (use `DepGraph.refBoxes` / `wb.precedentsOf`). Cell edits go through `wb.changed()` → `dirtyPoints()`: the `DepGraph` (`depgraph.js`, runs of filled-down formulas, verified lazily) finds dependents and clears only their cache entries (`cell.dirty` stops using the file's cached value); dynamic formulas (INDIRECT/OFFSET/RAND/TODAY/spill refs/relative names) are always dirtied; spills and >100k-point batches fall back to sheet-level `invalidate(si)`. Inside `transact` changes are queued in `wb.pending` and flushed on the next read or at the end. Any code that replaces `sheet.cells` or moves formulas must call `invalidateStructure()` (drops the graph). The graph is prebuilt in the background after load (`wb.prepareGraph()`). Big range reads use `rangeFast` (block columns + per-column row lists, cached per column version `bumpCol`) — don't mutate `Range.rows` from functions. `formulaShifter(raw)` generates filled formula texts without re-tokenizing.
- Sheet props added for Excel features: `outline` (`outline.js`: `{rows, cols, rowsColl, colsColl, below, right}`, collapse = hiddenRows/hiddenCols), `protect` (`protect.js`: `{on, hash, allow}`; cell styles `locked: false`, `hideFormula`; app guard `protectBlocked(protectAction(cmd))` in `run()`), `sparklines` (`sparkline.js`: groups with `items: [{r, c, ref}]`), `page` (`page.js`: orientation/paper/margins/scale/fit/header/footer/area/titleRows). All are adjusted in `shiftAxis` and written to/read from xlsx.
- Big data: large CSV/xlsx data is stored as column blocks (`block.js`, `sheet.blocks = [{ r0, c0, n, ver, dver, cols: [{ num: Float64Array|null, str: Int32Array|null, dict, fmt }], perm? }]`) instead of cell objects. `wb.getCell`/`getValue`/`styleAt` read blocks; `putCell` writes plain values into the block (formulas/other styles go to the Map and blank the block cell); `perm` is a row-order view used by sort (`sortOrder`/radix, `reorderRows`). Hidden rows can be bitmaps (`axis.js` `{__bits}`); never materialize a block per row in hot paths.
- Pivot/slicer computation goes through `cube.js` (columnar: `Column.dim()` codes/keys, `filterRows`, `groupAggregate`, `groupedColumn` for date/number grouping, rollup cache `aggregateQuery` for ≥200k rows). `pivotSourceData` returns `{ cube, si, ref, table }` with a lazy non-enumerable `rows` — don't spread results that carry the `rows` getter. `resolvePivot` → `{ def, header, cube, filters, groups, measures }`; `computePivot(res, def)` builds trees from groups.
- Pivot def extras: `values[i].showAs` (+ `baseField`, `baseItem` | `basePos: 'prev'|'next'`), `collapsed: { field: [item text] }` (xlsx `sd="0"`), `groups: { field: { by: 'months'|'number', start, size } }`, `subtotalTop`, `blankRows`, `repeatLabels`, `missingCaption`, `errorCaption`, `showExpand`. `pivotDetail(src, def, gridRow, gridCol)` returns the source rows of a value cell (drill-down). Toggle (+/−) buttons are `def.buttons` entries with `kind: 'toggle'`.
- Drawing objects live in sheet props `charts`, `images` (data-URL `src`), `shapes`; each has `id`, `x/y/w/h` (sheet px) and optional `z` (stacking). `app.js` tracks the selected object id in `chartSel` for all three kinds.
- Macros: `wb.vba = { bin (base64 vbaProject.bin), codeName, sheetCodes }` is preserved and written back as .xlsm; VBA is displayed, never executed.
- Number formats: builtin `style.numFmt` ids, or `numFmt: 'custom'` + `style.code` (Excel format code rendered by `formatCode()` in `format.js`). `styleForCode()` maps a code back to a builtin id when it renders identically; use it for anything that sets a format from a code (dialog, xlsx import).
- Tables: `sheet.tables` (`tables.js`): `{ id, name, r1..c2, header, totals, style, banded…, filter: {criteria, hidden, sort}|null, totalsFns }`; r1 = header row, r2 = totals row when present. Each table has its own filter; code that reads filters must go through `getFilter(key)`/`putFilter(key, f)` in `app.js` (key '' = sheet autofilter, else table id), and hidden rows are the union of `hiddenRows`, `filter.hidden` and every table's `filter.hidden`. Structured refs (`표1[열]`, `[@열]`) are `sref` tokens resolved via `ctx.structRef` (needs the formula cell, see `ctxFor(si, r, c)`); they are written to xlsx in canonical English form (`canonicalRef`).
- Slicers: drawing objects in `sheet.slicers` with `source: {kind:'table', table, column} | {kind:'pivot', field, pivots: [{sheet, name}]}` (legacy `self`/`sheet` still read; one slicer can drive several pivots — `slicerPivotTargets`); style is `sl.style` (`slicerstyle.js`) + optional `sl.custom` colors; their selection is not stored on the slicer but derived from the table filter criteria / `pivot.filters` (see `slicerModel`). In xlsx they are written as native Excel slicers (slicerCache + slicers parts, drawing `mc:AlternateContent`, workbook/sheet `extLst` with the x14/x15 GUIDs) and read back in `linkPivotsAndSlicers`.
- Pivot tables: source rows and filtering live in DOM-free `pivot.js` (`pivotSourceData`, `resolvePivot`, `computePivot`); `writePivot` in app.js renders the grid into `def.area`. xlsx export writes a real pivotCacheDefinition (saveData="0", refreshOnLoad="1") + pivotTableDefinition with all row/col/page/data fields; import rebuilds the full def.
- Excel fidelity: built-in table/pivot styles come from `stylepresets.js` (`PRESET_DXF`/`PRESET_STYLES`, generated from Excel's presetTableStyles; element i uses dxf `count-1-i`). `tablePresetCell` and `paintPivotPreset` apply elements in Excel's precedence, and draw outer (L/R/T/B) and inner (V/H) borders per region. Colors use `applyTint` (Windows integer HLS), and `setThemeColors(wb.theme)` sets the palette. `wb.baseStyle` is xlsx xf 0, used by `styleAt` for unstyled cells and as the start style in `setStyle`; it is written back as xf 0. `wb.fitRows` lists rows with no stored height, auto-fitted once on open (`fitRowsOnOpen`). Pivot `cellFmt` number formats replace (not merge with) the pivot's default decimals.
- Calculated fields: `checkCalc` (validation), `renameCalcRefs`, `excelCalcFormula` (DIVIDE/ROWS → Excel form; the original goes in the ignorable `tb:formula` attribute of the pivot cache), and `CALC_FUNCS` live in `pivot.js`. `ROWS()` uses measure slot -1 (row count). In `app.js`, `saveCalcFields` updates every pivot on the same source (`sameSourcePivots`, like Excel's shared pivot cache). The manager dialog is `calcFieldDialog(entry, name)`, and `listCalcFormulas` is Excel's List Formulas.
- Loading and recalculation performance: write pivot output inside `wb.transact` (one dependency propagation per pivot, not per cell). `DepGraph.propagate` tracks per-run coverage, so a filled-down fixed-range run is walked once. `COUNTIF` uses `fastCount` (per-range hash and sorted index, compares at 15 significant digits). Number→text conversion uses `numberText` (15 significant digits, like Excel).
- Chart series may reference defined names (`{ name, sheet? }`, file form `[0]!name`); `chartModelData` evaluates them on every render, so OFFSET-based ranges follow slicers.
- Big autosave: column blocks are written to IndexedDB in 16MB parts (`PART_BYTES`, record `partKeys`, `numParts/strParts/permParts`) and joined on load. One large structured clone would double memory and crash the tab.
- Conditional formatting: `condfmt.js` evaluates rules (`prepareCond`, `condFormatAt`). `sheet.cond` order is priority: index 0 wins over later rules and `stopIfTrue` stops later rules; `addCondRule` inserts at the front like Excel. xlsx `priority` maps to this order.
- Ribbon: contextual tabs use `context: 'table' | 'slicer'` in `TABS` and are shown from `ribbonState().context`. Alt key sequences (KeyTips, e.g. Alt+A+E) are defined in `KEYTIPS` in `app.js`.
- The grid is virtualized (`src/view.js`): only visible rows/cols are rendered; positions come from `Axis` (default size + sparse custom sizes/hidden). Never loop over all 1,048,576 rows — clamp to `wb.usedRange()` / `wb.extent()`.
- Data analysis (Excel add-ins): `analysis.js` = Analysis ToolPak tools (descriptive, correlation/covariance, regression, histogram, t/z/F tests, ANOVA, moving average / exponential smoothing as formulas, random numbers, sampling) returning `{ rows, heads, pct }` tables, plus `solveMin` (Solver: penalty + Nelder–Mead + coordinate polish, integer vars by rounding). `ets.js` = FORECAST.ETS AAA model (`fitEts`, `timeAxis` recognizes monthly date steps); fx-stat.js imports it (ets.js has no imports, so no cycle). App side: `dataAnalysisDialog`/`toolDialog` (modeless, `refInput` range fields filled by sheet selection via `refPick`), `forecastSheetDialog` (new sheet with FORECAST.ETS formulas + table + chart), `scenarioManager` (sheet prop `scenarios: [{name, comment, locked, cells: [{r,c}], values}]`, xlsx `<scenarios>`), `solverDialog`.
- Regex lookups: XLOOKUP/XMATCH `match_mode` 3 = regex (Excel 365); WIXEL extensions REGEXVLOOKUP, REGEXHLOOKUP, REGEXMATCHPOS, REGEXSEARCH (`lookupRegex` in fx-lookup.js: partial match, case-sensitive unless case_sensitivity 1).
- Browser library (`library.js`): recent documents (default 30, `opts.libMax`) as gzip JSON snapshots in IndexedDB keys `lib:index` / `lib:doc:<id>` / `lib:ver:<id>:<ts>` with up to 20 versions per doc (`VER_MAX`). app.js keeps `docId` (persisted in the autosave meta), calls `libraryFlush()` before any workbook is replaced (loadWorkbook*), on edits (debounced) and on explicit save (named version). Docs over `LIB_CELL_LIMIT` cells use only the big-doc autosave.
- Publish / view mode: `publishDialog` makes a read-only link (`#view=` + base64url(gzip(snapshot)), no server) or a server link (`?view=<id>`, server.js `/api/publish`, `/api/published/:id`, auto re-publish, viewers poll every 30s). `viewOnly` blocks `run()` commands that are not 'free'/`VIEW_CMDS`, `startEdit`, and all saving. `?doc=<name>` opens a server doc; `startCollabWatch` reloads it when someone else saved and there are no local changes.
- Pivot conditional formatting scope: rules may carry `pivot: { name, scope: 'selection'|'data'|'field', value, rowField, colField }`; `writePivot` records `pivotLayouts` (cell roles per pivot) and `refreshPivotCond` recomputes those rules' ranges on every redraw. xlsx: `<conditionalFormats>` in the pivot definition + `pivot="1"`, linked back by priority on import (`linkPivotCond`).
- Pivot charts may show a subset of metrics: `chart.pivot.values = [value display names]` (filtered in `resolveChart`). Exported to xlsx as a normal range chart over the pivot cells, with `wxPivot` in the WIXEL chart extension to restore the link.
- Timeline slicers are slicers with `timeline: true, level: 'Y'|'Q'|'M'|'D'` (rendered by `timelineHtml`, periods from `timelinePeriods` in view.js, filtering via the normal slicer model `apply`).
- Pivot areas block only cell-changing commands (`PIVOT_LOCKED`), not inserts of charts/slicers.
- GETPIVOTDATA is not a dynamic function: `depgraph.js` makes it depend on the column span of the pivot that contains its reference cell (whole sheet if none). `writePivot` calls `wb.dropGraph()` when a pivot's column span changes. `propagate(points, limit)` returns null above `DIRTY_LIMIT` (workbook.js) and `dirtyPoints` falls back to sheet-level `invalidate`. `wb.pivotMemo` caches the resolved pivot per def while `wb.version` is unchanged, and `pivotLookup` keeps a per-`res.groups` index (item texts → merged list).
- `setCellData` with the same value (numbers compared at 15 significant digits) and only a new style does not dirty dependents. While opening a file, pivot redraws run with `wb.holdDirty = true` (or `holdDirtyWhile`) and listeners muted, and changed cells are flushed once.
- xlsx read memoizes formula conversion by "shape" (`formulaShape`: row numbers replaced) when the conversion is the identity. Shared formulas use `formulaShifter`. The export `exportF` does the same.
- Pivot grouping: `def.groups[field] = { by: 'years'|'quarters'|'months'|'mdays'|'number'|…, start?, end?, base? }`. `mdays` is Excel's day group ("12월9일", no year). A group with `base` is a derived field (Excel's '월2' = months of '일'), exposed by `withDerivedFields(cube, groups)` in pivot.js. Filters on grouped fields are translated to raw values in `resolvePivot`. xlsx: read from `fieldGroup`/`rangePr`/`groupItems` (pivot field items then index groupItems). Written back by `excelGroupItems`: an in-place group is a `fieldGroup` on the base cache field; a derived group is an extra `databaseField="0"` cache field. Cache date items `<d>` are read as serials (`isoSerial`).
- Pivot `tieOrder` (from the file's rowItems/colItems) breaks ties in value sorts and Top-N filters like Excel's saved order. The Top-N count mode also keeps items tied with the Nth value. Text labels that look like numbers, or start with `'` or `=`, are written with a leading `'`.
- Error values should be returned, not thrown, on hot paths: throw/catch costs ~3.5 µs in Chrome, which made `IFERROR(x/0, 0)` over 80k cells take seconds. `binary()` returns `ERR.DIV0` and error operands, and `derefSoft` (formula.js) is used for function arguments, `ev.value` and `finish`. `deref` still throws for internal propagation.
- Cells: a formula stored in a text-formatted cell in a file keeps `fx: true` (a formula, like Excel). Typing into a text-formatted cell still gives text. Empty-text cells (`""` in xlsx) are raw `'` (COUNTA / pivot Count count them, and the cube's `ne()` treats `''` as non-empty).
- Pivot Top-N after opening a file: `def.tieState` (the filter key at save time) plus `tieOrder` reproduce exactly the items Excel saved while the filters are unchanged. `writePivot` deletes `tieState` on the first redraw that is not part of opening (`openingPivots`). Computed Top-N treats error/blank values as ties, so all of them show when fewer than N items are numeric. GETPIVOTDATA returns #REF! for items hidden under a collapsed parent.
- xlsx export: `xfAt` memoizes the composed style's xf id per (col style, row style, cell style) object triple. `exportF` caches "file form = app form" per shared AST. The browser autosave pauses while an export runs (`exportBusy`).
- Never use `Math.max(...arr)` / `push(...arr)` on data-sized arrays (stack overflow above ~100k items); use `maxOf` / `minOf` / `pushAll` from fxcore.js.
- All UI text is Korean. New functions need an entry in `src/funcinfo.js` (a test enforces this).
- Undoable changes go through `wb.transact(fn, meta())`; sheet-level props (freeze, filter, charts, merges, hidden rows…) use `wb.setSheetProp()` / methods that call `wb.snapshotAll()`. New sheet props must be added to `SHEET_PROPS` in `workbook.js` and adjusted in `shiftAxis`.
- `server.js` stores documents in `tabula/data/` (gitignored) via `/api/files`; the app falls back to localStorage when the API is absent.
- Legacy .xls (BIFF8): `xls.js` `readXls(bytes)` (CFB container → records → cells/styles/names/merges/pictures); formulas are decoded from ptg tokens to file-form text and go through `fromFileFormula(text, { legacy: true })`. app.js picks it by the OLE signature (D0 CF 11 E0). Saving always writes xlsx.
- Pictures: `im.crop = {l, t, r, b}` (fractions, xlsx `a:srcRect` / xls blip crop). EMF pictures are converted to SVG by `emf.js` (`emfToSvg`, GDI records only) for display; the original is kept in `im.emf` (data URL) and written back as `.emf` on export.
- Legacy implicit intersection: `xlfn.js` `SCALAR_ARGS` lists scalar arguments of non-lifted functions (SUMIF criteria, VLOOKUP lookup_value…) that get `@` from legacy files. Wildcard criteria: `~` escapes only `* ? ~`.
- Criteria functions are lifted over their criteria argument (`SUMIF/COUNTIF/AVERAGEIF` position 1, `*IFS` criteria positions): `COUNTIF(A1:A9, A1:A9)` returns an array (used by `SUMPRODUCT(1/COUNTIF(…))`).
- Charts: a series name that points at several cells joins the non-empty ones with a space. A multi-column category range becomes Excel's multi-level axis (`multiLevel` in chart.js → `data.catLevels`, drawn under the inner labels; xlsx `multiLvlStrRef`). When `seriesFmt[i].type` is set, a missing `axis` means the primary axis (the combo default "last series on secondary" applies only without a format).
- Cell rendering: a cell with a fill or pattern uses `background-clip:border-box` and covers its own right/bottom gridline like Excel; when it has no own top/left border its box starts 1px inside, so the neighbour's line is never painted over (redrawing it would double the line at non-100% zoom). Overflowing text (`.c.ovf`) sits above the empty neighbours' fill. Pattern fills are 8×8 pixel bitmaps (`PATTERN_BITS` in view.js), not gradients. Data bars are empty when every value is 0.
- Pivot redraw speed: `cleanStyle` returns the style object itself when it has no null/undefined values (styles are replaced, never mutated). `setCellData` compares cells with `sameCell` (no `cellData`/JSON). While opening a file `wb.noUndo` is set, so cell records carry no before/after copies. `writePivot` remembers the written style objects in a `CellMap` (`pivotWritten`) and compares by identity.
- xlsx fonts: black is the default colour, except an explicit RGB black (`rgb="FF000000"`), which is kept as direct formatting (it beats a table style's white header text, like Excel).
- A ZIP without an end-of-central-directory record but with a local header signature is reported as a truncated download/upload.
- Rendering: numbers never wrap; when the file's font is missing (`fontMissing` in view.js) a number up to 30% too wide is drawn smaller instead of `###`. Drawing anchor offsets are clamped to the cell size like Excel.
- Two collation orders: formula comparison / MATCH / LOOKUP binary search use `compareText` (fxcore: symbols < digits < Hangul < Latin, Korean Excel). Sorting (data sort, pivot items, filter lists, slicers) uses `compareSortValues` / the `Intl.Collator('en', {sensitivity:'accent'})` collators (digits < Latin < Hangul, case-insensitive). Pivot item sort puts Excel's built-in custom lists first (`CUSTOM_LISTS` in cube.js: 1월…12월, 요일, Jan…Dec, 분기, 간지).
- Pivot captions from the file: `def.fieldCaptions` (pivotField@name), `def.itemCaptions[field][item]` (item@n), `def.grandCaption` (grandTotalCaption); `subtotals` and `blankRows` may be `true`, `false` or an array of field names (per-field). OOXML `_xHHHH_` escapes are decoded by `unx` (xml.js).
- Pivot cache snapshots: Excel shows pivotCacheRecords until refresh. `wb.pivotSnapshots` (Map cachePath → {rows, ver}) is used by `pivotSourceData` while the source sheet's `_ev` is unchanged; Refresh sets it to null. The blank gap row between page filters and the table is not written by `writePivot` (users put titles there).
- Pivot items: empty text `""` is its own item key `EMPTY_TEXT` (cube.js), shown as an empty label and sorted first among texts; blank cells are `EMPTY` '(비어 있음)'. A page field shows an item name only when exactly one selected item exists in the source (items missing from the data are not counted), else '(다중 항목)'. The grand total of an empty pivot evaluates calculated fields on zero sums (`measures.emptyValue`); plain sums stay blank. Numeric error/missing captions are written as numbers (`captionRaw`); GETPIVOTDATA returns the error caption ('' → 0). Tabular/outline layouts show column field names (and '값') in the first header row; compact shows '열 레이블'.
- Off-axis hidden items (fields not on rows/cols/pages) are ignored like Excel unless a slicer on that field drives the pivot (`dropOffAxisFilters` after `linkPivotsAndSlicers`). `Sheet!#REF!` in formulas is an error token.
- Dates follow Excel's 1900 system: `dateParts`/`serialOf` (format.js) map serial 0 → 1900-01-00, 60 → 1900-02-29, and serials below 61 one day earlier than the JS epoch; weekday is `(serial + 6) % 7`. Use `serialToDate`/`dateToSerial`, never raw `EPOCH` arithmetic for day parts.
- Empty arguments: VLOOKUP/HLOOKUP `range_lookup` and MATCH `match_type` given as `,)` mean exact match. Array IF (`pickEach`) returns 0 for blank cells.
- `*IFS` with only equality criteria (numbers, text without wildcards/operators, booleans, '') go through `fastIfs` (fx-math.js): one hash index per (target rows, criteria rows) set, keyed by `critKey`/`cellKey` (numeric text matches numbers), memoized on the Range `rows` objects.
- xlsb pivots: calculated field PName lists are per field (`f.pnames`, and formulas are not memoized across fields), row/col items (`PTROWITEMS`/`PTCOLITEMS`, SXLI `r:u16 t:u16 n:u32 i:u32` + x array) give Excel's tie order, and cache field groups (`PCDFGROUP` par/base, `PCDFGRANGE` by/flags/start/end/step, group items) become `<fieldGroup>`.
- Unopenable files get Korean messages: password-protected (EncryptionInfo / FILEPASS) and DRM-encrypted (Fasoo etc. signature in zip.js).
- xlsb pivot fields: a display name (pivotField@name) follows the 20 fixed bytes of PTFIELD when `f1 & 0x20000000`. Tie order is also kept per parent item path (`def.tieByParent[field][parent texts joined by \u0001]`, used by `tieRank(d, field, parentNode)`), because an inner item (검색어) can sit at different places under different outer items. Pivot sums over groups whose values are all blank are blank (not 0); GETPIVOTDATA returns 0 for them.
- Opening files: pivot cells written differently from the file are not propagated synchronously when there is no dependency graph yet; `loadWorkbookAsync` hands them to `prepareGraph()` (which reuses a running build via `wb.graphRun`) and dirties them afterwards. Pivot sources are read with `wb.rangeRead`. Formula texts without A1 refs skip `scanRefs` (`noRefText`). Books with more than 2MB of pictures count as `bigBook()` (IndexedDB autosave in parts, no blocking localStorage/serialize right after opening).
- Cell text overflow follows Excel: left-aligned text spills right into an empty cell, right-aligned spills left, centered spills both ways only when both neighbours are empty (`.c.ovf` + flex justify).
- External workbook references: formulas may use `[N]Sheet!A1` / `'[N]Sheet name'!A1` (tokenizer `SHEET` accepts a `[N]` prefix; `quoteSheetName` leaves `[N]Name` unquoted). On xlsx import, `readExternalLinks` turns each externalLink's cached `sheetDataSet` into a read-only `veryHidden` sheet named `[N]Sheet` with sheet prop `external: N`, appended after all real sheets (`wb.ownSheetCount()`; `addSheet`/`moveSheet` keep them last), and keeps the raw externalLink XML/rels in `wb.externals`, which the writer puts back (`externalReferences`, parts, rels, content types) while writing only the first `ownSheetCount()` sheets. `*IF(S)` functions return #VALUE! for ranges on these sheets (closed workbook, like Excel). The unhide dialog lists only `hidden` sheets.
- Formula details matched to Excel: legacy implicit intersection (`@`) works across sheets (row/column of the formula cell); criteria that are error values match cells holding the same error (not propagated); a range argument that is an error (`SUMIFS(#REF!, …)`) returns it; a circular reference returns the file's saved value when there is one (the cell that closes the loop keeps its saved value).
- Built-in time formats 18–21, 45–47, 32/33/55/56 map to their exact codes (`TIME_CODES` in xlsx.js; 21 = h:mm:ss 24-hour) instead of the generic AM/PM `time`. `표[#All]`-style single special items are written without inner brackets.
- Pivot extras: `def.showHeaders === false` (Excel's field headers off) blanks row/column field captions; a pivot with no row and no column fields has no label column and no 총합계 label; `sort[field].at = [[otherField, item]]` sorts by the values under that item of the other axis (autoSortScope item references); value sorts treat blank as 0; a page field whose selection covers every item in the source shows '(모두)'; subtotal labels of numeric (date) items use the source column format (`res.fieldStyle`, set in `writePivot`).
- Spill candidates: `ensureSpills(area)` — when called while the spill queue is already running (a formula reading a range), it evaluates only the remaining candidates that could spill into that area (same sheet, anchor at or above/left of it); re-running the whole queue nested made files with thousands of candidates hang. `mayReturnArray` treats XLOOKUP with a one-column/one-row return range and INDEX with a row (and column) argument as single-cell, so whole-column lookups are not spill candidates.
- XLOOKUP of a blank value matches the first blank cell of the lookup range (for whole-column references, the first row below the data), returning 0 for a blank result like Excel.
- Sheet XML over 48MB is read in 8MB slices cut at `</row>` (`chunkedSheet`, byte search + per-slice `TextDecoder`) instead of one giant string.
- Pivot `valuesOnRows` (xlsx `dataOnRows="1"`, `-2` in rowFields): with several values, each row item becomes one row per value (value names in the extra last label column headed '값'; compact layout indents them under the item), column leaves carry `vi = -1` and `valueCells(rpath, kind, vi)` picks the value from the row; grand rows are '전체 <값 이름>'. `rowItems` entries carry `vi`. `cubeGroups` memo keys include the row-field count, so pivots with the same field on rows vs. columns never share group layouts.
- Circular references: every cell on the loop (from the re-entered cell up the evaluation stack) keeps its saved file value, not just the cell that closes it. `compareValues` compares numbers at 15 significant digits (sorting, MATCH/LOOKUP, MAX-equality).
- Formula sharing covers whole-column (`B:B`, `$E:$E`) and whole-row (`3:3`) references: `scanRefs` emits `{cols|rows}` segments (only the column/row part shifts), so filled-down `SUMIFS(Raw!K:K, …, D6)` rows share one AST. On xlsx read, a formula shape whose conversion changes the text (e.g. `Sheet!#REF!`) stores the converted shape (`out`) and later rows fill their row numbers into it (`fillShape`, verified when stored).
- Korean locale formats: `[$-412]` / `[$-ko-KR]` sections render `ddd`/`dddd` as 금 / 금요일 and `mmm`/`mmmm` as '2월' (`sec.korean`); `[$-F800]` / `[$-F400]` (system long date/time) render as Korean Windows does (`2021년 2월 1일 월요일`, `오후 2:24:00`). xlsx date helpers (`numberRaw`, `isoSerial`, `serialIso`) go through `dateParts`/`serialOf`, so dates before 1900-03-01 keep their serial (60 stays a number).
- Reference semantics for SUM-family functions (`REF_AS_RANGE` in formula.js: SUM, AVERAGE, COUNT, MAX, MIN, STDEV…): a single-cell reference argument is passed as a 1×1 Range, so text/logical values in referenced cells are ignored (`SUM(A1)` with text = 0) while direct text arguments still give #VALUE!.
- The spill queue skips cells that are currently being evaluated (`evaluating`): an `=INDIRECT(…)` array anchor reading its own source range must not re-enter itself (that looked like a circular reference and returned the saved scalar).
- Closed external workbooks: when a formula's own `*IF(S)` reads an external (`[N]Sheet`) range (`CLOSED_BOOK.hit`, fxcore), `evalCell` keeps the file's saved value if there is one (Excel shows it until a full recalc); without a saved value the result is #VALUE!.
