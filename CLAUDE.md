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
- DOM-free modules (unit-tested in Node): `formula.js`, `fxcore.js`, `fx-*.js`, `xlfn.js`, `flashfill.js`, `workbook.js`, `format.js`, `axis.js`, `xlsx.js`, `zip.js`, `xml.js`, `chart.js`, `pivot.js`, `series.js`, `csv.js`, `validation.js`, `shapes.js`, `vba.js`, `tables.js`, `textsplit.js`, `fmtpresets.js`, `condfmt.js`, `slicerstyle.js`.
- Formula engine: `formula.js` = tokenizer/parser/evaluator; functions live in `fx-math/stat/text/date/lookup/logic/fin.js` (import only `fxcore.js`, never `formula.js` — no cycles). Values: scalars, `Range` (2-D array, `.ref` when from cells), `RefValue` (unresolved reference; `evalAny` keeps refs, `evaluate` derefs), `Lambda`. Function flags: `lazy(fn)` gets AST args + `(ctx, ev)`; `refFn(fn)` gets RefValue args; `lift(fn, positions)` maps scalar functions over arrays and sets `fn.liftPos` (used by `mayReturnArray` to find spill candidates). Normal functions receive error values as arguments (not thrown).
- Dynamic arrays: `evaluateArray` returns a Range for multi-cell results; `Workbook.placeSpill` spills into empty cells (`#SPILL!` if blocked), `ensureSpills()` evaluates all `cell.maybeArray` formulas before reading empty cells. xlsx writes them as `t="array"` + `cm="1"` + `xl/metadata.xml`; `xlfn.js` converts app formulas ↔ file form (`_xlfn.`, `_xlpm.`, `_xleta.`, `SINGLE`, `ANCHORARRAY`, legacy implicit intersection `@`).
- Defined names: `wb.names = [{ name, ref: '=…', sheet|null, comment, hidden }]` (change via `wb.setNames`); evaluated by `wb.nameValue` (range → RefValue, constant, or LAMBDA). Cells may carry `cached` (file value for unsupported functions) and `link` (hyperlink URL or `#Sheet!A1`).
- Pivot defs use field names (also `styleDef` = custom style parts from the file's `<tableStyles>`, `errorCaption` = Excel's showError/errorCaption): `{ name, rows, cols, values: [{ field, agg, name, showAs, numFmt… }], pages, filters, calcFields: [{name, formula}], sort: {field: {dir, by}}, order: {field: [items]}, fieldFilters: {field: top|label|value}, style, styleOpts, rowCaption, cellFmt, captureFmt, layout, subtotals, grandRows, grandCols, top, left, area }` (old `rowField/colField/valueField` defs are still read via `normalizeDef`). Calculated fields are evaluated on field sums (`makeMeasures`). `computePivot` cells carry a `role` (rowHead, valueHead:i, rowItem:d, data:i, grandData…) used for pivot styles (`pivotStyleParts`/`roleStyle`) and for re-applying formats captured from the file (`cellFmt`). A sheet has `pivot` and `pivotsExtra`; in app.js use `allPivots`/`findPivotEntry`/`putPivotDef` and refer to pivots by `{sheet, name}`. Source reads are cached per sheet version (`wb.sheetVersion(si)`).
- Sheet size is `MAX_ROWS` = 10,000,000 (more than Excel's `EXCEL_MAX_ROWS` = 1,048,576); xlsx export drops rows beyond Excel's limit and `xlsxOverflow()` reports how many.
- Charts: `{ type ('combo' allowed), series: [{name, cat, val}], seriesFmt: [{type, axis: 'secondary', color, labels, marker, numFmt}], labels, legend, grouping, axes, pivot: {sheet, name} }`; `chartModelData` (chart.js) resolves ranges or pivot data (`pivotChartData`) — use it for rendering and export.
- In-cell pictures: cell `image: {src, alt}` with `raw: ''` (value is a `CellImage` from fxcore; `IMAGE()` returns one too). Keep `image` when rewriting a cell's style/comment. xlsx: richData (`metadata.xml` valueMetadata XLRICHVALUE + `xl/richData/*`, cells `t="e" vm="N"`).
- Recalc model: formula results are cached per sheet (`wb.caches[si]`). Imported sheets have `fileValues: true` and formula cells keep Excel's saved result in `cell.cached`; `getValue` returns it until that sheet (or a sheet it references) changes. `wb.invalidate(si)` clears only `si` and its dependents (`wb.affected(si)`, from the static `sheetDeps()` map built from formula ASTs: sheet refs, table refs, names → all, INDIRECT → all); `wb.invalidate()` with no argument is a full recalc (F9) and drops all file values — avoid it. Structural changes use `invalidateStructure()` (keeps file values). Formula ASTs are shared per formula text (`parseMemo`) and must never be mutated.
- Undo: cell edits record `cell` entries; sheet props use `propSnap` (per property); `CALC_NEUTRAL` props don't recalc. Row/column insert uses `snapshotSheet(si)` (that sheet only) + recorded cells in other sheets; add/delete/move sheet uses `snapshotList()` (sheet objects are kept, so `putSheet` restores contents in place); rename records a `rename` entry; names use `snapshotNames()`. `snapshotAll()` serializes the whole workbook — don't use it in new code (too slow for big files).
- Big files: `readXlsxAsync` / `writeXlsxAsync` / `wb.loadAsync` run step generators and yield to the browser (progress via `progressOverlay` in app.js); the sync `readXlsx` / `writeXlsx` drive the same generators (tests use them). Sheet XML is read by the DOM-free `scanRows` scanner (not `parseXml`); `readXlsx` returns `sheet.cells` as a `Map`. `zipAsync` compresses with the browser's `CompressionStream`. Workbooks over 50k cells autosave to IndexedDB per sheet, only changed sheets (`sheet._ev` edit counter, `_sid` id), as gzip Blob chunks (`saveBigToIdb` / `loadBigFromIdb`); over 300k cells the server is saved only on explicit [저장].
- Per-sheet `noGrid` (gridlines hidden, xlsx `showGridLines="0"`) is a sheet prop.
- Drawing objects live in sheet props `charts`, `images` (data-URL `src`), `shapes`; each has `id`, `x/y/w/h` (sheet px) and optional `z` (stacking). `app.js` tracks the selected object id in `chartSel` for all three kinds.
- Macros: `wb.vba = { bin (base64 vbaProject.bin), codeName, sheetCodes }` is preserved and written back as .xlsm; VBA is displayed, never executed.
- Number formats: builtin `style.numFmt` ids, or `numFmt: 'custom'` + `style.code` (Excel format code rendered by `formatCode()` in `format.js`). `styleForCode()` maps a code back to a builtin id when it renders identically; use it for anything that sets a format from a code (dialog, xlsx import).
- Tables: `sheet.tables` (`tables.js`): `{ id, name, r1..c2, header, totals, style, banded…, filter: {criteria, hidden, sort}|null, totalsFns }`; r1 = header row, r2 = totals row when present. Each table has its own filter; code that reads filters must go through `getFilter(key)`/`putFilter(key, f)` in `app.js` (key '' = sheet autofilter, else table id), and hidden rows are the union of `hiddenRows`, `filter.hidden` and every table's `filter.hidden`. Structured refs (`표1[열]`, `[@열]`) are `sref` tokens resolved via `ctx.structRef` (needs the formula cell, see `ctxFor(si, r, c)`); they are written to xlsx in canonical English form (`canonicalRef`).
- Slicers: drawing objects in `sheet.slicers` with `source: {kind:'table', table, column} | {kind:'pivot', field, pivots: [{sheet, name}]}` (legacy `self`/`sheet` still read; one slicer can drive several pivots — `slicerPivotTargets`); style is `sl.style` (`slicerstyle.js`) + optional `sl.custom` colors; their selection is not stored on the slicer but derived from the table filter criteria / `pivot.filters` (see `slicerModel`). In xlsx they are written as native Excel slicers (slicerCache + slicers parts, drawing `mc:AlternateContent`, workbook/sheet `extLst` with the x14/x15 GUIDs) and read back in `linkPivotsAndSlicers`.
- Pivot tables: source rows and filtering live in DOM-free `pivot.js` (`pivotSourceData`, `resolvePivot`, `computePivot`); `writePivot` in app.js renders the grid into `def.area`. xlsx export writes a real pivotCacheDefinition (saveData="0", refreshOnLoad="1") + pivotTableDefinition with all row/col/page/data fields; import rebuilds the full def.
- Conditional formatting: `condfmt.js` evaluates rules (`prepareCond`, `condFormatAt`). `sheet.cond` order is priority: index 0 wins over later rules and `stopIfTrue` stops later rules; `addCondRule` inserts at the front like Excel. xlsx `priority` maps to this order.
- Ribbon: contextual tabs use `context: 'table' | 'slicer'` in `TABS` and are shown from `ribbonState().context`. Alt key sequences (KeyTips, e.g. Alt+A+E) are defined in `KEYTIPS` in `app.js`.
- The grid is virtualized (`src/view.js`): only visible rows/cols are rendered; positions come from `Axis` (default size + sparse custom sizes/hidden). Never loop over all 1,048,576 rows — clamp to `wb.usedRange()` / `wb.extent()`.
- All UI text is Korean. New functions need an entry in `src/funcinfo.js` (a test enforces this).
- Undoable changes go through `wb.transact(fn, meta())`; sheet-level props (freeze, filter, charts, merges, hidden rows…) use `wb.setSheetProp()` / methods that call `wb.snapshotAll()`. New sheet props must be added to `SHEET_PROPS` in `workbook.js` and adjusted in `shiftAxis`.
- `server.js` stores documents in `tabula/data/` (gitignored) via `/api/files`; the app falls back to localStorage when the API is absent.
