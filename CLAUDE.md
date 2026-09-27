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
- DOM-free modules (unit-tested in Node): `formula.js`, `workbook.js`, `format.js`, `axis.js`, `xlsx.js`, `zip.js`, `xml.js`, `chart.js`, `pivot.js`, `series.js`, `csv.js`, `validation.js`, `shapes.js`, `vba.js`.
- Sheet size is `MAX_ROWS` = 10,000,000 (more than Excel's `EXCEL_MAX_ROWS` = 1,048,576); xlsx export drops rows beyond Excel's limit and `xlsxOverflow()` reports how many.
- Drawing objects live in sheet props `charts`, `images` (data-URL `src`), `shapes`; each has `id`, `x/y/w/h` (sheet px) and optional `z` (stacking). `app.js` tracks the selected object id in `chartSel` for all three kinds.
- Macros: `wb.vba = { bin (base64 vbaProject.bin), codeName, sheetCodes }` is preserved and written back as .xlsm; VBA is displayed, never executed.
- The grid is virtualized (`src/view.js`): only visible rows/cols are rendered; positions come from `Axis` (default size + sparse custom sizes/hidden). Never loop over all 1,048,576 rows — clamp to `wb.usedRange()` / `wb.extent()`.
- All UI text is Korean. New functions need an entry in `src/funcinfo.js` (a test enforces this).
- Undoable changes go through `wb.transact(fn, meta())`; sheet-level props (freeze, filter, charts, merges, hidden rows…) use `wb.setSheetProp()` / methods that call `wb.snapshotAll()`. New sheet props must be added to `SHEET_PROPS` in `workbook.js` and adjusted in `shiftAxis`.
- `server.js` stores documents in `tabula/data/` (gitignored) via `/api/files`; the app falls back to localStorage when the API is absent.
