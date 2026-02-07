# CLAUDE.md

## Project Overview

**네이버 파워링크 광고 순위 체크 프로그램** (Naver Power Link Ad Rank Checker)

A Python CLI tool that checks the ranking position of a specific business in Naver's paid ad search results (파워링크) for given keywords. It checks both PC and mobile ad search pages separately and reports rankings in a formatted table or CSV export.

## Repository Structure

```
revinchu/
├── CLAUDE.md                    # AI assistant guide (this file)
├── .gitignore                   # Excludes __pycache__, debug HTML, CSV, venv, .env
├── requirements.txt             # Python dependencies (3 packages)
├── keywords.txt                 # Sample keyword list (one per line, # for comments)
├── naver_ad_rank_checker.py     # Main rank checking program (478 lines)
└── debug_html_structure.py      # HTML structure analysis utility (141 lines)
```

**Output directories (gitignored):**
- `debug_html/` — Raw HTML files saved when using `--debug` flag
- `debug_*.html` — HTML files from `debug_html_structure.py`
- `*.csv` — Exported result files

## Tech Stack

- **Language**: Python 3.11+ (uses `list[T]` generics syntax from 3.9+)
- **HTTP client**: `requests` >= 2.28.0
- **HTML parser**: `beautifulsoup4` >= 4.12.0 with `lxml` >= 4.9.0 backend
- **No framework** — standalone CLI scripts using `argparse`

## Setup

```bash
pip install -r requirements.txt
```

No build step, compilation, or virtual environment setup is required beyond installing dependencies.

## Key Commands

### Run rank checker

```bash
# With keyword file
python naver_ad_rank_checker.py -k keywords.txt -t 하늘마음

# Single keyword query
python naver_ad_rank_checker.py -q 부산가려움증 -t 하늘마음

# Export results to CSV
python naver_ad_rank_checker.py -k keywords.txt -t 하늘마음 -o result.csv

# Debug mode (saves raw HTML to debug_html/ folder)
python naver_ad_rank_checker.py -q 부산가려움증 -t 하늘마음 --debug

# Custom delay between requests (default: 1.0s)
python naver_ad_rank_checker.py -k keywords.txt -t 하늘마음 --delay 2.0
```

### CLI Arguments

| Argument | Required | Description |
|----------|----------|-------------|
| `-k` / `--keywords-file` | Yes (or `-q`) | Path to keyword list file (one per line) |
| `-q` / `--query` | Yes (or `-k`) | Single keyword to check |
| `-t` / `--target` | Yes | Business name to search for (substring match) |
| `-o` / `--output` | No | CSV output file path |
| `--debug` | No | Save fetched HTML to `debug_html/` |
| `--delay` | No | Seconds between requests (default: 1.0) |

Note: `-k` and `-q` are mutually exclusive.

### Analyze HTML structure (debugging)

```bash
python debug_html_structure.py 부산가려움증
```

Fetches both PC and mobile ad pages, prints detailed HTML structure analysis (7 sections), and saves raw HTML files for inspection.

## Architecture

### Entry Points

- `naver_ad_rank_checker.py:main()` (line 391) — Main program entry point
- `debug_html_structure.py:main()` (line 101) — Debug utility entry point

### Data Model

```python
@dataclass AdResult           # naver_ad_rank_checker.py:45 — Single ad (rank, title, description, url)
@dataclass RankCheckResult    # naver_ad_rank_checker.py:54 — Full result for one keyword (PC + mobile ranks, ad lists)
```

`RankCheckResult.pc_rank` / `mobile_rank` = 0 means "not found" (미발견).

### Core Functions (`naver_ad_rank_checker.py`)

| Function | Line | Purpose |
|----------|------|---------|
| `fetch_page(url, user_agent)` | 68 | HTTP GET with Korean-locale headers, timeout=15s |
| `parse_pc_ads(html)` | 84 | Parse PC ad results using 3 fallback strategies |
| `parse_mobile_ads(html)` | 179 | Parse mobile ad results using 3 fallback strategies |
| `find_target_rank(ads, target)` | 254 | Substring-match target business in ad list |
| `check_keyword(keyword, target, delay)` | 268 | Orchestrate PC + mobile check for one keyword |
| `load_keywords(filepath)` | 298 | Read keyword file (skip blanks and `#` comments) |
| `print_results(results, target)` | 315 | Print formatted table + detailed ad listing to stdout |
| `export_csv(results, target, filepath)` | 359 | Write CSV with UTF-8 BOM (`utf-8-sig`) |
| `save_debug_html(keyword, html, device)` | 381 | Save raw HTML to `debug_html/` directory |

### URL Patterns

| Platform | URL Template |
|----------|-------------|
| PC | `https://ad.search.naver.com/search.naver?where=ad&query={query}` |
| Mobile | `https://m.ad.search.naver.com/search.naver?where=m_expd&query={query}` |

Keywords are URL-encoded via `urllib.parse.quote()`.

### HTML Parsing Strategy (3-level fallback)

Both `parse_pc_ads()` and `parse_mobile_ads()` use the same 3-level strategy:

**Strategy 1 — CSS selector-based** (most reliable):
- Tries multiple list selectors: `ul.lst_type li`, `ol.lst_type li`, `div.lst_type > div`, etc.
- Within each item, looks for title via: `a.lnk_tit`, `.tit_area a`, `a.tit`, `.area_text_title a`, `h3 a`, `a[class*='tit']`, `.title a`
- Extracts description from: `.ad_dsc`, `.dsc_area`, `.desc`, `p`

**Strategy 2 — Data attribute-based**:
- Selects `[data-cr-area] > div` or `[data-area] > div`
- Extracts first `<a>` text as title

**Strategy 3 — Heading/link fallback**:
- PC: Scans all `<a>` tags, filters by parent being h2/h3/h4/strong/em or class containing "tit"; skips navigation text
- PC wider: Falls back to all h2/h3/h4 text
- Mobile: Scans h2/h3/h4 tags, takes linked or heading text; skips navigation text

### Business Name Matching (`find_target_rank`)

Uses **bidirectional substring matching**:
- `target.lower() in ad_title.lower()` — target is part of title
- `ad_title.lower() in target.lower()` — title is part of target

Example: searching "하늘마음" matches "하늘마음한의원 덕천".

Returns `(0, "")` if no match found (displayed as "미노출").

### Request Flow

For each keyword:
1. Fetch PC ad page → parse → find target rank
2. `time.sleep(delay)` (default 1s)
3. Fetch mobile ad page → parse → find target rank
4. `time.sleep(delay)` between keywords

In debug mode, HTML is saved before parsing in the main loop (lines 443-456), and `check_keyword` also fetches (resulting in double fetches during debug).

## Code Conventions

### Naming
- **Functions**: `snake_case` — `fetch_page`, `parse_pc_ads`, `find_target_rank`
- **Classes/dataclasses**: `PascalCase` — `AdResult`, `RankCheckResult`
- **Constants**: `UPPER_SNAKE_CASE` — `PC_URL_TEMPLATE`, `REQUEST_TIMEOUT`, `DEFAULT_DELAY`
- **Variables**: `snake_case` — `pc_rank`, `mobile_rank`, `matched_title`

### Type Hints
- Used on all function signatures and dataclass fields
- Uses Python 3.9+ generic syntax: `list[AdResult]`, `tuple[int, str]`

### Error Handling
- Network errors caught per-request with `requests.RequestException`
- Errors printed to stderr in Korean: `print(f"  [PC 오류] ...", file=sys.stderr)`
- Graceful degradation: if one request fails, the tool continues with remaining keywords

### Output
- All user-facing output is in **Korean**
- stdout for normal output, stderr for errors
- CSV uses `utf-8-sig` encoding (UTF-8 with BOM) for Korean Excel compatibility
- Terminal output includes formatted table and detailed ad listing with ★ markers

### Constants (top of `naver_ad_rank_checker.py`)
```python
PC_URL_TEMPLATE       # PC ad search URL pattern
MOBILE_URL_TEMPLATE   # Mobile ad search URL pattern
PC_USER_AGENT         # Chrome 131 on Windows 10
MOBILE_USER_AGENT     # Safari on iPhone iOS 17.0
REQUEST_TIMEOUT = 15  # HTTP timeout in seconds
DEFAULT_DELAY = 1.0   # Delay between requests in seconds
```

Note: `debug_html_structure.py` duplicates the User-Agent constants (not imported from main module).

## Testing

No automated tests. This is a scraping tool whose correctness depends on Naver's live HTML structure. Verify manually by running against known keywords and comparing output against the Naver ad page in a browser.

## Important Notes for AI Assistants

- **Naver HTML changes frequently**: If parsing stops working, run `debug_html_structure.py` to analyze current HTML structure, then update selectors in `parse_pc_ads()` and `parse_mobile_ads()`
- **Rate limiting**: Keep delay between requests >= 1s to avoid being blocked by Naver
- **User-Agent strings**: Must be kept up-to-date with current browser versions; Naver may block outdated UA strings
- **Proxy**: Do not set `http_proxy`/`https_proxy` unless explicitly required; the script accesses Naver directly
- **Encoding**: Naver pages use UTF-8; CSV export uses UTF-8 with BOM for Korean Excel compatibility
- **Debug double-fetch**: In `--debug` mode, HTML is fetched twice per keyword (once for saving, once for parsing in `check_keyword`); this is by design but worth noting
- **No package structure**: This is a flat script project — no `__init__.py`, no setup.py/pyproject.toml
- **Duplicate constants**: User-Agent strings are defined in both Python files independently; changes must be synced manually
