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
├── medical_platform/            # WIZ X 해외환자 유치 웹 플랫폼 (Flask + SQLite), see its README.md
├── cloudflare/                  # Cloudflare Workers + D1 배포 (build.py, worker.py, wrangler.jsonc), see its README.md
└── kv_reference/                # KV 아틀라스: 광고 키비주얼·이미지 사이트 레퍼런스 색인 (정적 HTML/JS), see its README.md
```

## Medical Platform (`medical_platform/`)

Multilingual (en/ko/zh/ja/vi) hospital discovery site + lead-capture form + internal CRM for WIZ X.

- Run: `python -m medical_platform.app` · Test: `python -m pytest medical_platform/tests`
- Public routes are prefixed with `/<lang>/`; admin at `/admin` (Korean UI). Roles: admin / coordinator / marketer (no PII).
- Seed hospitals and partners are fictional `[Sample]` data (`is_demo=1`). Never add unverified real-hospital claims
  (명의, 유명인 방문, prices) without evidence/source fields.
- Do not collect medical records in the consult form; hospital referrals require third-party-share consent.
- Add UI strings to `i18n.T` for all five languages.
- DB access goes through `db.get_db()`; it returns sqlite3 locally or a D1 adapter (`DB_BACKEND=d1`) on Cloudflare.
  Use only execute/executemany/fetchone/fetchall/lastrowid so both work; tests run against both backends.
- Email: `notify.py` (SMTP/Resend/console). Never put patient contact details in staff emails.
- UI style: editorial / public-service look (no gradients, no rounded card grids, no emoji). Fonts, icons and Leaflet are vendored
  under `static/vendor/` (OFL/ISC/BSD); use the `icon('name')` template helper for Lucide icons.

## KV Atlas (`kv_reference/`)

Ad key-visual database tool for designers and marketers. Static HTML/JS (`data.js` → `app.js` → `live.js` → `ai.js`, then `KVA.boot()`).
Published as a claude.ai artifact with `sample` (Claude, images), `db`, `user`, `downloads` capabilities; see its README.md.

- Static taxonomy and classic references live in `data.js`; shared data lives in the artifact db (`ads`, `trends`, `overlays`, `meta/status`).
- Every capability can be absent (local file): features must degrade, never break the page.
- A weekly scheduled task writes trends/ads into the db; keep its prompt in sync when taxonomy values change.
- Do not embed or hotlink copyrighted campaign images; KV thumbnails are generated color/technique diagrams that link out to search.
- Only add campaign facts (year, agency) that are well documented; leave an entry out rather than guess.

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
