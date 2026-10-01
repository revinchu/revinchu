# AGENTS.md — 저장소 안내 (OpenAI Codex · 기타 코딩 에이전트용)

이 저장소에는 서로 관계없는 두 프로젝트가 있습니다.

| 경로 | 내용 | 상태 |
|------|------|------|
| `tabula/` | **WIXEL(위셀)** — 엑셀과 같은 브라우저 스프레드시트 (순수 ES 모듈, 의존성 0) | 주 개발 대상. 버전 2.0.0 |
| 루트 `*.py` | 네이버 파워링크 광고 순위 체크 CLI (Python) | 별도 소규모 도구 |

**WIXEL 작업이라면 먼저 [`tabula/AGENTS.md`](tabula/AGENTS.md) 를 읽으세요.** 인수인계 문서는 `tabula/docs/codex/` 에 있습니다.

## 공통 규칙
- 사용자에게 보이는 모든 글자(UI · 메시지 · 토스트 · 대화상자)는 **한국어**.
- 원본 인계는 `claude/offsro-excel-implementation-zmro3r`의 `f6cbc6e`입니다. 검증 도구·문서를 보완한 `codex/handoff-validation`의 인계 커밋에서 새 작업을 분기하세요. 원격 반영 여부는 `tabula/docs/codex/06_검증결과.md`와 실제 Git 상태를 확인하세요.
- 커밋 전에 `cd tabula && npm test`가 모두 통과해야 합니다. 실제 개수와 결과는 `tabula/docs/codex/06_검증결과.md`를 참고하세요.
- 기존 체크아웃 경로를 유지하고, 새 로컬 작업은 `D:\Codex\Workspaces`, 임시 파일은 `D:\Codex\Temp`, 캐시는 `D:\Codex\Caches`에 만드세요.
- 사용자가 올린 실제 업무 파일(xlsx · xlsb 등)은 **절대 커밋하지 마세요** (개인 · 회사 데이터).

## Python 순위 체크 도구 (참고)
- `pip install -r requirements.txt` → `python naver_ad_rank_checker.py -k keywords.txt -t 업체명`
- 네이버 HTML 이 자주 바뀌므로 파싱이 깨지면 `debug_html_structure.py` 로 구조를 확인하고 선택자를 고칩니다.
- 요청 간격 1초 이상 유지, CSV 는 `utf-8-sig`.
