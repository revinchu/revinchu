# SEO · AEO · GEO 체크리스트

검색엔진(SEO), 답변 엔진(AEO: 구글 AI 개요·추천 스니펫·음성 비서), 생성형 AI(GEO: ChatGPT·Claude·Perplexity·Gemini 등)가
이 사이트의 정보를 **찾고 → 이해하고 → 인용**하게 만들기 위한 점검표입니다.

- `[x]` 사이트에 구현됨 · `[ ]` 해야 할 일 · **(직접)** 계정 로그인·소유권 인증이 필요해 운영자가 해야 하는 일
- 마지막 점검: 2026-10-29

---

## 1. 크롤링·색인 (기술 SEO)

- [x] `robots.txt` — 전체 공개, `/admin`·상담 접수 완료 페이지만 제외, 사이트맵 위치 명시
- [x] AI 크롤러 명시적 허용 — GPTBot, OAI-SearchBot, ChatGPT-User, ClaudeBot, Claude-User, Claude-SearchBot, PerplexityBot, Google-Extended, Applebot-Extended, CCBot, Amazonbot, meta-externalagent 등
- [x] 국내·해외 검색 크롤러 허용 — Googlebot, Bingbot, Yeti(네이버), Daum, Baiduspider, YandexBot
- [x] `sitemap.xml` — 5개 언어 전체 페이지, 병원 상세, 매거진 글, `lastmod`, `hreflang` 대체 링크 포함
- [x] 서버 렌더링(HTML에 본문이 그대로 있음) — 자바스크립트 없이도 크롤러가 내용을 읽음
- [x] 언어별 고유 URL (`/en/`, `/ko/`, `/zh/`, `/ja/`, `/vi/`) + `hreflang` + `x-default`
- [x] `canonical` — 번역본이 없는 글은 원문(영어/한국어)을 canonical 로 지정해 중복 색인 방지
- [x] `meta robots` — `max-image-preview:large, max-snippet:-1` (큰 이미지·긴 스니펫 허용)
- [ ] **(직접)** 커스텀 도메인 연결 — `workers.dev` 보다 자체 도메인이 신뢰도·브랜드 검색에 유리. 연결 후 `SITE_URL` 시크릿 변경
- [ ] **(직접)** 커스텀 도메인 연결 시 Cloudflare 대시보드 → **AI Crawl Control / Bot 설정**에서 "AI 봇 차단"이 꺼져 있는지 확인 (새 도메인은 기본 차단일 수 있음)
- [ ] 404 페이지 다국어화, 깨진 링크 정기 점검 (월 1회)
- [ ] 페이지 속도 점검 (PageSpeed Insights, 모바일 LCP 2.5초 이내 목표). 첫 요청이 느린 무료 플랜 특성 → 유료 플랜 검토

## 2. 검색엔진 등록 (직접)

- [ ] **(직접)** Google Search Console — 도메인 소유권 인증, `sitemap.xml` 제출, 색인 생성 요청
- [ ] **(직접)** Bing Webmaster Tools — 사이트맵 제출 (ChatGPT 검색·Copilot 이 Bing 색인을 사용)
- [ ] **(직접)** 네이버 서치어드바이저 — 소유권 확인, 사이트맵·RSS 제출 (국내 검색)
- [ ] **(직접)** 다음(카카오) 웹마스터도구 등록
- [ ] **(직접)** 바이두 站长平台 (중국어 사용자), Yandex Webmaster (러시아어권 확장 시)
- [ ] IndexNow 연동 (Bing·Yandex 에 새 글을 즉시 알림) — 글 발행 시 자동 호출

## 3. 구조화 데이터 (Schema.org / JSON-LD)

- [x] 모든 페이지: `Organization` + `WebSite`(사이트 검색 `SearchAction`)
- [x] 매거진 글: `Article`(제목·설명·요약·이미지·작성/수정일·저자·출처 `citation`·`speakable`)
- [x] 매거진 글: `BreadcrumbList`
- [x] 매거진 글: `FAQPage` (본문 FAQ 섹션에서 자동 생성)
- [x] 매거진 목록: `CollectionPage` + `ItemList`
- [x] 병원 상세: `Hospital` / `Dentist` (이름·현지어 이름·주소·좌표·전화·진료과·공개 가격 `Offer`)
- [x] Open Graph / Twitter 카드 (공유 미리보기 이미지·제목·설명·언어)
- [ ] 관광지에 `TouristAttraction`, 숙소에 `LodgingBusiness`, 식당에 `Restaurant` 추가
- [ ] **(직접)** 구글 리치 결과 테스트(search.google.com/test/rich-results)로 오류 확인

## 4. AEO — 답변 엔진이 바로 인용하는 글쓰기

- [x] 글 맨 위 **핵심 요약(Quick answer)** 상자 — 2~4문장으로 질문에 직접 답함 (`.answer-box`, `speakable` 지정)
- [x] 소제목을 실제 검색 질문 형태로 ("How much does…", "Do I need…")
- [x] FAQ 섹션 3~4개 질문 + 짧고 완결된 답변
- [x] 숫자·표 중심 정보 (가격대, 소요 시간, 요일별 일정 등) — 답변 엔진이 표를 그대로 인용
- [x] 목차(앵커 링크) — 섹션 단위 인용·딥링크
- [x] 작성일·수정일 표시 (최신성 신호)
- [ ] 새 글마다: 대표 질문 1개를 정하고, 첫 문단에서 40~60단어로 답하기
- [ ] 한국어 글도 네이버 지식스니펫 형식(정의 → 목록 → 표)으로 다듬기

## 5. GEO — 생성형 AI가 신뢰하고 인용하게

- [x] `/llms.txt` — 사이트 소개, 핵심 페이지, 전체 글 목록 (llmstxt.org 형식)
- [x] `/llms-full.txt` — 영어 글 전문(마크다운)을 한 파일로 제공
- [x] 모든 글에 **공식 출처** 목록 (정부·공공기관 링크) — `citation` 구조화 데이터로도 제공
- [x] **1차 데이터** — 심평원 공개 가격을 직접 집계한 표 (다른 곳에 없는 정보 = 인용 가치)
- [x] 편집 원칙 고지 — "공식 출처 기반, 의학적 조언 아님"
- [x] 일관된 고유명사 — 병원·지역·명소 이름을 5개 언어로 통일 (`cloudflare/name_i18n.py`)
- [ ] 브랜드 엔티티 강화 — 회사 소개 페이지(About), 연락처, 운영 주체, 사업자 정보
- [ ] 외부 언급 확보 — 보도자료, 관광공사·지자체 협력, 여행 커뮤니티(레딧·트립어드바이저) 정보성 답변
- [ ] **(직접)** 위키데이터/구글 비즈니스 프로필 등록 (브랜드 확정 후)
- [ ] 분기마다 주요 AI(ChatGPT·Perplexity·Gemini·Claude)에 대표 질문 10개를 물어보고 인용 여부 기록

## 6. 콘텐츠 — 다양한 관심사로 유입, 오래 머물게

- [x] 매거진 카테고리 5개: K-의료, K-뷰티, K-팝, K-푸드, 여행 팁
- [x] 긴 정보성 글 10편 (영어·한국어) + 공공누리 사진, 표, 팁 상자
- [x] 글 → 병원 검색·가격·컨시어지·상담으로 이어지는 내부 링크
- [x] "함께 읽으면 좋은 글" 3편, 홈 화면 최신 글 4편, 상단 메뉴·푸터 카테고리 링크
- [ ] 중국어·일본어·베트남어 번역본 (현재는 영어 원문을 보여 주고 canonical 을 원문으로 지정)
- [ ] 발행 주기: 주 1~2편. 우선순위 주제 예시:
  - K-의료: 라식·라섹 비교, 임플란트 과정·기간, 한방 추나, 대학병원 국제진료센터 이용법, 의료비자 실전
  - K-뷰티: 성수 팝업 지도, 피부 타입별 선크림, 올리브영 세일 달력, 남성 그루밍
  - K-팝: 공연장별 교통·좌석 팁, 팬카페 가입법, 촬영지 투어, 한국어 응원법
  - K-푸드: 채식·할랄 맛집, 비 오는 날 메뉴, 편의점 음식, 계절 음식
  - 여행: 계절별 서울·부산, 비 오는 날 실내 코스, 한강 즐기기, 휠체어·유모차 동선
- [ ] 글마다 키워드 조사(구글 트렌드·네이버 데이터랩·Answer the Public)로 실제 질문 반영
- [ ] 6개월마다 가격·규정이 바뀌는 글(비자, 세금 환급, 가격표) 재검토 후 `updated` 갱신

## 7. 신뢰(E-E-A-T)와 규제

- [x] 가격은 공공데이터 원문·기준일과 함께 표시, 외국인 가격은 별도 확인 안내
- [x] 공개 정보 업체는 "제휴처 아님" 명시
- [x] 의료 정보에 "의학적 조언 아님" 고지
- [ ] 의료 감수자(의사) 검토 받은 글은 감수자 이름·자격 표시 (실제 감수 후에만)
- [ ] 의료광고 심의 대상 여부 확인 (병원 홍보성 문구, 치료 전후 사진, 할인 표현 금지)
- [ ] 개인정보 처리방침·쿠키 고지 다국어 최신화

## 8. 측정

- [ ] **(직접)** Google Analytics 4 또는 Cloudflare Web Analytics 연결
- [ ] 관리자 대시보드 `page_events` 에 매거진 조회(`view_magazine`, `view_article`) 집계 표시
- [ ] 핵심 지표: 자연 검색 유입, 글 평균 체류 시간, 글 → 상담 신청 전환율, AI 유입(referrer: chatgpt.com, perplexity.ai 등)

---

### 구현 위치 (개발자용)

| 항목 | 파일 |
|---|---|
| robots.txt, sitemap.xml, llms.txt, llms-full.txt | `medical_platform/public.py` |
| 메타 태그, canonical, hreflang, OG, 공통 JSON-LD | `medical_platform/templates/public/base.html` |
| 글 엔진 (마크다운, 목차, FAQ 추출) | `medical_platform/magazine.py` |
| 글 원고 | `medical_platform/content/articles/<slug>.<lang>.md` |
| 글·병원 JSON-LD | `public.py` 의 `article_ld`, `hospital_ld` |
| 고유명사 번역 | `cloudflare/name_i18n.py`, `medical_platform/i18n.py` (`DISTRICTS`) |

### 새 글 쓰는 법

1. `medical_platform/content/articles/<영문-슬러그>.en.md` 와 `.ko.md` 를 만든다 (다른 언어는 `.zh.md`, `.ja.md`, `.vi.md`).
2. 앞머리에 `title, description, category, published, updated, hero, hero_alt, hero_credit, answer, tags` 를 쓴다.
3. 본문은 질문형 `##` 소제목, 표, 팁 상자(`> `), 사진(`![설명](주소 "출처")`), `## FAQ`, `## Sources` 순서를 권장.
4. 사진은 저작권이 확인된 것만 (공공누리 1·3유형은 출처 표시). 내부 링크는 `/{lang}/...` 형식.
5. 배포: `python cloudflare/build.py && uv run pywrangler deploy` — 사이트맵·llms.txt 에 자동 반영.
