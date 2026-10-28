# KV 아틀라스 (`kv_reference/`)

디자이너와 마케터를 위한 광고 키비주얼 데이터베이스 도구. 정적 HTML/JS로 만들었고, claude.ai 아티팩트로 열면 공유 DB와 Claude(보는 사람의 계정)가 연결된다.

- 배포본: https://claude.ai/artifact/TNfSXR9tLTGXfhPRwtZ25U (비공개, 공유는 페이지의 Share 메뉴)
- 로컬: `python -m http.server -d kv_reference 8000` 후 http://localhost:8000. 로컬에서는 클래식 데이터만 보이고 AI·공유 DB 기능은 꺼진다.

## 파일

| 파일 | 역할 |
|------|------|
| `data.js` | 정적 데이터: 사이트 180곳, 클래식 키비주얼 67건, 기법 20종, 무드, 분야, 업종, 국가·권역, 광고 규격, 다국어 KV 키워드 |
| `app.js` | 코어: 레퍼런스 저장소(클래식 + 공유 DB), 필터, 상세, 사이트 디렉터리, 통합 검색, 무작위·훈련, 내 보드, 규격, 라우팅 |
| `live.js` | 공유 DB(`db` 기능): `ads`, `trends`, `overlays`, `meta/status` 구독과 쓰기, 트렌드 화면, 데이터 관리(가져오기·내보내기) |
| `ai.js` | Claude 기능(`sample` 기능): 역분석 랩, 상세 역분석, 콘셉트 스튜디오, 트렌드 브리핑, 붙여넣기 구조화, 브라우저 이미지 특징 계산 |
| `ingest/tvcf_ingest.py` | TVCF 우수 인쇄광고 목록 → 가져오기용 JSON (TVCF에 접속되는 PC에서 실행) |

## 화면

- **키비주얼 DB**: 출처(클래식/TVCF/트렌드 수집/직접 분석), 권역, 국가, 분야, 업종, 시대, 기법, 매체·유형, 무드로 필터한다.
- **트렌드**: 매주 예약 작업이 조사한 권역·업종별 리포트(신호, 시사점, 사례, 출처)를 보여 주고, AI 트렌드 브리핑을 만든다.
- **역분석 랩**: 이미지를 올리면 다음을 해 준다.
  1. 색·구성이 비슷한 레퍼런스를 바로 찾는다. k-means 팔레트(Lab 거리), dHash, 밝기·채도·에지 밀도를 쓴다.
  2. Claude가 분류, 커뮤니케이션 컨셉(타깃·인사이트·문제·약속·컨셉·빅 아이디어·행동·이유), 생각의 고리(과제→인사이트→컨셉→시각 번역→실행), 시각 문법, 역프롬프트(Midjourney, FLUX, GPT Image, Imagen, Firefly, 영상, 네거티브), 변주 아이디어, 다국어 검색어, DB에서 고른 유사 레퍼런스를 추론한다.
  3. 결과를 DB에 등록하면 다음 검색부터 유사도 비교 대상이 된다.
- **콘셉트 스튜디오**: 질문 한 줄에서 출발해 아래 단계를 차례로 진행한다. 단계마다 수정 지시를 넣어 다시 생성할 수 있고, 결과는 Markdown 기획서로 복사하거나 내려받는다.
  1. 브리프
  2. 타깃 세그먼트
  3. KPI
  4. 인사이트·컨셉(점수 포함)
  5. KV 시안·매체 전개·프롬프트
  6. 영상 스토리보드
  7. 반론·리스크 점검
- **데이터**: TVCF 등 목록을 붙여 넣으면 Claude가 정리해 등록한다. JSON 가져오기와 JSON·CSV 내보내기, 수집 로그도 이 화면에 있다.
- 기존 화면: 사이트 디렉터리, 통합 검색, 무작위·훈련(브리프를 콘셉트 스튜디오로 보내기), 내 보드, 광고 규격.

## 썸네일 미리보기와 이미지 검색 (직접 호스팅)

claude.ai 아티팩트는 보안 정책상 외부 이미지를 표시하지 않는다. 그래서 썸네일과 이미지 검색은 같은 페이지를 Cloudflare Worker(`worker/`)로 호스팅한 주소에서 켜진다. 페이지가 같은 주소의 `/api/health`를 찾으면 `hosted.js`가 호스팅 모드로 바뀐다.

- **카드 썸네일:** 레퍼런스에 이미지 주소가 있으면(MoMA 5,248건 등) 원본 사이트에서 바로 불러온다. 없으면 출처 링크의 대표 이미지(og:image 등)를 `/api/preview`로 찾아 원본 주소에서 불러온다. Worker는 이미지 주소만 돌려주고 파일은 저장·중계하지 않는다.
- **통합 검색 → 이미지로 보기:** `/api/images`가 여러 출처의 썸네일과 원본 링크를 모아 보여 준다.
  - 키 없이 켜지는 출처: Openverse, Wikimedia Commons, Art Institute of Chicago, The Met
  - 키를 넣으면 켜지는 출처: Brave 이미지 검색(웹 전체), Google 프로그래머블 검색, Unsplash, Pexels, Pixabay
- **AI 기능:** `/api/claude`가 서버의 Anthropic API 키로 `claude-opus-5-5`를 호출한다. 거절된 요청은 `fallbacks: "default"`로 다른 모델이 다시 받는다. 쓰려면 페이지의 데이터 탭에서 접근 토큰(`ACCESS_TOKEN`)을 입력한다.
- 공유 DB(트렌드, 등록한 역분석)는 아티팩트에만 있다. 호스팅 주소에서는 클래식·대량 데이터와 이미지 기능만 동작한다.

```bash
cd kv_reference/worker
npm install
npx wrangler login                       # 또는 CLOUDFLARE_API_TOKEN 환경 변수
npx wrangler secret put ANTHROPIC_API_KEY  # 선택: AI 기능
npx wrangler secret put ACCESS_TOKEN       # 선택: AI 기능을 쓸 사람에게 줄 토큰
npx wrangler secret put BRAVE_API_KEY      # 선택: 웹 전체 이미지 검색
npm run deploy                           # → https://kv-atlas.<계정>.workers.dev
npm test                                 # 미리보기 추출·이미지 검색 단위 테스트
```

## 대량 데이터 (수십만 건 규모)

수집 스크립트가 JSONL을 만들고, `build_bulk.py`가 이를 출처·국가별 샤드(최대 4만 건, 약 8MB)로 나눠 `bulk/`에 쓴다. 페이지(`bulk.js`)는 `bulk/manifest.json`을 읽어 처음 8만 건을 불러오고, 나머지는 출처 칩이나 국가를 고를 때 불러온다. 아티팩트 파일 한도 안에서 약 100만 건까지 담을 수 있다. 30만 건 합성 데이터로 검색 약 0.3초, 메모리 약 180MB를 확인했다.

| 수집기 | 원천 | 규모 | 필요한 것 |
|------|------|------|------|
| `ingest/museums.py` | MoMA·The Met CC0 메타데이터 (포스터·광고물) | 11,824건 (포함됨) | 없음 |
| `ingest/google_ads_transparency.py` | Google 광고 투명성 센터 BigQuery 공개 데이터 (이미지 광고) | 국가당 수만~수십만 건 | Google Cloud 서비스 계정 키, 쿼리 비용(월 1TB 무료) |
| `ingest/meta_ad_library.py` | Meta 광고 라이브러리 API | 검색어당 최대 수천 건 (EU·정치 광고 중심) | Meta 개발자 액세스 토큰 |
| `ingest/pitt_ads.py` | Pitt Image Ads Dataset 주석 (토픽·감정·행동·이유) | 64,832건 | people.cs.pitt.edu 접속, 연구 목적 사용 |
| `ingest/tvcf_ingest.py` | TVCF 우수 인쇄광고 목록 | 목록 페이지 단위 | tvcf.co.kr 접속 |

```bash
python kv_reference/ingest/museums.py --moma Artworks.csv --met MetObjects.csv -o museums.jsonl
python kv_reference/ingest/google_ads_transparency.py --countries KR,JP,US --per-country 50000 --since 2026-01-01 --dry-run
python kv_reference/ingest/build_bulk.py museums.jsonl gads.jsonl -o kv_reference/bulk
```

저장소에는 뮤지엄 샤드만 둔다. 외부 대량 샤드(`gads-*`, `meta-*`, `pitt-*`, `tvcf-*`)는 .gitignore 로 빼고 아티팩트에만 올린다.
Google·Meta·Pitt 수집기는 인증과 접속이 없는 환경에서 작성해 실제 응답으로는 아직 검증하지 않았다.

## 데이터 갱신

- 예약 작업 "KV 아틀라스 주간 트렌드 수집"이 매주 월요일 10:17(KST)에 새 세션을 띄워 웹 검색으로 사례와 트렌드를 조사하고, `ArtifactData`로 DB에 쓴다.
- TVCF 자동 수집은 Claude Code 환경의 네트워크 허용 도메인에 `tvcf.co.kr`이 있어야 동작한다. 막혀 있으면 데이터 화면의 붙여넣기나 `ingest/tvcf_ingest.py`를 쓴다.

## 데이터 원칙

- 클래식 키비주얼은 원본 이미지를 싣지 않는다. 썸네일은 대표 색과 기법을 도식화한 것이다.
- 수집 데이터는 출처 링크와 메타데이터만 저장하고 이미지 파일은 복제하지 않는다. 역분석 랩에서 사용자가 등록한 이미지만 320px 썸네일로 저장된다.
- 연도·제작사·날짜는 출처로 확인된 것만 적는다. AI 역분석 결과는 추론이며 실제 제작 의도와 다를 수 있다.
- 분류값(분야·업종·기법·무드·국가)은 `data.js`의 목록을 쓴다. 새 값을 추가하면 예약 작업의 프롬프트도 함께 고친다.

## 참고한 오픈소스와 연구

- 컨셉 역추론의 "행동·이유" 형식: Pitt Image Ads Dataset (Hussain et al., CVPR 2017)의 "I should … because …" 주석 방식
- 팔레트 추출: Color Thief(MIT)의 지배색 추출 아이디어를 k-means로 다시 구현
- 지각 해시: blockhash-js(MIT)·dHash 방식
- 역프롬프트: CLIP Interrogator의 이미지→프롬프트 접근. 브라우저 CLIP 모델은 아티팩트 보안 정책상 외부 모델 파일을 받을 수 없어 Claude 비전으로 대신했다.
