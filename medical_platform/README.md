# WIZ X 해외환자 유치 플랫폼 MVP (가칭 MEDIPRISM)

해외 환자가 한국 병원을 **지도·목적별로 탐색**하고 **투명한 가격**을 확인한 뒤, 의료·여행 니즈를 담아 **상담 DB를 남기는** 웹사이트와, 위젝스 직원이 그 DB를 실제 방문까지 관리하는 **내부 CRM**입니다.

## 실행

```bash
pip install -r medical_platform/requirements.txt
python -m medical_platform.app          # http://127.0.0.1:5000  (관리자: /admin)
python -m pytest medical_platform/tests # 테스트
```

처음 실행하면 `medical_platform/instance/platform.db`(SQLite)가 생성되고 **가상 샘플 데이터**가 들어갑니다.

| 초기 계정 | 비밀번호 | 역할 |
|---|---|---|
| admin | admin1234 | 관리자 (전체) |
| coord_en / coord_zh / coord_ja | coord1234 | 코디네이터 (영/중/일·베) |
| marketer | market1234 | 마케터 (집계만, 개인정보 마스킹) |

> ⚠️ 운영 전 반드시 `SECRET_KEY` 환경변수 설정, 초기 비밀번호 변경, 샘플 병원·제휴사 교체.

## 공개 사이트 (EN · 한국어 · 简体中文 · 日本語 · Tiếng Việt)

| 경로 | 내용 |
|---|---|
| `/{lang}/` | 검색, **목적별 탐색**(가성비 / 한국인에게 유명 / 외국인에게 유명 / 명의 / 유명인 방문 / VIP / 외국인 특화), 진료과, 지도 |
| `/{lang}/hospitals` | **지도 + 리스트**, 진료과·지역·특화·지원언어·가격대 필터 |
| `/{lang}/hospitals/{slug}` | 병원 소개, **국내가 vs 외국인 적용가**(포함/불포함·확인일), 의료진 자격, **특화 태그의 근거·출처**, 주변 숙소·K푸드, 광고 표시 |
| `/{lang}/concierge` | 숙소·K푸드·통역·공항/교통·관광·VIP 의전 제휴사, 5일 여행 동선 예시 |
| `/{lang}/consult` | **상담 DB 수집 폼**: 연락처/메신저, 국적·거주국·상담언어, 진료 분야, 선호 병원 유형, 희망 병원, 예산, 일정·인원, 컨시어지 니즈, 숙소·음식·여행 선호, VIP 요청, 유입 경로, 동의 4종 |

- 광고 파라미터(UTM)·랜딩·리퍼러를 first-touch로 저장해 리드에 연결
- 허니팟 봇 차단, CSRF, 30일 내 동일 이메일/메신저 중복 표시
- 병력·검사결과는 받지 않음 (기획서 4-3)

## 관리자 (`/admin`)

- **대시보드**: 조회→상담→유효→병원연결→견적→예약→방문 퍼널, SLA(30분) 초과 미응대, 평균 최초 응대, 언어·국가·진료·UTM·캠페인·예산·컨시어지 니즈·병원 선호별 집계
- **환자 CRM**: 상태 10단계, 언어 기반 자동 배정, 우선순위(VIP·고예산 자동 '높음'), 상담·연락 기록, 다음 연락일, 병원 연결(**제3자 제공 동의 없으면 차단**), 컨시어지 요청-제휴사 매칭, 동의 갱신/철회, 개인정보 삭제(비식별화), CSV(엑셀용 BOM) 다운로드
- **병원 DB**: 검증 상태(수집 예정→자료 확보→검증 완료→게시→갱신 필요→비공개), 다국어 소개, 특화 태그+근거, 가격, 의료진, 확인일/다음 확인일, 광고 계약 여부
- **제휴 서비스 / 직원·권한 / 접속·감사 로그**

## 구조

```
medical_platform/
  app.py      앱 팩토리, CSRF
  db.py       스키마 + 샘플 데이터
  i18n.py     5개 언어 문자열, 코드 목록
  public.py   공개 사이트
  crm.py      리드 생성·자동 배정·중복·비식별화
  admin.py    관리자
  templates/ static/ tests/
```

## 운영 전 체크

- 번역문은 초안 → 원어민·의료광고 검수
- 개인정보 처리방침(`/privacy`)은 템플릿 → 법률 검토
- 외국인환자 유치업 등록·수수료·광고 규정 검토 후 상담/예약 기능 운영 범위 확정
- 운영 DB는 PostgreSQL + 백업, HTTPS 적용
- 지도는 Leaflet + OpenStreetMap 타일 (트래픽 증가 시 유료 타일/카카오·구글 지도로 교체)

## 웹 배포 (PythonAnywhere 무료 계정 예시)

결과 주소: `https://<사용자이름>.pythonanywhere.com` — 파일 저장소가 유지되므로 SQLite 상담 DB가 재시작 후에도 남습니다.

1. https://www.pythonanywhere.com 에서 무료(Beginner) 계정 가입
2. **Consoles → Bash** 에서:
   ```bash
   git clone -b claude/claude-md-mlbz4ceanrc28v1i-FFlyB https://github.com/revinchu/revinchu.git
   pip install --user flask
   ```
   (저장소가 비공개라면 GitHub 개인 액세스 토큰으로 clone)
3. **Web → Add a new web app → Manual configuration → Python 3.x** 선택
4. Web 탭의 **WSGI configuration file** 내용을 아래로 교체 (`<사용자이름>`과 비밀키 변경):
   ```python
   import os, sys
   sys.path.insert(0, "/home/<사용자이름>/revinchu")
   os.environ["SECRET_KEY"] = "충분히-길고-무작위인-문자열"
   from medical_platform.wsgi import application
   ```
5. **Reload** 클릭 → `https://<사용자이름>.pythonanywhere.com` 접속, `/admin` 로그인 후 **초기 비밀번호 즉시 변경**
6. 코드 업데이트 시: Bash에서 `cd revinchu && git pull` 후 Web 탭 Reload

다른 호스팅(Render·Railway·자체 서버)은 `gunicorn medical_platform.wsgi:application` 으로 실행하고,
상담 DB 보존을 위해 `DATABASE_PATH`를 영구 디스크 경로로 지정하세요.

## 디자인 원칙과 사용한 오픈소스

'AI가 만든 듯한' 기본 패턴(보라·청록 그라데이션, 둥근 카드 3단, 이모지 아이콘, 중앙 정렬 히어로)을 피하고,
정부·의료 공공 디자인 시스템(GOV.UK / NHS)의 원칙 — 라벨은 입력창 위, 밑줄 링크, 고대비 노란 포커스, 장식 없는 정보 우선 — 위에
에디토리얼(잡지·가이드북) 톤을 얹었습니다. 홈 첫 화면에 실제 DB의 '최근 검증 가격' 표를 노출해 서비스의 핵심 가치를 바로 보여줍니다.

| 구성 | 오픈소스 | 라이선스 | 위치 |
|---|---|---|---|
| 본문·한중일 글꼴 | [Pretendard](https://github.com/orioncactus/pretendard) | SIL OFL 1.1 | `static/vendor/pretendard/` |
| 영문 제목 글꼴 | [Source Serif 4](https://github.com/adobe-fonts/source-serif) (Fontsource 배포본) | SIL OFL 1.1 | `static/vendor/source-serif-4/` |
| 아이콘 | [Lucide](https://github.com/lucide-icons/lucide) | ISC | `static/vendor/lucide/icons/` (`icon()` 템플릿 함수로 인라인) |
| 지도 | [Leaflet](https://github.com/Leaflet/Leaflet) + OpenStreetMap 타일 | BSD-2 / ODbL | `static/vendor/leaflet/` |

외부 CDN 없이 모두 프로젝트 안에 포함되어 있습니다(지도 타일만 OpenStreetMap에서 로드). 각 폴더에 라이선스 파일이 있습니다.
