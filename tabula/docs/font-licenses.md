# 웹 글꼴 카탈로그와 라이선스 근거

검증일: 2026-10-05. 카탈로그는 `src/web-font-catalog.js`, 갱신 도구는 `tools/update-font-catalog.mjs`입니다. 운영 중 전체 글꼴을 내려받지 않고, 등록된 가족의 글꼴만 필요할 때 요청합니다. 이 문서는 글꼴 목록과 배포 조건의 확인 기록입니다.

## 목록의 범위

| 구분 | 가족 수 | 근거 |
|---|---:|---|
| Google Fonts 공식 제공 목록 | 1,950 | [공식 메타데이터](https://fonts.google.com/metadata/fonts) |
| Google Fonts 등록 | 1,948 | Noto Emoji·Noto Color Emoji 2개 제외. Material 아이콘 계열은 수집 규칙에서 제외 |
| 원배포처에서 확인한 추가 웹 글꼴 | 43 | 아래 한국어/CDN 목록 |
| 합계 | **1,991** | 가족 이름의 대소문자를 무시한 중복 0 |
| 한국어 지원 | **68** | Google 38 + 추가 CDN 30. `scripts.includes('korean')`으로 집계 |
| Cafe24 | 36 | 한국어 23 + 영문 장식 글꼴 13 |

Google Fonts 저장소 스냅샷은 [9710da1eacb3be272583c3224dcb70f9da6eadbb](https://github.com/google/fonts/tree/9710da1eacb3be272583c3224dcb70f9da6eadbb)입니다. 전체 트리 23,736개 항목이며 `truncated:false`를 확인했습니다. API 키를 사용하지 않습니다. 최신 서비스 메타데이터가 저장소보다 앞서는 경우가 있으므로 둘의 시점을 같다고 가정하지 않습니다. 메타데이터·트리 SHA-256, 정확한 가족 수, 배포 주소는 `WEB_FONT_CATALOG_META`에 기록합니다.

각 Google 가족은 해당 저장소 디렉터리의 실제 라이선스 파일을 연결합니다. 당시 디렉터리 또는 라이선스 파일이 없는 다음 10가족은 Google의 `download/list?family=...` 응답 안에 포함된 라이선스 전문을 직접 확인해 `licenseText`로 보존했습니다. 응답의 바이너리 다운로드 링크를 따라가지 않습니다.

- Edu NSW ACT Cursive / Edu NSW ACT Hand Pre
- Edu QLD Hand / Edu SA Hand
- Edu VIC WA NT Hand / Edu VIC WA NT Hand Pre
- Kumar One Outline / M PLUS Rounded 1c / Playwrite NZ Basic Guides / Tinos

현재 라이선스 분포는 OFL-1.1 1,949가족, Apache-2.0 35가족, Ubuntu Font Licence 1.0 5가족, 여기어때 자체 라이선스 2가족입니다. 모든 무료 글꼴에 같은 조건이 적용된다고 표시하지 않습니다.

## 사용자가 제시한 한국어 글꼴 목록 확인

[사용자 참고 글](https://flowworks.io/blog/free-commercial-fonts)은 조사할 목록으로만 사용했습니다. 실제 등록과 고지는 아래 원배포 자료에 근거합니다.

| 글꼴/그룹 | 등록 내용 | 버전·배포 주소와 원문 조건 |
|---|---|---|
| Pretendard | 100–900, 9굵기 | [원저장소](https://github.com/orioncactus/pretendard), [1.3.9 동적 서브셋 CSS](https://cdn.jsdelivr.net/gh/orioncactus/pretendard@v1.3.9/dist/web/static/pretendard-dynamic-subset.min.css), [라이선스](https://github.com/orioncactus/pretendard/blob/v1.3.9/LICENSE) |
| SUIT | 100–900, 9굵기 | [공식 안내](https://sun.fo/suit/), [2.0.5 CSS](https://cdn.jsdelivr.net/gh/sun-typeface/SUIT@v2.0.5/fonts/static/woff2/SUIT.css), [라이선스](https://github.com/sun-typeface/SUIT/blob/v2.0.5/LICENSE) |
| Wanted Sans | 400–1000, 7굵기 | [공식 웹 배포 문서](https://github.com/wanteddev/wanted-sans/blob/v1.0.3/packages/wanted-sans/documentation/webfonts/README.md), [1.0.3 CSS](https://cdn.jsdelivr.net/gh/wanteddev/wanted-sans@v1.0.3/packages/wanted-sans/fonts/webfonts/static/split/WantedSans.min.css), [라이선스](https://github.com/wanteddev/wanted-sans/blob/v1.0.3/OFL.txt) |
| IBM Plex Sans KR | Google Fonts, 100–700 | [공식 가족](https://fonts.google.com/specimen/IBM+Plex+Sans+KR) |
| NanumSquareNeo | 300·400·700·800·900 | [NAVER 배포처](https://campaign.naver.com/nanumsquare_neo/), [공식 CSS](https://hangeul.pstatic.net/hangeul_static/css/nanum-square-neo.css), [NAVER 라이선스 전문](https://help.naver.com/service/30016/contents/18088?osType=PC) |
| Gmarket Sans | 300·500·700 | [공식 배포와 OFL 안내](https://corp.gmarket.com/fonts/), 공식 사이트 CSS가 참조하는 script.gmarket.com의 WOFF2를 HTTPS로 사용 |
| S-Core Dream | 자동 웹 배포에 미등록 | [원배포처](https://s-core.co.kr/company/font/)는 원본 OTF ZIP을 제공하고 파일 수정·변환을 제한. 원본과 동일한 허용된 웹 배포 경로를 확인하지 못해 제3자 변환 WOFF를 임의 등록하지 않음. 설치된 원본 글꼴 사용을 제한하는 것은 아님 |
| Cafe24 Ssurround | 써라운드·에어 포함 Cafe24 36가족 | [Cafe24 공식 페이지](https://fonts.cafe24.com/), [한국어 목록](https://img.cafe24.com/csdstatic/freefonts/data/fonts_ko.json), [영문 목록](https://img.cafe24.com/csdstatic/freefonts/data/fonts_en.json) |
| Jalnan | 잘난체2 / 잘난체 고딕 | [공식 배포처](https://gccompany.co.kr/font), [공식 사용 가이드](https://drive.google.com/file/d/1TFGKe-QLRnx2IHpJO3ZH3pVUYjQwMx7L/view), 2023-09-27판. 옛 잘난체와 새 잘난체2를 같은 글꼴로 대체하지 않음 |
| Baemin | Google Fonts의 주아·도현·기랑해랑·연성 | [배민 공식 배포 안내](https://www.woowahan.com/fonts), 가족별 Google 원저장소 라이선스 연결 |
| Gowun | 고운돋움 400 / 고운바탕 400·700 | [고운돋움](https://fonts.google.com/specimen/Gowun+Dodum), [고운바탕](https://fonts.google.com/specimen/Gowun+Batang) |

Pretendard, SUIT, Wanted Sans는 위의 명시된 릴리스 버전을 고정합니다. NAVER·Gmarket·Cafe24·여기어때는 원배포처가 제공하는 실제 자산 경로를 사용하며, 버전이 없는 주소를 불변이라고 주장하지 않습니다. Cafe24 경로에 포함된 v1.0/v1.1/v2.0 등도 그대로 보존합니다. 제공자가 자산을 교체할 가능성은 남습니다.

## 저작권과 내보내기 고지

OFL은 글꼴 자체만 판매하는 행위를 허용하지 않으며, 재배포·임베딩할 때 저작권과 라이선스 고지가 필요합니다. Apache·Ubuntu·여기어때 라이선스는 각 원문 조건을 따릅니다. 글꼴로 만든 일반 문서와 글꼴 프로그램 자체의 재배포는 구별합니다. [SIL 공식 라이선스 전문](https://openfontlicense.org/open-font-license-official-text/)

카탈로그 모든 가족에는 `rawLicenseUrl` 또는 `licenseText`가 있습니다. Google의 10개 예외 가족, NAVER, Gmarket, Cafe24, 여기어때는 원문 텍스트를 함께 제공합니다. `licenseUrl`은 사용자가 확인할 원배포 자료입니다. HTML/SVG/PDF 등으로 글꼴 프로그램을 포함하는 내보내기는 이 고지를 함께 보존해야 합니다. 글꼴 이름만 기록한 XLSX에는 운영체제나 Excel에서 해당 글꼴을 사용할 수 있는지에 따른 차이가 남습니다.

- **NAVER:** 공식 안내의 영문 원문과 번역을 보존했습니다. 본문에 NanumSquareNeo가 예약 글꼴 이름으로 명시되어 있습니다.
- **Gmarket:** 공식 페이지의 OFL 허용 문구와 공식 WOFF `name` 테이블의 실제 저작권 `Copyright © 2019 eBay Korea Co., Ltd. All rights reserved.`를 확인했습니다. 현재 사이트 이름으로 저작권자를 임의 변경하지 않습니다.
- **Cafe24:** 공식 사이트의 공통 라이선스 안내는 제작한 모든 글꼴의 상업 이용·웹폰트·소프트웨어 임베딩·재배포를 허용하고 OFL을 적용합니다. 공식 써라운드 ZIP의 `License-Ssurround.pdf`도 대조했습니다. 등록한 36가족 각각의 공식 WOFF `name` 테이블에서 저작권 문구를 읽어 보존하고, 공통 안내와 SIL 1.1 전문을 함께 제공합니다. 연도·표기 차이를 한 문자열로 덮어쓰지 않았습니다. Angbanana의 원본에는 `\\250\\317` 형태의 표기가 있어 원본 문자열을 그대로 보존했습니다.
- **여기어때:** 공식 PDF의 License 페이지 전문을 `licenseText`로 포함했습니다. 수정·개작 재배포 금지, 상표권 등록 제한 등 자체 조건이 있으므로 OFL이라고 표시하지 않습니다. 공식 사이트에 연결된 Framer 자산을 그대로 사용하며 글꼴 바이너리를 변환하지 않습니다.

## 카탈로그 필드

`WEB_FONT_CATALOG`는 다음 구조의 객체 배열입니다.

```js
{
  family: 'IBM Plex Sans KR',
  label: 'IBM Plex Sans KR',
  aliases: ['IBM 플렉스 산스 KR'],
  source: 'google', // 또는 'cdn'
  category: 'sans-serif',
  scripts: ['korean', 'latin', 'latin-ext', 'cyrillic'],
  weights: [100, 200, 300, 400, 500, 600, 700],
  styles: ['normal'],
  variants: ['100', '200', '300', '400', '500', '600', '700'],
  license: 'OFL-1.1',
  licenseUrl: '원배포 라이선스',
  rawLicenseUrl: '라이선스 원문 텍스트 주소',
  homepage: '원배포 홈페이지'
}
```

CDN 글꼴은 `cssUrl` 또는 `faces:[{url,weight,style}]`를 가집니다. 지원 굵기는 CSS가 선언한 실제 값입니다. Cafe24의 각각 분리된 굵기 가족은 원배포 CSS처럼 400으로 등록합니다. Pretendard와 Wanted Sans의 동적 서브셋은 원 CSS의 unicode-range를 사용합니다.

`variants`는 Google 메타데이터의 실제 키입니다. Bona Nova·Cardo·Neuton 등 14가족은 굵기와 기울임의 모든 조합을 제공하지 않습니다. `weights × styles`의 단순 곱으로 CSS2 요청을 만들면 HTTP 400이 발생할 수 있으므로 실제 `variants`만 요청해야 합니다. italic만 있는 Molle도 normal을 임의 요청하지 않습니다. Wanted Sans는 1000 굵기도 지원합니다.

## 네트워크 범위와 검증

실행 중 글꼴 요청에 필요한 origin은 다음 7개로 제한합니다.

- https://fonts.googleapis.com
- https://fonts.gstatic.com
- https://cdn.jsdelivr.net
- https://hangeul.pstatic.net
- https://script.gmarket.com
- https://img.cafe24.com
- https://framerusercontent.com

라이선스 원문 수집용 raw.githubusercontent.com과 메타데이터 갱신용 fonts.google.com·api.github.com은 글꼴 바이너리 origin과 구별합니다. 셀 값, 문서 이름, 도형 텍스트를 URL의 `text=` 등에 넣지 않습니다. 글꼴 이름만 요청합니다.

초기 검증은 CDN CSS 3개와 직접 face 46개(총 49요청)의 정상 응답을 확인했습니다. 별도로 CSS가 참조하는 표본 글꼴도 HEAD로 확인했습니다. Cafe24는 HEAD에 CORS 헤더가 없지만 실제 GET에는 `Access-Control-Allow-Origin: *`가 있으므로, 36개 모두 GET Range 1바이트 응답 헤더로 재확인했습니다. 글꼴 전체 파일은 이 목록 수집 과정에서 저장하지 않았습니다. 추가로 한국어 전체 38가족, 비대칭 굵기·기울임, 저장소 누락 예외, 주요 문자권을 포함한 Google 60가족의 실제 CSS2 응답도 모두 정상임을 확인했습니다. 같은 캐시로 생성한 카탈로그 SHA-256이 원본과 일치합니다. 브라우저 적용·내보내기 동작은 별도 UI 회귀 검증 대상입니다.

## 갱신

```powershell
node tools/update-font-catalog.mjs --cache-dir D:/Codex/Temp/wixel-font-catalog --verify-network
```

Google 공식 목록과 저장소 트리를 다시 수집하고, 확인되지 않은 라이선스·불완전한 트리·지원하지 않는 variant 형식이 있으면 실패합니다. `--verify-network`는 curated CSS/face의 HTTP와 CORS를 검사합니다. CORS가 HEAD에서 생략된 경우에만 GET Range 응답 헤더를 확인하며 글꼴 본문을 저장하지 않습니다. 실패할 경우 새 카탈로그를 쓰지 않습니다.

```powershell
node tools/update-font-catalog.mjs --cache-dir D:/Codex/Temp/wixel-font-catalog --offline --date 2026-10-05 --output D:/Codex/Temp/web-font-catalog-rebuilt.js
```

`--offline`은 이미 검증한 메타데이터와 라이선스 manifest 캐시만 사용합니다. 같은 캐시·날짜·도구로 동일한 파일을 재생성할 수 있습니다. `--reuse-cache`는 유지보수자가 명시적으로 기존 캐시를 우선 사용할 때 쓰며 신규 자료 수집과 구별합니다. 일반 앱 빌드는 수집 도구를 실행하지 않습니다.

추가 CDN 목록은 도구 안의 정적으로 검증된 `curated` 배열입니다. 목록 자동 갱신 시 출처 불명의 URL을 추가하지 않습니다. 새 CDN 가족을 추가하거나 버전을 변경할 때는 원문 조건, 실제 CSS의 가족 이름·굵기, HTTPS·CORS, 내보내기 고지를 다시 확인해야 합니다.
