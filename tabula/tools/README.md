# 검증 도구 (tools/)

`npm test` 외에 브라우저 동작과 엑셀 파일 회귀를 확인하는 스크립트입니다.
실제 파일을 비교할 때는 **수정 전에 기준 결과를 저장하고, 수정 후 같은 파일과 명령으로 다시 측정**하세요.
차트·피벗 옵션·화면 표시의 검증 범위는 [`14_차트와표시검증.md`](../docs/codex/14_차트와표시검증.md), 최신 단축키 결과는 [`13_단축키점검.md`](../docs/codex/13_단축키점검.md), 공개 배포는 [`12_공개배포와운영검증.md`](../docs/codex/12_공개배포와운영검증.md), 셀 스타일·지우기의 실행 결과와 미검증 범위는 [`11_셀스타일과지우기.md`](../docs/codex/11_셀스타일과지우기.md), 기본 3.0 검증은 [`09_WIXEL3_구현과검증.md`](../docs/codex/09_WIXEL3_구현과검증.md), 원본 인계 검증은 [`06_검증결과.md`](../docs/codex/06_검증결과.md)에 기록합니다.

## 명시적 파일 저장 창 검사

```powershell
node tools/save-dialog.mjs
node tools/server-save-dialog.mjs
```

`save-dialog.mjs`는 격리 브라우저에서 파일 선택 API를 메모리 핸들로 바꾸고 저장창 호출·실제 XLSX 내용·취소·권한/쓰기 실패·미지원 환경의 확인 후 다운로드를 검사합니다. OS 창 외형과 운영체제 권한 확인창 자체는 수동 검사 대상입니다. `server-save-dialog.mjs`는 로컬 소스 서버와 가상 온라인 API로 이름·폴더 선택, 기존 파일 덮어쓰기의 revision 확인, 문서 전환 보호를 검사합니다. 실제 사용자 파일이나 온라인 보관함에는 쓰지 않습니다.

## 준비 (Windows PowerShell)

기존 체크아웃 경로는 유지합니다. 새 작업은 `D:\Codex\Workspaces`, 임시 파일은 `D:\Codex\Temp`,
캐시는 `D:\Codex\Caches`에 둡니다. 아래 경로는 현재 체크아웃 예시입니다.

```powershell
Set-Location D:\위셀\wixel-3\tabula
New-Item -ItemType Directory -Force D:\Codex\Temp, D:\Codex\Caches\npm, D:\Codex\Caches\ms-playwright | Out-Null
$env:TEMP = 'D:\Codex\Temp'
$env:TMP = 'D:\Codex\Temp'
$env:npm_config_cache = 'D:\Codex\Caches\npm'
$env:PLAYWRIGHT_BROWSERS_PATH = 'D:\Codex\Caches\ms-playwright'
npm install --no-save --package-lock=false playwright
node node_modules/playwright/cli.js install chromium
```

Playwright는 검증 도구용입니다. 앱의 `package.json`과 잠금 파일에 의존성을 추가하지 마세요.
`check.mjs`는 Node만 사용하므로 Playwright나 실행 중인 서버가 필요 없습니다.

서버용 터미널에서 다음을 실행합니다. 테스트용 문서 저장 경로를 사용하며, 종료는 `Ctrl+C`입니다.

```powershell
Set-Location D:\위셀\wixel-3\tabula
$env:HOST = '127.0.0.1'
$env:PORT = '5178'
$env:TABULA_DATA = 'D:\Codex\Temp\wixel-server-data'
npm start
```

검증용 터미널에서는 브라우저 경로를 같은 값으로 지정하고 스크립트를 실행합니다.

```powershell
Set-Location D:\위셀\wixel-3\tabula
$env:PLAYWRIGHT_BROWSERS_PATH = 'D:\Codex\Caches\ms-playwright'
$env:WIXEL_URL = 'http://127.0.0.1:5178/'
node tools/smoke.mjs
$LASTEXITCODE
```

이미 외부 폴더에 설치한 Playwright를 재사용할 수도 있습니다. `PLAYWRIGHT_MODULE`은
`import()`가 읽을 수 있는 **파일 URL**을 사용합니다. 아래는 해당 위치에 설치되어 있을 때의 예입니다.

```powershell
$env:PLAYWRIGHT_MODULE = 'file:///D:/Codex/Temp/wixel-tools/node_modules/playwright/index.mjs'
```

현재 폴더의 `node_modules`를 쓰려면 `Remove-Item Env:PLAYWRIGHT_MODULE -ErrorAction SilentlyContinue`로 이 설정을 해제합니다.
서버 포트를 바꿨다면 `WIXEL_URL`도 맞추세요. 브라우저 스크립트는 생성한 테스트 브라우저에서 데이터를 초기화하고 명령을 실행하므로 별도 테스트 서버를 사용합니다.
스모크·키보드·파일 비교 브라우저 도구는 초기화 스크립트에서 `window.TABULA_STATIC = true`를 설정해
서버 문서 자동 저장·복원으로 검사 데이터가 바뀌는 것을 막습니다. 도구의 성공은 서버 API 인증·보안 검증을 의미하지 않습니다.
`ui-regressions.mjs`도 별도 브라우저에서 합성 문서를 사용합니다. 인증 화면 시나리오는 API 응답을 모의하며, 실제 서버 API는 `npm test`의 서버 테스트에서 검증합니다.

### Git Bash 사용 시

Windows Git Bash에서도 같은 D: 경로를 사용합니다. 아래 환경 변수는 명령을 실행하는 각 터미널에 적용합니다.

```bash
cd /d/위셀/wixel-3/tabula
mkdir -p /d/Codex/Temp /d/Codex/Caches/npm /d/Codex/Caches/ms-playwright
export TEMP='D:/Codex/Temp' TMP='D:/Codex/Temp'
export npm_config_cache='D:/Codex/Caches/npm'
export PLAYWRIGHT_BROWSERS_PATH='D:/Codex/Caches/ms-playwright'
export WIXEL_URL='http://127.0.0.1:5178/'
npm install --no-save --package-lock=false playwright
node node_modules/playwright/cli.js install chromium
```

서버용 터미널에서는 같은 폴더에서 다음 명령을 실행하고, 다른 터미널에서 `node tools/smoke.mjs`를 실행합니다.

```bash
HOST=127.0.0.1 PORT=5178 TABULA_DATA='D:/Codex/Temp/wixel-server-data' npm start
```

외부 Playwright 모듈은 `export PLAYWRIGHT_MODULE='file:///D:/Codex/Temp/wixel-tools/node_modules/playwright/index.mjs'`처럼 지정합니다.

## 스크립트와 판정

| 파일 | 하는 일 | 사용법 | 성공 기준 |
|------|---------|--------|-----------|
| `smoke.mjs` | 샘플 데이터에서 실행 가능한 앱 명령을 호출해 콘솔 오류·예외 확인. 파일 선택 등 일부 명령 제외 | `node tools/smoke.mjs` | `bad 0` 및 종료 코드 0 |
| `keys.mjs` | 직접 단축키의 실제 선택·서식·계산·QAT 결과 확인 | `node tools/keys.mjs` | 요약의 `bad`·`pageErrors`가 모두 0, 종료 코드 0 |
| `keytips.mjs` | 리본 키팁 등록표·전 중간 입력 경로·키 순서 실제 동작 확인 | `node tools/keytips.mjs` | `failed 0`, 종료 코드 0; 등록·경로·동작 범위 별도 출력 |
| `ui-regressions.mjs` | 비연속 선택·대화상자·메뉴·피벗 이동·인증 재시도의 실제 브라우저 회귀 검사 | `node tools/ui-regressions.mjs` | 모든 항목 통과, 페이지 오류 0, 종료 코드 0 |
| `check.mjs` | Node에서 수식을 다시 계산해 파일 저장값과 비교하고 xlsx 저장/읽기 왕복 검사 | `node tools/check.mjs 파일.xlsx` | `formula mismatches: 0`, `roundtrip issues: 0`, 종료 코드 0 |
| `brcheck.mjs` | 브라우저에서 파일을 열고 피벗을 다시 그린 뒤 수식 비교 및 로딩 시간 기록 | `node tools/brcheck.mjs 파일.xlsx` | 브라우저 수식 불일치·오류 없음 및 종료 코드 0 |
| `pvcmp.mjs` | 피벗 영역별 파일 저장값과 WIXEL 재계산 결과 비교 | `node tools/pvcmp.mjs 파일.xlsx` | `pivots N differing 0`, 종료 코드 0 |

- `check.mjs`·`pvcmp.mjs`: 종료 코드 **0** = 검사상 차이 없음, **1** = 비교 차이, **2** = 입력 문제 또는 검사 중 실행 오류. 비정상 종료만 보고 차이의 원인을 단정하지 말고 출력도 확인하세요.
- `smoke.mjs`·`keys.mjs`·`brcheck.mjs`: 검출한 오류·실패는 종료 코드 **1**. `brcheck.mjs`의 입력 문제는 **2**입니다. 모듈 설치·브라우저 시작·서버 연결 등 실행 자체의 실패도 비정상 종료이므로 결과 로그를 함께 확인하세요.
- `ui-regressions.mjs`도 검사 실패·페이지 오류가 있으면 종료 코드 **1**입니다. `WIXEL_UI_SCREENSHOT`에 절대 PNG 경로를 지정하면 피벗 이동 창을 캡처합니다.
- PowerShell은 명령 직후 `$LASTEXITCODE`, Bash는 `$?`로 종료 코드를 확인합니다. 뒤에 실행한 명령의 코드로 덮이지 않게 바로 기록하세요.
- `check.mjs`·`pvcmp.mjs`는 **xlsx/xlsm/xlsb**만 받습니다. 옛 `.xls`(BIFF8)는 `brcheck.mjs`로 확인하세요.
- `TODAY/NOW/RAND` 같은 수식의 날짜·난수 차이는 **자동 제외되지 않습니다**. 실제 차이와 구분해 수동 검토 근거를 남기세요.
- `pivots 0 differing 0`은 피벗이 없는 파일을 뜻하므로 피벗 기능 검증 성공으로 세지 마세요.
- 스모크와 왕복 검사는 실제 Excel에서의 표시·복구 메시지 여부를 보장하지 않습니다. Excel 365 확인란 실검증은 별도입니다.

## 회귀 비교 방법 (PowerShell)

실제 업무 파일은 예를 들어 `D:\Codex\Temp\wixel-samples`에 둡니다. 파일과 셀 값이 담긴 로그는
저장소·전달 ZIP에 넣지 마세요. 아래 예시는 `tabula` 폴더에서 실행합니다.

1. **수정 전**에 입력 파일 목록과 해시를 한 번 저장합니다. 수정 후에는 이 목록을 다시 만들지 않습니다.

```powershell
$sampleRoot = 'D:\Codex\Temp\wixel-samples'
$reportRoot = 'D:\Codex\Temp\wixel-regression'
New-Item -ItemType Directory -Force $reportRoot | Out-Null
$sampleFiles = @(Get-ChildItem -LiteralPath $sampleRoot -File |
    Where-Object { $_.Extension.ToLowerInvariant() -in '.xlsx', '.xlsm', '.xlsb' } |
    Sort-Object FullName)
if ($sampleFiles.Count -eq 0) { throw '검증할 엑셀 파일이 없습니다.' }
$sampleFiles | ForEach-Object {
    [pscustomobject]@{ Path = $_.FullName; SHA256 = (Get-FileHash -LiteralPath $_.FullName -Algorithm SHA256).Hash }
} | Export-Csv -LiteralPath (Join-Path $reportRoot 'inputs.csv') -NoTypeInformation -Encoding UTF8
```

2. 다음 블록을 `$phase = 'before'`로 실행합니다. 그 뒤 코드를 수정하고 같은 블록을 `$phase = 'after'`로 실행합니다.
   파일 해시가 달라지면 중단하므로 같은 입력으로 비교할 수 있습니다. 각 실행의 전체 출력과 종료 코드를 저장합니다.

```powershell
$phase = 'before' # 수정 후에는 'after'
$reportRoot = 'D:\Codex\Temp\wixel-regression'
$logPath = Join-Path $reportRoot ($phase + '.txt')
Set-Content -LiteralPath $logPath -Value "phase $phase" -Encoding UTF8
foreach ($item in (Import-Csv -LiteralPath (Join-Path $reportRoot 'inputs.csv'))) {
    if ((Get-FileHash -LiteralPath $item.Path -Algorithm SHA256).Hash -ne $item.SHA256) {
        throw "입력 파일이 변경되었습니다: $($item.Path)"
    }
    Add-Content -LiteralPath $logPath -Value "== $($item.Path)" -Encoding UTF8
    foreach ($toolName in 'check', 'pvcmp') {
        $output = & node "tools/$toolName.mjs" $item.Path 2>&1
        $resultCode = $LASTEXITCODE
        Add-Content -LiteralPath $logPath -Value "tool $toolName" -Encoding UTF8
        Add-Content -LiteralPath $logPath -Value ($output -join [Environment]::NewLine) -Encoding UTF8
        Add-Content -LiteralPath $logPath -Value "exit $resultCode" -Encoding UTF8
    }
}
```

3. `before.txt`와 `after.txt`를 파일별로 비교합니다. 다음 명령으로 주요 판정 줄을 볼 수 있습니다.
   총합만 같아도 특정 파일이 나빠질 수 있으므로 개별 파일·피벗의 차이와 실행 오류를 확인하세요.

```powershell
Select-String -LiteralPath D:\Codex\Temp\wixel-regression\before.txt, D:\Codex\Temp\wixel-regression\after.txt `
    -Pattern '^== ', '^tool ', '^formula mismatches:', '^roundtrip issues:', '^pivots ', '^exit '
```

나빠진 파일이 있으면 원인을 해결하기 전에는 커밋하지 않습니다. 이번 결과, 변동 수식 등 검토한 예외,
사용한 기준 커밋과 Node 버전을 함께 기록하세요. 실제 사용자 파일이 없다면 생성 샘플 검사와 미실행 범위를 명확히 구분합니다.

## 이전 인계의 실제 파일 측정 기록

아래는 **2026-10-01, 원본 인계 `f6cbc6e`에서 전달받은 기록**이며 이번 검증에서 재측정한 결과가 아닙니다.
입력 파일은 사용자 개인 자료로 저장소에 없습니다.

- 피벗이 있는 실제 파일 50개: 대부분 `differing 0`으로 기록됨.
- 남은 차이(다른 피벗 수 / 전체): g2.xlsb 3/30, m3.xlsb 2/19, s3.xlsx 1/1, v37.xlsx 6/56, w3.xlsb 1/8, x3.xlsx 1/12, y60.xlsb 27/62, y61.xlsb 5/10, y62.xlsb 3/8, y63.xlsb 4/5.
- 분석 대상은 `docs/codex/02_보완항목내역서.md` P1-1, 이번 검증 실적은 `docs/codex/06_검증결과.md`를 참고하세요.


## WIXEL 3.0 추가 검사

모든 브라우저 검사는 별도 테스트 문서로 실행하며 실제 사용 중인 탭은 변경하지 않습니다. 실행 전에 서버를 시작하고 `WIXEL_URL`을 지정하세요. Playwright가 프로젝트 밖에 설치되어 있으면 `PLAYWRIGHT_MODULE`에 해당 `index.mjs`의 file URL을 지정합니다.

| 명령 | 검사 범위 |
|---|---|
| `node tools/grid-rendering.mjs` | 합성 시트의 24개 테마·DPR·앱 배율 조건에서 PNG 실제 픽셀 두께, 이중선·채우기 경계, 숨김·고정·스크롤·클릭 좌표·F2 편집 정렬 |
| `node tools/chart-ui.mjs` | 콤보 계열별 종류·주축/보조축, 즉시 미리보기·취소·확정·실행 취소, 가져온/숨긴 계열, 두 견본과 일반 차트 전환 |
| `node tools/pivot-options.mjs` | 합성 XLSX 파일 열기 UI, 오류 표시 옵션·계산 결과·재저장 왕복; 사례 수와 assertion 수를 별도 출력 |
| `node tools/keytips.mjs` | 키팁 등록 중복·접두 충돌·명령 존재, 전 중간 입력 경로, 대표 명령의 실제 결과, 작은 화면 패널·QAT 배치 |
| `node tools/format-regressions.mjs` | 차트 축/계열 연속 변경, 3D 회전·깊이, 도형 채우기·효과·텍스트와 페이지 오류 |
| `node tools/security-regressions.mjs` | CSP 우회 테스트 환경에서 HTML/SVG 공격 차단, 인쇄·링크·정상 차트/도형 표현 |
| `node tools/vault-ui-integration.mjs` | 로컬 workerd의 보관함 연결·저장·다른 기기 복원·게시 관리 UI |
| `node tools/wixel3-ui.mjs` | 시작·저장 화면, Ctrl+Space, QAT, 개인 보관함 충돌, Google Sheets, 피벗과 셀 서식 |
| `node tools/cell-style-regressions.mjs` | 새 스타일의 포함 항목·서식 편집·취소, 수정·복제·삭제·실행 취소, 셀/행/열/블록 연결, XLSX/WIXEL 스타일 병합과 이름 충돌, 표준 스타일 및 시트 보호 |
| `node tools/clear-ui.mjs` | 지우기 메뉴·우클릭의 하이퍼링크 지우기/제거, 기존 지우기 명령, 비연속 선택·보호·실행 취소, 조건부 서식과 값·메모·그림 보존, 앞자리 0·문자 수식·숫자 타입 유지 |
| `node --test cloudflare/backend.test.js` | 개인 보관함 격리·버전·청크·게시 API |
| `node cloudflare/integration.mjs` | 실제 로컬 workerd API; 별도 Cloudflare dev 서버 필요 |

`format-regressions.mjs`는 저장소 `.local/`에 차트와 도형 화면을 캡처합니다. Cloudflare 설정·실행 환경과 한도 검사 명령은 `cloudflare/README.md`를 확인하세요. 실제 공개 배포 후의 동작은 로컬 통과와 별도로 확인해야 합니다.

`cell-style-regressions.mjs`와 `clear-ui.mjs`는 별도 브라우저에서 **스크립트가 만든 합성 문서만** 검사합니다. 스타일 병합용 XLSX/WIXEL도 검사 중 생성하며 사용자 업무 파일을 읽지 않습니다. 모든 항목 통과·페이지 오류 없음·종료 코드 0을 확인하고, 실제 파일 회귀와 실제 Excel 앱 검증은 별도로 기록하세요.

`cell-style-regressions.mjs`의 스타일·병합 대화상자 캡처는 기본적으로 `D:\Codex\Temp\wixel3-cell-style.png`, `D:\Codex\Temp\wixel3-style-merge.png`에 저장합니다. `WIXEL_STYLE_SCREENSHOT`, `WIXEL_STYLE_MERGE_SCREENSHOT`으로 다른 D: 절대 경로를 지정할 수 있습니다. 최종 실행 수치는 [`11_셀스타일과지우기.md`](../docs/codex/11_셀스타일과지우기.md)를 확인하세요.

### 차트·피벗·키팁 검사의 범위

`chart-ui.mjs`와 `pivot-options.mjs`는 새 브라우저에서 스크립트가 만든 합성 문서만 사용합니다. `WIXEL_URL`로 소스 서버 또는 검사할 번들을 지정합니다. `chart-ui.mjs`는 계열별 새 선택으로 세로 막대·꺾은선·영역 및 축 0/1을 검사하며 Excel의 모든 차트 조합을 보장하지 않습니다. `pivot-options.mjs`는 플래그와 문구의 조합을 검사하므로 실제 업무 피벗 전체의 계산 회귀를 대신하지 않습니다. 모든 항목 통과·페이지 오류 없음·종료 코드 0을 확인하세요.

차트 화면 캡처는 `D:\Codex\Temp\wixel3-combo-chart.png`가 기본이며 `WIXEL_CHART_SCREENSHOT`에 다른 D: 절대 PNG 경로를 지정할 수 있습니다. 차트 엔진 및 표준 XML 왕복은 `test/chart-combo.test.js`, `test/chart-combo-xlsx.test.js`가 별도로 확인합니다. 브라우저 미리보기 통과와 XLSX 왕복 통과는 서로 다른 검증입니다.

`keytips.mjs`는 매 시나리오를 별도 컨텍스트에서 실행하고 원격 쓰기 요청을 차단합니다. 요약의 등록 키 수·검사한 중간 경로 수·실제 동작 키 수를 구분해 기록하세요. 한글 물리 키와 AltGr 검사는 합성 KeyboardEvent를 포함하므로 OS의 실제 IME나 브라우저 예약 키까지 검증하는 것은 아닙니다. `keys.mjs`의 직접 단축키 검사와 함께 실행합니다.

`grid-rendering.mjs`는 합성 시트를 새 컨텍스트에 만들고 원격 쓰기를 차단합니다. 흰색·검정 테마, DPR 1·1.25·2, 앱 배율 50·75·100·125%의 24개 조건에서 PNG를 직접 읽어 선 두께를 검사합니다. `GRID_SCREENSHOTS`에 D: 절대 폴더를 지정하면 조건별 PNG를 저장합니다. `ok: true`, 페이지 오류·쓰기 요청 없음과 종료 코드 0을 확인하고 `cases`·`checks`를 별도로 기록하세요. 작은 합성 시트의 `renderMs`는 대형 문서 성능 수치가 아닙니다. 실제 모니터 전환·브라우저 자체 확대/축소와 모든 환경까지 검증하지는 않습니다. 최종 결과와 한계는 [14 문서](../docs/codex/14_차트와표시검증.md)를 확인하세요.

## 차트·그림·필터·설정 확장 검사

`chart-upgrade-ux.mjs`, `chart-map.mjs`, `picture-ux.mjs`, `interaction-ux.mjs`, `settings-ux.mjs`는 위 Playwright 설정과 `WIXEL_URL`을 사용합니다. 합성 데이터만 생성하며 쓰기 API를 차단합니다. `chart-ui.mjs`의 기존 콤보 회귀도 함께 실행하세요. 실행 결과와 검증 제한은 [15_Excel_UI확장검증.md](../docs/codex/15_Excel_UI확장검증.md)에 기록합니다. `update-map-data.mjs`는 검사 명령이 아니라 지도 공개 데이터의 재생성 도구이며 [16번 문서](../docs/codex/16_지도데이터.md)의 고정 버전·출처를 사용합니다.


## 실제 Excel와 복구·계산·접근성 회귀

합성 파일만 사용하는 Windows Excel 검증(Excel 설치 필요):

```powershell
node tools/excel-fixtures.mjs D:/Codex/Temp/wixel-production-excel
& tools/excel-interop.ps1 -FixtureRoot D:/Codex/Temp/wixel-production-excel
node tools/excel-roundtrip.mjs D:/Codex/Temp/wixel-production-excel
```

일반 열기에 실패한 파일을 복구하여 통과시키지 않습니다. 생성기 목록의 합성 XLSX만 읽고 별도 `excel-saved` 폴더에 저장합니다. 사용자 Excel 세션을 닫거나 업무 파일을 수정하지 않습니다.

추가 브라우저 검사: `calculation-trust-ux.mjs`(저장값/오류 표시·Undo), `picture-ux.mjs`(효과 포함), `grid-a11y.mjs`(Chromium 접근성 트리), `performance-regression.mjs`(합성 문서/CPU 제한/자료 보존).
성능 수치는 장치·자동 저장·브라우저 설정에 따라 달라지며 제품 전체의 점수나 SLA가 아닙니다.

로컬 Cloudflare workerd 복구 검사:

```powershell
$env:WIXEL_WORKER_URL='http://127.0.0.1:8787'
node cloudflare/recovery-integration.mjs
node tools/recovery-ux.mjs
```

두 도구는 localhost를 확인하고 별도 무작위 보관함을 사용하며 종료 시 합성 문서와 이력을 삭제합니다. 공개 배포 주소나 실제 보관함 키를 넣지 마세요.
결과 및 범위는 [신뢰성 검증 문서](../docs/codex/17_신뢰성_복구_실제Excel검증.md)를 참조하세요.

### 후속 신뢰성 회귀

- `node tools/excel-chart-gallery.mjs D:/Codex/Temp/wixel-gallery-excel`: 61종 합성 차트 생성. `excel-interop.ps1 -FixtureRoot`로 실제 Excel을 거친 뒤 `--verify D:/Codex/Temp/wixel-gallery-excel/excel-saved`로 다시 비교합니다.
- `node tools/date-system.mjs`: 1900/1904 문서의 표시·입력·파일 저장·Undo.
- `node tools/pivot-reliability.mjs`: 외부 시트 원본 수식 변경과 자동 새로 고침·Undo.
- `node tools/storage-atomic.mjs`: 분리된 브라우저 컨텍스트의 실제 IndexedDB에 고장을 주입하여 이전 대용량 저장본 보존을 검사합니다. 사용자 탭의 저장소를 사용하지 않습니다.

- `node tools/library-atomic.mjs`: loopback 소스 서버 전용, 두 탭과 실제 IndexedDB의 문서·이력·목록 동시성/고장 주입.
- `node tools/library-recovery.mjs`: 초기 복원 후 충돌 사본·원본 재열기·실패 후 재시도·저장 중 편집의 화면 흐름.
- `tools/ui-regressions.mjs`, `tools/wixel3-ui.mjs`는 내부 모듈을 직접 가져오는 항목이 있으므로 소스 서버에서 실행합니다. 배포본 검사는 `keys`, `keytips`, `chart-upgrade-ux`, `picture-ux`, `recovery-ux` 등 번들 호환 도구를 사용합니다.

- `node tools/excel-pivot-fixtures.mjs D:/Codex/Temp/wixel-pivot-excel`: 합성 피벗 8종 생성. `excel-interop.ps1 -FixtureRoot`는 Excel 새로 고침과 29개 기대값을 검사합니다. 이후 `--verify D:/Codex/Temp/wixel-pivot-excel/excel-saved`로 WIXEL 재읽기 비교합니다.


## 피벗 표시·팝업 접근키·한글 지연 입력

아래 도구는 위 Playwright 환경 변수와 `WIXEL_URL`을 사용합니다. 새 컨텍스트의 합성 문서만 사용하고 원격 쓰기를 차단합니다.

| 명령 | 검사 범위 |
|---|---|
| `node tools/pivot-display-options.mjs` | XLSX 8개 옵션 조합, 체크 상태·표시 행·취소·Undo/Redo·저장 왕복 |
| `node tools/pivot-classic-grid.mjs` | 표 안 필드 클릭 메뉴·실제 HTML 드래그·필드 창 연결·Σ 값 이동·보호 |
| `node tools/keytips-ime.mjs` | 키팁 후 지연 beforeinput/composition/input, 반복 눈금선 전환, 다음 정상 한글/영문 |
| `node tools/find-replace-access-keys.mjs` | Ctrl+H, Alt+A/D/N/E/T/C/F/I, 중첩 서식 창과 포커스·Undo |
| `node tools/dialog-access-keys.mjs` | 공통 대화상자·메뉴·하위 메뉴, 중복키·AltGr·동적/숨김/비활성 제어, 별도 팝업 |

`dialog-access-keys.mjs`의 공통 UI 단독 검사는 소스 모듈을 가져오므로 소스 서버에서 실행합니다. 다른 네 도구는 최종 번들에서도 실행합니다. `keytips-ime.mjs`는 Chromium에 합성 이벤트를 보내므로 Windows IME 자체를 검증한 결과와 구분해야 합니다. 상세 범위는 [19번 문서](../docs/codex/19_피벗표시와팝업접근키.md)를 확인하세요.

`pivot-classic-grid.mjs`는 기본 A7 배치를 검사합니다. `CLASSIC_TOP=0`을 지정하면 A1에 피벗을 두고 원본 데이터를 별도 합성 시트로 옮겨 같은 드래그 회귀를 실행합니다. `CLASSIC_SCREENSHOT`으로 D: PNG 경로를 지정하면 기본 화면과 드래그 중 띠 위치를 함께 저장합니다.


## 셀·행·열 우클릭 회귀

| 명령 | 범위 |
|---|---|
| `node tools/context-menu.mjs` | 실제 우클릭/Shift+F10/메뉴 키, 명시 접근키, 선택·치수·숨김·Undo, 보호/읽기 전용, 미니 도구, 작은 화면 |
| `node tools/context-mini-toolbar.mjs` | 공통 UI의 배치/포커스/접근키/연속 서식/닫힘. 소스 서버 전용 |
| `node tools/context-menu-safety.mjs` | 셀 밀기 보호, 잘라내기·붙여넣기 자료 보존, 변경 범위 권한, 치수 변경의 수식 저장값 유지 |

`WIXEL_URL`과 기존 Playwright 환경 변수를 사용한다. 분리된 컨텍스트의 합성 문서만 만들며 클립보드를 모의 구현하고 원격 쓰기를 차단한다. 사용자 업무 파일이나 열린 탭은 건드리지 않는다. `CONTEXT_MENU_SCREENSHOT`에 D: 절대 PNG 경로를 지정하면 메뉴를 캡처한다. 구체적 한도·결과는 [20번 문서](../docs/codex/20_셀행열_우클릭메뉴.md)를 확인한다.


## 차트 범례 항목 이름

`node tools/chart-legend.mjs`는 합성 피벗의 값 열만 선택한 일반 원형/3D 원형 생성, 이름 열을 포함한 선택, 피벗 차트 전용 경로, 기존 숫자 범위 차트, 단일 계열 범례와 지표 선택을 검사한다. `WIXEL_URL`로 소스 또는 배포용 번들을 지정한다. 새 컨텍스트에서만 실행하고 원격 쓰기를 차단한다. `CHART_LEGEND_SCREENSHOT`으로 D: 절대 PNG 경로를 지정할 수 있다. [21번 문서](../docs/codex/21_차트범례_항목이름.md)에서 결과와 자동 이름 보완의 범위를 확인한다.


## 대용량 문서 성능과 필터·개체 회귀

- `node tools/filter-performance.mjs`: 합성 5만 고유 항목의 가상화, 전체 선택·검색·End/Home·Undo와 다른 열의 사용자 지정·색상·상위 조건을 검사한다. 소스/번들 모두 지원한다.
- `node tools/object-rendering.mjs`: 합성 167개 차트·슬라이서·그림·도형의 화면 밖 생성 생략, 캐시·서식·선택·Undo·스크롤·틀 고정·원본 개수 보존을 검사한다. 소스/번들 모두 지원한다.
- `node tools/real-workbook-performance.mjs "D:/업무파일.xlsx" "D:/Codex/Temp/result.json"`: localhost 전용 실제 파일 열기·필터·적용·Undo·스크롤 성능 측정. 실제 파일은 커밋하지 않는다. 원격/API 요청을 차단하고 별도 브라우저 컨텍스트를 사용한다. 결과는 시간·개수·함수별 CPU 표본만 기록하며 셀값·수식·시트명·원본 프로파일을 출력하지 않는다.

실제 파일 측정은 `WIXEL_PERF_CPU=4`로 CPU 제한, `WIXEL_PERF_COLUMNS=6,7,8,9,10,11,12,13`으로 0부터 시작하는 열 번호, `WIXEL_PERF_MODES=existing,cleared`로 기존 필터/조건 해제 상태를 지정한다. 각 작업은 최대55초 watchdog을 사용한다. 사용자 탭이나 실제 보관함은 수정하지 않는다. `PLAYWRIGHT_MODULE`, `PLAYWRIGHT_BROWSERS_PATH`, `WIXEL_URL`은 위와 같다.


`WIXEL_PERF_RESTORE=1`을 함께 지정하면 원래의 필터 조건을 다시 적용하고, 실제 대용량 IndexedDB 자동 저장 완료 후 같은 격리 컨텍스트의 새 페이지에서 복원한다. 값·서식·필터·날짜 체계·개수의 일치 여부만 기록한다. 임시 프로필에서만 실행하므로 사용자 보관함을 변경하지 않는다.

`node tools/release-update.mjs`는 새 배포 알림의 간격·실패 처리·제목줄 버튼·명시 저장/취소를 합성 응답으로 검사한다. 문서를 자동 새로고침하거나 전송하지 않는지 확인하며 소스/번들 모두 지원한다.


## 리본 키팁 전수 연결과 보고서 셀 스타일

- `node tools/ribbon-keytip-coverage.mjs`: 일반 8개·상황별 7개 탭의 실제 DOM 조작/키팁 등록/배지/입력 포커스 일치, HJ·HH·HFC의 선택·적용·Undo·Esc, held Alt/F10, 한글 지연 조합 회귀를 검사한다. `WIXEL_KEYTIP_FILTER`로 검사 이름을 제한할 수 있다.
- `node tools/cell-style-presets.mjs`: 새 24개 스타일의 적용·숫자 형식 보존·Undo, 이름 중복 우선순위, 검색·키보드 이동과 좁은 화면의 갤러리를 검사한다.

기존 `keytips.mjs`, `keytips-ime.mjs`, 셀 스타일 검사를 함께 유지한다. `WIXEL_URL`, `PLAYWRIGHT_MODULE`, `PLAYWRIGHT_BROWSERS_PATH`를 사용하며 새 격리 컨텍스트의 합성 문서만 사용하고 원격 쓰기를 차단한다. 소스와 배포용 번들에서 실행할 수 있다. IME 검사는 DOM 이벤트 회귀이며 실제 Windows IME 드라이버의 모든 동작을 보장하는 검사는 아니다. [23번 문서](../docs/codex/23_리본키팁과스타일.md)에서 Excel 호환경로와 WIXEL 보충 키의 범위 및 검증 결과를 확인한다.


### 선·자유곡선과 도형 서식

- `node tools/shape-drawing.mjs`: 갤러리·그리기 완료/취소·점 편집·Undo·보호·선 서식·닫힌 경로를 합성 문서에서 검사합니다.
- `node tools/shape-text.mjs`: 긴 일반/리치 텍스트를 회전 0/90/270도, 줄바꿈 켬/끔, 배율 50/100/150%에서 검사합니다. 실제 글자 경계와 축소 맞춤, 패널 입력을 확인합니다.
- 두 도구는 `WIXEL_URL`로 소스 또는 빌드 주소를 선택하고 격리 브라우저 문서에서 실행합니다. 공개 API 쓰기를 차단합니다.
- 표준 XLSX 저장과 실제 Excel 왕복 결과·한계는 [도형 검증 문서](../docs/codex/24_자유곡선과도형서식.md)를 참고하세요.


### 차트 요소 직접 편집

`node tools/chart-direct-editing.mjs`는 격리된 합성 문서에서 제목·범례 이동/삭제, 선택 요소 서식 동기화, 계열·점 색, 원형 조각 분리와 확대·축소, 실행 취소, 보호·시트 전환을 검사합니다. `WIXEL_URL`로 소스/번들/공개 주소를 고르고 `WIXEL_CHART_DIRECT_FILTER`로 검사 이름을 제한할 수 있습니다. 외부 API 쓰기를 차단하며 실제 문서·사용자 탭에는 접근하지 않습니다. 동작·저장 범위는 [차트 직접 편집 문서](../docs/codex/25_차트요소직접편집.md)를 참고하세요.

### 온라인 그림 검색 회귀

`node tools/online-pictures.mjs`는 격리 브라우저와 합성 검색 응답으로 Creative Commons 기본 해제·라이선스 필터·다중 출처·검색 경합·중복·원본 삽입·실행 취소·시트 전환을 검사합니다. `WIXEL_URL`로 소스/배포 번들을 지정합니다. 외부 API는 모의 응답이며 실제 업무 문서와 사용자 브라우저 탭을 사용하지 않습니다. 실제 공개 API 연결 여부는 별도로 확인합니다.


### 페이지 레이아웃·SmartArt·붙여넣기·검토 회귀

- `node tools/smartart.mjs`: 20배치/8범주 갤러리, 텍스트·그림·Undo·보호·크기/대칭.
- `node tools/object-group.mjs`: 도형/그림 여러 개 선택, 그룹화/해제·변환·보호·선택 창.
- `node tools/page-layout.mjs`: 인쇄 미리보기·PDF 실제 다운로드, 병합 셀·개체 원위치·페이지 경계, 테마/너비/높이/배율.
- `node tools/paste-special.mjs`: 12모드, 연산·전치·링크·빈 셀·Alt키, 취소·원본 불변·문서 전환·클립보드 지연.
- `node tools/review-ux.mjs`: 편집 허용 범위 암호/취소·보호·Undo·다른 문서 안전성과 메모 표시.

기존 `WIXEL_URL`, `PLAYWRIGHT_MODULE`, `PLAYWRIGHT_BROWSERS_PATH`를 사용합니다. 새 격리 브라우저의 합성 자료만 사용하며 API 쓰기를 차단합니다. 소스와 실제 배포용 번들을 모두 검사하고 결과·지원 범위는 [27번 문서](../docs/codex/27_페이지레이아웃과편집.md)에 기록합니다. PDF 결과는 D: 임시 폴더에 생성하며 사용자 문서나 기존 브라우저 탭을 사용하지 않습니다.


## 모바일 작업 모드

로컬 서버 또는 배포 번들을 대상으로 `WIXEL_URL`을 지정해 실행합니다. 실제 사용자 문서를 사용하지 않는 합성 검사입니다.

```sh
node tools/mobile-work-mode.mjs
node tools/mobile-touch.mjs
node tools/mobile-popups.mjs
node tools/mobile-layout.mjs
```

`mobile-layout.mjs`의 전체 리본 카탈로그 검사는 소스 서버가 필요합니다. 번들에서는 `MOBILE_LAYOUT_CATALOG=0`으로 실제 앱 화면·모드 전환·가상키보드 검사만 실행합니다. 실제 기기의 IME, iOS/Android 브라우저 바 및 다운로드 UI는 에뮬레이션 결과와 별도로 확인해야 합니다. 자세한 범위는 `docs/codex/28_모바일작업모드.md`를 참고하세요.


## 홈 화면 아이콘

`WIXEL_URL`을 실행 중인 소스 서버 또는 배포 주소로 지정하고 `node tools/home-icons.mjs`를 실행합니다. 홈 화면 메타데이터, manifest 파싱, HTTP MIME, 아이콘 실제 디코드 크기·불투명도·안전 영역 및 로컬 파일과의 해시를 검사합니다. 이는 실제 휴대전화 홈 화면 설치 테스트와 별개입니다.


## 모바일 밀도·CSV·숨김 행·온라인 열기 회귀

- `node tools/csv-import-quality.mjs`: 로컬 소스/번들 파일 입력으로 16만 행과 대형 UTF-8/UTF-16 LE·BE/CP949 합성 CSV를 검사합니다. `CSV_FIXTURE_DIR`은 D: 경로만 허용하며 기본 `D:/Codex/Temp/wixel-csv-quality`, `CSV_RESULT_PATH`는 JSON 결과 저장 위치입니다.
- `node tools/csv-row-boundaries.mjs`: LF/CR/CRLF 및 중간 빈 행·인용 줄바꿈을 가진 같은 자료를 소형 UTF-8/대형 UTF-16으로 읽고 전체 행렬 해시를 비교합니다.
- `node tools/axis-window-benchmark.mjs`: Node에서 보이는 축 항목 열거 결과·조회 수와 시간 비교. Axis 생성 비용은 제외합니다.
- `node tools/hidden-grid-performance.mjs`: 로컬 소스/번들에서 숨김 행·열, 표시 메모, 틀 고정·스크롤·Undo/Redo를 검사합니다. 소스에서만 `WIXEL_PERF_BASELINE_REF=a6db8e5`로 이전 `view.js`/`axis.js`를 대체해 비교할 수 있습니다. 수정본은 이 변수를 해제합니다.
- `node tools/axis-overlap.mjs`: 나란한 두 표의 필터가 같은 행을 숨길 때 좌표·클릭 위치·스크롤·Undo/Redo와 원본 비트맵 보존을 검사합니다.
- `node tools/server-open-guard.mjs`: 로컬 **소스 전용**. 지연된 온라인 열기, 응답 순서 역전, 편집 보존, 401 인증 재시도와 오류 안내를 가상 API로 검사합니다.
- 기존 `mobile-work-mode.mjs`는 촘촘하게/여유롭게 전환·설정 기억·원본 문서 보존을 포함합니다. `mobile-layout.mjs` 카탈로그는 두 밀도×여섯 화면×15개 탭입니다.

브라우저 도구의 `WIXEL_URL`과 Playwright 환경 변수는 위와 같습니다. 신규 도구는 격리 컨텍스트·합성 자료만 사용하며 외부 쓰기를 막습니다. 시간은 해당 합성 조건의 관측값이며 전체 문서의 속도 보장이 아닙니다. 근거·남은 과제는 [30. 재평가와 품질 개선](../docs/codex/30_재평가와품질개선.md)에 기록합니다.


## 도형·그림의 바로가기 버튼

`node tools/drawing-links.mjs`는 합성 문서에서 도형·그림·그룹 자식의 링크 이동, 시트명·이름 범위, 같은 시트 맨위로, Ctrl+클릭/Enter/K, 링크 편집·제거·Undo, 끊어진 링크 안내, 보호 상태와 모바일 탭을 검사합니다. `WIXEL_URL`로 소스 또는 배포 번들을 지정합니다. 기존 Playwright 환경 변수를 사용하며 격리 컨텍스트에서 외부 연결·API 쓰기를 차단합니다. XLSX 왕복·내부 주소·구조 변경 검사는 `test/drawing-hyperlinks.test.js`, `test/hyperlink.test.js`에 있습니다. 실제 업무 파일은 이 도구나 저장소에 포함하지 않습니다. 지원 범위는 [31번 문서](../docs/codex/31_도형버튼과하이퍼링크.md)에 기록합니다.


## 셀 테두리 픽셀 비교

`node tools/grid-rendering.mjs`는 가는 선·이중선·중간/굵은 선, 채움+넘침 텍스트, 회색/검정 교차점, 점쇄선, 틀 고정·숨김·스크롤·필터·편집을 검사합니다. `WIXEL_URL`로 소스/배포 번들을 정하고 `GRID_SCREENSHOTS`를 D: 폴더로 지정하면 PNG와 좌표 JSON을 보관합니다. DPR 1/1.25/1.5/2 × 배율 50/75/80/100/125% × 두 테마의 장치 픽셀을 판독합니다. 검사 조건·한계는 [32번 문서](../docs/codex/32_셀테두리정밀표시.md)에 기록합니다.


## 확대 슬라이더 전환 회귀

`node tools/grid-zoom.mjs`는 100→400→100→200→50→115% 슬라이더 입력만으로 표시를 바꾸며 강제 다시 그리기를 사용하지 않습니다. DPR 1/1.25/1.5/2, 틀 고정 네 조합에서 선 두께·머리글·첫 셀 경계·글자 여백을 검사합니다. `WIXEL_URL`로 소스/배포 번들을 지정하고 `GRID_ZOOM_SCREENSHOTS`에 D: 폴더를 지정하면 PNG·좌표 JSON을 보관합니다. 검사 시간은 캡처·파일 저장을 포함하므로 렌더러 성능 수치로 쓰지 않습니다. XLS ROW/SCL/XFExt 합성 가져오기·XLSX 왕복 회귀는 `test/xls-view.test.js`, 근거는 [33번 문서](../docs/codex/33_배율변경과셀여백.md)를 참조합니다.


`node tools/merged-row-layout.mjs`는 사용자 파일 대신 병합 제목·일반 행·수동 높이·여러 행 병합을 가진 합성 문서를 사용합니다. 실제 배율 슬라이더(115/185/400%), F2 편집, 명령·홈 리본·행 머리글 더블클릭의 자동 맞춤, 실행 취소/다시 실행, 일반 행의 확대·축소와 XLSX 내보내기 시 원본 높이 보존을 확인합니다. 저장 선택기는 메모리 모의 구현이며 실제 파일이나 원격 문서를 쓰지 않습니다. `WIXEL_URL`로 소스/최종 번들을 지정하고 `MERGED_ROW_SCREENSHOTS`에 D: 절대 폴더를 지정하면 제목 캡처·좌표를 보관합니다. 글꼴 보정은 셀별로 적용되는지와 글자 영역을 검사하며, 모든 운영체제의 글꼴 모양이 Excel과 동일하다는 검사는 아닙니다. 결과의 실패·페이지 오류·원격 쓰기가 모두 없고 종료 코드 0인지 확인하세요.

## 틀 고정 원점과 큰 배율의 키보드 탐색

`node tools/freeze-navigation.mjs`는 합성 문서만 사용하여 저장된 고정 시작 행·열과 본문 시작 위치, 숨긴 행, 방향키·Shift 선택 확장·클릭 좌표를 검사합니다. 고정 영역이 화면보다 큰 330% 배율에서는 행·열·양축의 임시 스크롤 전환을, 배율 축소·창 크기 변경에서는 저장된 고정 설정 복원을 확인합니다. F2 편집과 실행 취소/다시 실행도 포함합니다. `WIXEL_URL`로 소스/배포 번들을 지정하고 `FREEZE_NAV_SCREENSHOTS`에 D: 절대 폴더를 지정하면 PNG와 좌표 JSON을 보관합니다. 새 격리 브라우저에서 API·외부 네트워크와 원격 쓰기를 차단하며 사용자 문서를 읽지 않습니다. 이 도구는 WIXEL의 접근 가능성과 설정 보존을 검증하며, 실제 Excel의 큰 배율 처리와 완전히 같다는 판정은 하지 않습니다.


## IMPORTRANGE 값·자료형·표시 보존

`node tools/importrange-fidelity.mjs`는 격리 브라우저에서 `/api/health`와 `/api/fetch` 응답을 합성 CSV/시트 목록으로 모의합니다. 실제 Google 문서나 사용자 파일은 읽지 않으며 나머지 API·외부 연결·원격 쓰기는 차단합니다. 혼합 자료형, 빈 행·열, 유한 범위의 빈칸 보충과 열린 범위, 큰 금액·백분율 표시, 좁은 열 축소, 사용자 서식 우선 및 실행 취소, QUERY 집계·머리글·다수 자료형 정책, F9 갱신을 검사합니다. 긴 분산 텍스트의 왼쪽·오른쪽·가운데 맞춤에서 이웃 숫자·false·0은 넘침을 막고 실제 빈 칸만 허용하는지도 검사합니다.

일반 공유 주소나 문서 ID에서 범위에 시트 이름을 생략하면 **URL의 `gid`를 무시하고 첫 번째 탭**을 사용합니다. 시트 이름을 명시한 범위는 공개 시트 목록에서 정확한 `gid`를 찾습니다. 웹 게시 주소(`/d/e/…/pub`)의 `gid` 선택은 별도 확장 동작으로 유지하여 각각 검사합니다. `WIXEL_URL`로 소스/배포 번들을 지정하고 `IMPORTRANGE_SCREENSHOTS`에 D: 절대 폴더를 지정하면 PNG와 상세 JSON을 보관합니다. 모의 응답 검사는 외부 Google 서비스의 가용성이나 권한 설정을 보장하지 않습니다.

## 공통 팝업 레이아웃과 키보드 동작

`node tools/dialog-layout.mjs`는 1366×900, 1024×700, 390×844, 320×640 화면에서 이동 옵션, 긴 체크 설명이 있는 시트 보호, 셀 서식 6탭, WIXEL 옵션 9범주, 피벗 옵션, 차트 종류·서식, 도형·그림 서식, 찾기·바꾸기, 선택하여 붙여넣기를 검사합니다. 제목·닫기·확인/취소 버튼이 화면에 있고 입력 중심이 다른 요소에 가려지지 않는지, 본문 가로 넘침·설명 잘림·체크박스 정렬을 확인합니다. Tab 순환·Escape·Alt 접근키와 확인/취소/Undo의 실제 모델 결과도 포함합니다.

`WIXEL_URL`은 소스 또는 배포 번들 주소이며, `WIXEL_DIALOG_OUT`에 D: 절대 경로를 지정하면 대표 PNG와 상세 좌표 JSON을 보관합니다(기본 `D:/Codex/Temp/wixel-dialogs/current`). 좁은 범위 재검사는 `WIXEL_DIALOG_WIDTHS=320,390`, `WIXEL_DIALOG_FILTER=도형 서식`으로 지정할 수 있습니다. 새 격리 컨텍스트의 합성 문서·모의 클립보드만 사용하고 API·외부 연결·원격 쓰기를 차단합니다. Chromium 화면 크기 시뮬레이션이며 실제 모바일 OS·소프트키보드·화면 읽기 프로그램 검증은 아닙니다. 전체 리본 명령 검사는 별도 `smoke.mjs`로 실행합니다.

`node tools/dialog-toast.mjs`는 로컬 소스 서버(`/src/ui.js`)에서 직전 알림이 새 창 버튼을 가리지 않는지, 창 내부의 새 안내·입력 검증 오류·자동 숨김·중첩/모델리스 창 및 시작 소개 알림 동작을 8개 시나리오로 검사합니다. 실제 사용자 문서와 원격 API는 사용하지 않습니다.


## 도형 서식의 옵션 계층과 즉시 적용

`node tools/shape-format-parity.mjs`는 합성 도형에서 도형 옵션 3범주와 텍스트 옵션 2범주, 방향키 탐색, 채우기·선 라디오, 그라데이션 중지점, 비율 잠금, 선 미리 보기, 실제 적용·실행 취소를 검사합니다. 재구성 뒤 접힘 상태·입력 초점·스크롤 유지, 도형 A에서 B로 선택 전환 시 B에만 적용, 320px 화면 접근도 포함합니다. `WIXEL_URL`로 소스/배포 번들을 선택하고 `WIXEL_SHAPE_FORMAT_OUT`에 D: 출력 폴더를 지정합니다(기본 `D:/Codex/Temp/wixel-dialog-keyboard/shape-format`). 새 격리 브라우저에서 사용자 문서를 읽지 않고 API·외부 요청·원격 쓰기를 차단합니다. 실제 Excel 화면 전체와의 픽셀 일치 검증은 아닙니다.


## 팝업 기본 버튼·상시 단축키·차트 대화상자

- `node tools/dialog-keyboard-parity.mjs`: 로컬 소스 서버(`/src/ui.js`) 전용. 모달 바깥 초점/Enter, 동적 본문, 단계별 기본 동작, 비활성 기본 버튼, 중첩·모델리스 창, 상시 단축키·작은 화면 닫기·키 안정성을 검사합니다.
- `node tools/text-to-columns-ux.mjs`: 로컬 소스/번들 전용. Alt A E→Enter 3회, 기본 탭/일반 서식, 고급 숫자 설정, 옵션 접근키·초점, 실제 변환/덮어쓰기/Undo, 320/390px를 검사합니다. 한글 검사는 합성 물리키·조합 이벤트이며 실제 OS IME 구동이 아닙니다.
- `node tools/chart-dialog-ux.mjs`: 소스/배포 번들. 데이터 범위 변경 취소, 입력 직후 Enter, 계열 목록 키보드, 현재/다른/새 시트 이동과 Undo, 차트 유형별 적용 옵션, 모바일 창을 검사합니다. `CHART_DIALOG_OUT` 기본 경로는 `D:/Codex/Temp/wixel-dialog-keyboard/charts`입니다.

`WIXEL_URL`과 Playwright 환경 변수는 위와 같습니다. 모두 격리 브라우저와 합성 문서를 사용하며 API·외부 요청·원격 쓰기를 차단합니다. 지원 범위와 한계는 [38번 문서](../docs/codex/38_대화상자키보드와서식동작.md)를 참고하세요.

## 도형 그리기와 테두리 펜의 모드 격리

`node tools/drawing-mode-isolation.mjs`는 실제 리본 메뉴·마우스 드래그로 테두리 그리기/눈금/지우기와 도형·자유곡선 모드를 전환합니다. 도형 완료·취소 뒤 범위 선택의 셀 값·서식·Undo 불변, 명시적 펜의 1회 적용·자동 종료·다음 클릭/범위 선택 불변·재활성화·Undo, 선 스타일의 설정 전용 동작과 마지막 테두리 재적용, 시트/개체 전환, 서식 복사와 보호 시트, 드래그 도중 펜 시작·Escape 취소를 검사합니다. `WIXEL_URL`은 로컬 소스 또는 번들 주소이며 격리 컨텍스트의 합성 문서만 사용하고 API·외부 연결·원격 쓰기를 차단합니다.

수정 전 재현은 **소스 서버에서만** `DRAWING_BASELINE_REF=684ba67`로 이전 `app.js`를 격리 브라우저에 제공할 수 있습니다. `DRAWING_TEST_FILTER=outline → 도형 완성`으로 해당 경로 하나를 선택하면 이전 버전은 셀 서식/Undo 불변 검사에서 실패해야 합니다. 수정본 검증에서는 `DRAWING_BASELINE_REF`를 해제합니다. 이전 코드는 디스크의 제품 파일을 바꾸지 않으며 실제 사용자 문서를 사용하지 않습니다.

반복 적용 문제의 수정 전 재현은 소스 서버에서 `DRAWING_BASELINE_REF=e26c98a`와 `DRAWING_TEST_FILTER=명시 outline`을 사용합니다. 한 번 적용 후 다음 셀 클릭이 서식과 Undo를 바꾸는 검사에서 실패해야 합니다. `DRAWING_TEST_FILTER=선 스타일`로 선 종류만 선택해도 펜이 자동 시작되던 경로도 검사할 수 있습니다.


## 미디어·그리기·SVG 조합·문서 출력

새 회귀는 `node tools/creative-workspace.mjs`, `node tools/shape-merge-ui.mjs`, `node tools/online-media.mjs`, `node tools/svg-geometry-render.mjs`, `node tools/svg-icon-audit.mjs`입니다. 기존 `online-pictures.mjs`, `smartart.mjs`, `drawing-mode-isolation.mjs`도 함께 확인합니다. `WIXEL_URL`은 로컬 소스 또는 컴파일 번들, Playwright 환경 변수는 위와 같습니다. 모든 브라우저 검사는 격리 합성 문서를 사용합니다. 스크린샷·임시 파일 기본 경로는 D:입니다. 공급자 모킹과 실제 공개 API 읽기, 신규 인증 공급자의 미설정 상태를 구분합니다.

지원 범위·API 키 설정·안전한 SVG 거부·영상 XLSX 호환과 HTML/PDF 한계는 [40번 문서](../docs/codex/40_미디어그리기와문서출력.md)를 참고하세요.


## 피벗 필터 메뉴와 검색 선택

`node tools/pivot-filter-menu.mjs`는 합성 피벗의 실제 행/열/보고서 드롭다운 버튼을 클릭하여 검색 대체/현재 선택 추가, 전체 선택, 빈 결과·0개 체크·IME Enter, 조건 하위 메뉴, Undo, 다중 선택 모드, 보호·늦은 대상 변경, 작은 화면을 검사한다. `WIXEL_URL`은 로컬 소스 또는 번들 주소, `WIXEL_PIVOT_FILTER_OUT`은 D: 출력 폴더(기본 `D:/Codex/Temp/wixel-pivot-filter/source`)다. `WIXEL_PIVOT_FILTER_FILTER`로 이름을 포함하는 시나리오만 실행할 수 있다. 일반 필터 회귀에는 `tools/interaction-ux.mjs`와 `tools/filter-performance.mjs`를 사용한다. 외부 요청과 문서 API 쓰기를 차단하며 사용자 원본 파일을 변경하지 않는다.
