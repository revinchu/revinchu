# 검증 도구 (tools/)

`npm test` 외에 브라우저 동작과 엑셀 파일 회귀를 확인하는 스크립트입니다.
실제 파일을 비교할 때는 **수정 전에 기준 결과를 저장하고, 수정 후 같은 파일과 명령으로 다시 측정**하세요.
차트·피벗 옵션·화면 표시의 검증 범위는 [`14_차트와표시검증.md`](../docs/codex/14_차트와표시검증.md), 최신 단축키 결과는 [`13_단축키점검.md`](../docs/codex/13_단축키점검.md), 공개 배포는 [`12_공개배포와운영검증.md`](../docs/codex/12_공개배포와운영검증.md), 셀 스타일·지우기의 실행 결과와 미검증 범위는 [`11_셀스타일과지우기.md`](../docs/codex/11_셀스타일과지우기.md), 기본 3.0 검증은 [`09_WIXEL3_구현과검증.md`](../docs/codex/09_WIXEL3_구현과검증.md), 원본 인계 검증은 [`06_검증결과.md`](../docs/codex/06_검증결과.md)에 기록합니다.

## 셀 텍스트 넘침 검사

`node tools/cell-text-overflow.mjs`는 합성 긴 글자를 빨강, 이웃 글자를 파랑으로 그린 뒤 PNG의 실제 픽셀을 검사합니다. 숨긴 열·값/수식/0/FALSE/배열 자식·병합·가운데 좌우·빈 채우기·틀 고정·스크롤을 DPR 1/1.5와 75/100/150%에서 비교합니다. 업무 파일을 읽지 않고 원격 쓰기를 차단합니다.

`WIXEL_URL`에 로컬 소스/빌드 서버를 지정하고 `TEXT_OVERFLOW_BROWSER`는 chromium(기본) 또는 webkit으로 지정합니다. `CELL_TEXT_OVERFLOW_OUTPUT`에 D:의 출력 폴더를 주면 합성 화면과 JSON을 보관합니다. `TEXT_OVERFLOW_BASELINE_VIEW`는 수정 전 view.js를 메모리 응답으로 교체해 소스 서버에서 회귀 검출을 확인합니다. 공개 합성 검사는 정확한 주소를 `WIXEL_ALLOWED_TEST_URL`에도 명시해야 합니다. 합격 기준은 bad 0·페이지 오류 0·종료 코드 0이며 자세한 범위는 [숨긴 열과 셀 텍스트 넘침](../docs/codex/106_숨긴열과셀텍스트넘침.md)을 참고하세요.

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
- `node tools/storage-connection.mjs`: 로컬 소스 서버의 격리 Chromium·WebKit 프로필에서 실제 IndexedDB 연결 종료/재열기, 진행 중 쓰기 중단의 중복 실행 방지, CAS 충돌, gzip 셀 청크와 문서 복구 식별자 왕복, Web Lock 대기 중 편집 버전을 검사합니다. `WIXEL_URL`, `WIXEL_CONNECTION_OUT`, `WIXEL_ENGINES`(기본 `chromium,webkit`)를 사용합니다. 실제 iPad의 OS 종료나 업무 파일 부하 검사를 대신하지 않습니다.

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


## 슬라이서 리본·cm 치수·그룹 배치

- `node tools/slicer-ribbon-parity.mjs`: 합성 피벗/슬라이서에서 캡션·설정·보고서 연결, 페이지형 스타일/키보드, cm 단추/전체 크기, 다중 서식·맞춤·앞뒤 순서·선택 창, 그룹화/해제/복사/삭제/숨김 구성원, 보호·늦은 선택 변경·Escape/IME·Undo와 320/390px 화면을 검사합니다. `WIXEL_SLICER_RIBBON_OUT`은 D: 출력 폴더, `WIXEL_SLICER_RIBBON_FILTER`는 선택할 시나리오 이름 일부입니다.
- `node tools/slicer-arrange.mjs`: 슬라이서/도형 혼합 이동, 공통 변위·원점 제한·격자 붙임, noMove/보호 원자 차단, Escape/시트 전환 취소, 그룹 크기·맞춤·분배·쌓임 순서와 Undo/Redo를 검사합니다. 수정 전 재현은 소스에서 `SLICER_ARRANGE_BASELINE=eb6162d`, `SLICER_ARRANGE_FILTER=혼합`로 실행합니다.

`WIXEL_URL`로 소스 또는 최종 번들을 지정하고 기존 Playwright 환경 변수를 사용합니다. 새 격리 브라우저와 합성 자료만 사용하며 외부/API 요청과 쓰기를 차단합니다. 실제 Excel 화면 전체의 픽셀 일치나 모든 조합의 동등성을 판정하는 도구는 아닙니다. 표준 XLSX·실제 Excel 왕복 근거는 [47번 문서](../docs/codex/47_슬라이서리본과배치.md)를 참고하세요.


## 데이터 유효성 검사

`node tools/validation-parity.mjs`는 실제 데이터 리본 메뉴와 네 탭을 통해 8종 제한·8종 비교, 설명 메시지, 중지/경고/정보, 같은 설정 적용, 범위 선택과 취소, IME 메타데이터, 잘못된 값 표시, 이름 목록과 숫자 정밀도 보존, Alt/Enter/Escape, 보호·늦은 대상 변경·Undo와 320/390px 화면을 검사합니다. `WIXEL_URL`은 소스 또는 번들 URL, `WIXEL_VALIDATION_OUT`은 D: 출력 폴더(기본 `D:/Codex/Temp/wixel-validation/source`), `WIXEL_VALIDATION_FILTER`는 시나리오 이름 일부입니다. 격리 브라우저의 합성 문서만 사용하며 API·외부 요청·원격 쓰기를 차단합니다. IME composition 이벤트 검사는 실제 운영체제의 입력기 전환 검사가 아닙니다. 표준 파일과 네이티브 Excel 검증·제한은 [48번 문서](../docs/codex/48_데이터유효성검사.md)에 기록합니다.


## 개체·셀 범위의 그림/SVG 저장

`node tools/image-export.mjs`는 실제 우클릭·리본·복사 메뉴에서 사진·도형·아이콘·다중 개체·셀 범위를 PNG/JPEG/SVG로 저장하고 파일 내용을 다시 읽는다. 투명/흰 배경·해상도·회전·색·숫자 서식·범위 크기, 초과 크기 거절, 취소·권한·쓰기 실패·다운로드 확인, 문서/선택 변경 차단·원본/Undo 보존과 320px 화면을 검사한다. `WIXEL_URL`로 소스/번들을 지정하며 출력 변수와 제한은 [49번 문서](../docs/codex/49_그림과SVG저장.md)를 참고한다. 격리 합성 문서와 메모리 파일 선택기를 사용하며 실제 사용자 파일·원격 쓰기를 사용하지 않는다.


## 갤러리 글자와 접근키 배치

`node tools/gallery-access-captions.mjs`는 소스 서버의 공통 팝업에서 SVG 텍스트/실제 이름 분리와 동적 접근키를 4개 화면 크기로 검사한다. `GALLERY_CAPTION_BASELINE=ad2dadc`는 이전 공통 UI를 메모리에서만 제공하는 수정 전 재현 옵션이다. `node tools/chart-gallery-layout.mjs`는 소스/배포 번들의 추천·모든 차트, 긴 한글·검색·키보드·시트 배율·작은 화면의 실제 글자 경계를 검사한다. `WIXEL_CHART_GALLERY_OUT`과 `WIXEL_CHART_GALLERY_FILTER`로 D: 출력 경로와 개별 검사를 지정한다. 확대 모의·검증 범위는 [50번 문서](../docs/codex/50_차트갤러리와팝업글자배치.md)에 기록한다. `dialog-layout.mjs`는 `WIXEL_DIALOG_EXCLUDE`에 이름 일부를 지정해 다른 전용 도구에서 검사하는 팝업을 제외할 수 있다.

`node tools/popup-gallery-layout.mjs`는 표·셀·SmartArt·아이콘·그림 스타일 갤러리 5종을 320/390/1024px에서 검사한다. 셀 스타일 이름과 접근키의 잘림, 미리보기 내부 키 삽입, 가로 넘침을 확인하며 `WIXEL_POPUP_OUT`에 D: 출력 폴더를 지정한다. 소스와 최종 번들에 같은 도구를 사용한다.


## 직접 만든 피벗의 보고서 필터

`node tools/pivot-report-layout.mjs`는 실제 피벗 생성/필드 추가 UI로 보고서 필터만 있는 상태와 본문 포함 상태, 행·열 우선과 줄바꿈 배치, 본문 앵커, 선택/우클릭/이름/표시 옵션, 주변 셀 보존, Undo/Redo 및 XLSX 저장 후 다시 열기를 검사한다. 소스/번들은 `WIXEL_URL`, D: 결과 위치는 `WIXEL_PIVOT_REPORT_OUT`을 사용한다. `WIXEL_PIVOT_REPORT_NATIVE_FIXTURE`로 개인 자료가 없는 Excel 합성 파일을 지정해 가져오기와 비교할 수 있다. 네이티브 Excel 근거와 실제 파일 검증의 범위는 [51번 문서](../docs/codex/51_피벗보고서필터.md)를 참고한다.


## 모바일 한 줄 도구와 우클릭

`node tools/mobile-compact-chrome.mjs`는 상·하단 한 줄 기하, 작은 글꼴/크기 입력, 실제 마우스 휠 가로 탐색, 우클릭 도구의 겹침/키힌트/브라우저 기본 메뉴 차단 및 문서 불변을 검사한다. `WIXEL_URL`로 소스나 Cloudflare 번들을 선택하고 `WIXEL_MOBILE_COMPACT_OUT`으로 D: 캡처 폴더, `WIXEL_MOBILE_COMPACT_FILTER`로 시나리오 이름 일부를 지정한다. `mobile-layout.mjs`는 전체 상황별 리본의 두 밀도, `mobile-popups.mjs`는 팝업 최초 배치·현재 차트 요소 전환·파일 저장 확인과 실제 PDF를 함께 검사한다. 합성 문서만 사용하며 브라우저 검증 범위는 [52번 문서](../docs/codex/52_모바일한줄도구와우클릭.md)를 참고한다.


## 로컬 기능 감사 배치

`node tools/functional-audit.mjs --list`는 기존 도구를 기능·실행 전제별로 분류한 manifest를 만듭니다. 기본 실행은 격리 합성 브라우저/순수 Node 검사만 동시 2개까지 실행하며, `--only settings-ux,keys`, `--concurrency 1`, `--timeout 600`, `--out D:/Codex/Temp/wixel-functional-audit`로 범위를 지정할 수 있습니다. `WIXEL_URL`은 현재 로컬 소스 서버로 지정합니다. 소스 모듈을 직접 읽는 도구가 있으므로 전체 배치를 공개 주소나 번들에 그대로 적용하지 않습니다.

각 도구의 종료 코드·실행 시간·요약·오류·실행 전후 주요 제품 파일 해시·원문 로그를 저장합니다. `--resume`은 기록된 주요 파일 해시가 같은 기존 통과 항목만 재사용합니다. 새 실행의 해시 범위는 `src/**/*.js` 전체(추가·삭제 포함), `styles.css`, `index.html`이며 `sourceScope`에 정확한 목록을 저장합니다. 이전 감사 기록의 5개 파일 범위는 그대로 남기고 새 전체 범위 결과로 간주하거나 재개에 재사용하지 않습니다. 이 해시는 서버가 실제 제공하는 번들의 일치까지 증명하지 않습니다. 동일 릴리즈의 최종 통합 확인은 별도 빌드 검사로 구분합니다. 명시적 파일 입력이 필요한 비교/Excel 도구, 실제 보관함 쓰기 도구, 다른 담당자의 검사, 단독 실행해야 하는 성능 도구와 데이터 갱신 스크립트는 자동 배치에서 제외하고 사유를 manifest에 남깁니다.

`audit-playwright.mjs`와 `audit-network-preload.mjs`는 이 배치 전용 가드입니다. 새 headless 브라우저만 허용하며 외부/API/쓰기 요청의 실제 전송을 차단합니다. 개별 도구가 제공하는 합성 API 응답은 허용하므로 서버 오류·충돌·저장 동작을 원격 문서 없이 시험합니다. 파일 선택기는 각 도구의 메모리 모의 구현을 사용하며 실제 사용자 문서나 열린 브라우저에 연결하지 않습니다. 가드가 막은 요청과 제품 오류는 구분해야 합니다.

`node tools/functional-audit.mjs --only command-connection-audit`는 등록된 COMMANDS/MENUS와 모든 리본 조작·키팁의 연결 목록을 JSON으로 저장합니다. 함수 메뉴의 `fn:` 동적 라우터도 별도로 기록합니다. 이 연결 검사나 명령 열기 smoke, 기존 기능 도구의 통과 수는 모든 기능의 정확성 또는 Excel 완전 동등성을 뜻하지 않습니다. 실패 원인은 제품 결함·낡은 검사 계약·환경/fixture 부족으로 조사한 후 기록합니다.


## 모바일 팝업 밀도와 메뉴 구성

`node tools/mobile-popup-density.mjs`는 실제 모바일 버튼 전환과 데스크톱 복원, 320px 세로/844px 가로 화면의 우클릭 항목 숨김·미니 도구·전체 글꼴/크기 목록·키보드 하위 메뉴·셀 서식·빠른 분석·무늬·그리기를 검사한다. 모바일 격자 Ctrl+1/Shift+F10/한 칸 이동과 Undo도 확인한다. `WIXEL_URL`로 로컬 소스/번들 서버를, `WIXEL_MOBILE_DENSITY_OUT`으로 D: 결과 경로를, `WIXEL_MOBILE_DENSITY_FILTER`로 이름 부분을 선택한다. 외부/API 쓰기를 차단한 합성 문서 검사이며 실제 iPhone Safari 인증과 구분한다. 치수와 검증 범위는 [55번 문서](../docs/codex/55_모바일팝업밀도.md)를 참고한다.


## iPhone 마우스 손바닥 이동

`node tools/mobile-hand-pan.mjs`는 포인터만 전달되는 버튼 조합과 호환 Mouse Events 중복, 먼저 뜨는 우클릭, 첫 buttons3, 버튼 지연·해제·취소·시트 전환, 펜 구분, 상단 손바닥 버튼의 한 버튼/터치 이동 및 문서 불변을 검사한다. `WIXEL_BROWSER=chromium|webkit`, 로컬 `WIXEL_URL`, D: 결과용 `WIXEL_HAND_PAN_OUT`, 부분 이름 필터 `WIXEL_HAND_PAN_FILTER`를 사용한다. WebKit에서는 Chromium CDP 터치 주입 2개를 이유와 함께 제외하고 나머지 포인터/마우스/키보드 경로를 검사한다. 실제 iPhone/Bluetooth 인증과 구분하며 [56번 문서](../docs/codex/56_아이폰손바닥이동.md)에 범위를 기록한다.


## 피벗 필드 끌기와 권한

`node tools/pivot-field-drag.mjs`는 네 영역 추가·영역 간 이동·재정렬·중복 값·Σ 위치·Undo·취소·자동 스크롤·작은 화면을 실제 마우스 및 모의 Pointer/CDP 터치로 확인한다. `WIXEL_BROWSER=chromium|webkit`, `WIXEL_URL`, D: 출력 경로 `WIXEL_PIVOT_DRAG_OUT`, 부분 이름 `WIXEL_PIVOT_DRAG_FILTER`를 사용한다. `tools/pivot-pane-permissions.mjs`는 피벗 허용 보호·최종본·읽기 전용·계산 필드 공유 및 오래된 편집기 적용을 검사한다. 개인 문서나 원격 쓰기 없이 합성 문서를 사용하며 실제 기기 검증과 구분한다. 범위와 남은 차이는 [57번 문서](../docs/codex/57_피벗필드이동과검수.md)를 참고한다.


필드 목록 제거 회귀는 `WIXEL_PIVOT_DRAG_FILTER="목록으로 제거"`로 선택한다. 마우스·모의 Pointer/Touch·HTML DragEvent, 중복 집계와 체크 상태, Undo/Redo의 결과·선택 보존을 검사하며 `WIXEL_PIVOT_DRAG_BASELINE=3056fba`는 기존 app/helper/styles를 메모리에서 제공한다. 공개 합성 검사에는 정확한 URL을 `WIXEL_ALLOWED_TEST_URL`에도 지정한다. D: 출력 경로를 요구하고 검사 결과·범위는 [97번 문서](../docs/codex/97_피벗필드목록으로제거.md)에 기록한다.

## 개체 스타일과 모바일 시트 목록

- `node tools/mobile-slicer-gallery.mjs`: 모바일 320/390px 두 밀도와 데스크톱에서 갤러리 이전·다음·더 보기, 실제 스타일 변경·Undo, 키보드·작은 메뉴 경계를 검사합니다. 출력은 `WIXEL_SLICER_GALLERY_OUT`으로 지정합니다.
- `node tools/object-style-design.mjs`: 표·피벗·슬라이서 스타일 만들기·수정·복제·삭제·기본값·공유 사용처·요소 상속·명시 해제·숫자 형식·보호·오래된 편집기·실제 채우기와 XLSX 재열기를 검사합니다. `WIXEL_OBJECT_STYLE_OUT`과 부분 이름 `WIXEL_OBJECT_STYLE_FILTER`를 사용합니다.
- `node tools/mobile-sheet-actions.mjs`: 시트 숨기기 취소와 이동/복사의 앱 목록, 마우스·터치·키보드, 취소·Undo/Redo·보호·오래된 문서, 모바일 경계를 검사합니다. standalone/display-mode 모사와 반복 취소의 pageclose/crash/navigation/newpage·포커스·렌더·직렬화 계측, 0/21/34px 하단 안전영역 및 47px 상단/양옆 노치 검사를 포함합니다. 실제 iPhone OS 제스처 검사는 아닙니다. 출력 변수는 `WIXEL_SHEET_ACTIONS_OUT`입니다.
- `node tools/pivot-style-capture.mjs`: 가져온 피벗의 스타일 변경·지우기 실제 표시와 직접 지정 서식 보존을 확인합니다. 출력 변수는 `WIXEL_PIVOT_STYLE_OUT`입니다.

`WIXEL_URL`에 로컬 소스 또는 최종 번들을 지정합니다. `WIXEL_BROWSER`는 도구별로 Chromium/WebKit를 선택하며 각 파일의 기본값을 참고하세요. 격리 컨텍스트의 합성 문서만 사용하고 외부/API 쓰기를 차단합니다. 실제 Excel 정상 열기/재저장과 WebKit 화면 검사의 구분, iPhone OS 홈 이탈의 미확인 상태는 [58번 문서](../docs/codex/58_개체스타일과모바일목록.md)에 기록합니다.

## 초록색 상단 탭 가로 탐색

`node tools/mobile-header-scroll.mjs`는 Chromium/WebKit의 320·390px 모바일, 데스크톱에서 수동 모바일 전환, 모바일을 끈 좁은 창에서 리본 탭의 양방향 끌기·휠·클릭·파일 메뉴·취소/재진입을 검사합니다. Pointer Events와 호환 mousemove의 중복 이동, MouseEvents만 전달되는 경로, 터치 이벤트 전달과 문서·Undo 불변도 확인합니다. 모바일 WebKit에서는 Playwright가 실제 휠 입력을 지원하지 않아 WheelEvent 전달/소비로 구분하며, 데스크톱 WebKit과 Chromium에서는 실제 마우스 휠을 사용합니다. 실제 iPhone/Bluetooth 마우스 검증은 포함하지 않습니다.

`WIXEL_URL`은 로컬 소스/번들, `WIXEL_HEADER_SCROLL_OUT`은 D: 출력 폴더, `WIXEL_HEADER_SCROLL_FILTER`는 이름 일부, `WIXEL_BROWSERS=chromium,webkit`은 실행 엔진을 지정합니다. `WIXEL_HEADER_BASELINE_REF`를 지정하면 소스 서버에서 해당 Git 커밋의 `mouse-work.js`, `mobile-work.js`, `styles.css`만 격리 컨텍스트에 응답하여 수정 전 동작을 비교합니다. 제품 파일·사용자 문서는 바꾸지 않으며 외부·API·쓰기 요청을 차단합니다.

## 표 요약행 숨김과 복원

`node tools/table-totals-preservation.mjs`는 합성 표에서 요약행의 평균·사용자 수식·라벨·직접 서식·빈 칸을 숨겼다 복원하고, 필터·정렬·Undo/Redo·보호·원본 값 변경·복원 위치의 사용자 내용 보존 및 숨긴 XLSX 저장 후 다시 열기를 검사합니다. `WIXEL_URL`로 소스/번들 앱을, `WIXEL_BROWSER=chromium|webkit`로 엔진을, `WIXEL_TOTALS_OUT`으로 D: 출력 폴더를 지정합니다. `WIXEL_TOTALS_FILTER`는 시나리오 이름 일부를 선택합니다. 새 격리 컨텍스트의 합성 자료만 사용하며 외부/API/쓰기를 차단합니다. XLSX 생성은 로컬 소스 writer를 사용하고, 다시 열기는 지정한 앱에서 수행합니다. 앱 내부 셀 서식 보존과 XLSX 표준 함수·수식·라벨 보존의 범위 차이, 표시 옵션 24개 관찰 및 네이티브 Excel 근거는 [60번 문서](../docs/codex/60_필터단추와표시상태보존.md)에 기록합니다.


## 필터 단추와 Excel 조건 보존

`node tools/table-filter-display.mjs`는 표 디자인의 단추 표시와 데이터 필터 해제를 구분하고, 실제 단추·조건·숨긴 행·합계·슬라이서·부분 단추·Undo/Redo·머리글·기본 줄무늬·보호·모바일·XLSX 재열기를 확인합니다. `WIXEL_URL`은 소스/최종 번들, `WIXEL_BROWSER=chromium,webkit`는 엔진, `WIXEL_FILTER_DISPLAY_OUT`은 D: 출력 경로, `WIXEL_FILTER_DISPLAY_FILTER`는 시나리오 이름 일부입니다. `WIXEL_FILTER_NATIVE_FILE`을 지정하면 Excel에서 만든 합성 파일의 단추 숨김도 검증합니다. 미지원 원형 조건의 재적용 차단과 ~ 리터럴의 저장 후 재적용을 포함하며, 외부/API 쓰기를 차단합니다. [60번 문서](../docs/codex/60_필터단추와표시상태보존.md)에 실제 Excel·표준 검증과 한계를 구분해 기록합니다.


## 홈 화면 앱 버전 안내와 시트 상태 저장

- `node tools/mobile-release-notice.mjs`: 모바일 도구의 버전 진입·새 버전 감지·열린 도구 문구 갱신·파일 저장 취소에서 문서/선택/Undo 불변과 추가 이동 없음을 검사합니다. `WIXEL_URL`, `WIXEL_BROWSER`(chromium/webkit), `WIXEL_RELEASE_NOTICE_OUT`을 사용합니다. standalone 표시는 모의 값입니다.
- `node tools/sheet-state-storage.mjs`: 로컬 소스 서버의 격리 실제 IndexedDB에 합성 문서만 저장합니다. 30만 셀 숨김/Undo/Redo의 0청크 재사용, 셀·다른 메타 변경, 16MB 분할 청크 GC, 중단·CAS·Web Locks 부재를 검사합니다. `WIXEL_URL`, `WIXEL_SHEET_STORAGE_OUT`을 사용합니다.

검증 결과와 iPhone 실기기 한계는 [61번 문서](../docs/codex/61_아이폰홈화면시트조작.md)에 기록합니다.


## 모바일 외부 키보드 입력

`node tools/mobile-hardware-keyboard.mjs`는 격리 합성 문서에서 실제 Playwright 키 입력·마우스·터치를 사용해 첫 문자, Enter/Tab/Escape, 수식, 찾기, 글꼴, 설정 보존·복귀와 Undo를 검사합니다. IME 이벤트는 합성이며 실제 iPhone 시스템 키보드·Bluetooth 검사가 아닙니다. `WIXEL_URL`, `WIXEL_BROWSER`(chromium/webkit), `WIXEL_KEYBOARD_OUT`, `WIXEL_KEYBOARD_FILTER`를 지원하며 외부/API 요청과 원격 쓰기를 차단합니다. 세부 범위는 [62번 문서](../docs/codex/62_모바일외부키보드입력.md)를 참고합니다.

## 개선 계획의 첫 안전성 배치

- `node tools/snapshot-memory.mjs`: 별도 Node 프로세스의 30만 셀 old/Blob 직렬화 3회·동일 byte 비교. `WIXEL_SNAPSHOT_BENCHMARK`는 결과 JSON 경로다. JS heap 관측은 iPhone/전체 peak 메모리가 아니다.
- `node tools/library-atomic.mjs`: 기존 원자 저장 + Blob/JSON 블록 복원 70검사. 소스 서버 전용, `WIXEL_BROWSER=chromium|webkit`. Windows WebKit은 IDB Blob 검사를 위해 D:의 새 persistent 프로필을 사용한다.
- `node tools/command-recovery.mjs`: F4 보호 및 온라인 충돌 사본 실패/문서 변경 경로. `WIXEL_URL`, `WIXEL_BROWSER`, `WIXEL_RECOVERY_OUT`을 사용한다. 모든 온라인 API는 메모리 모의다.
- `node tools/mobile-clipboard.mjs`: 모바일 내부 셀 Ctrl/Meta+C/X/V, native 이벤트 부재/중복, 외부 사본·보호·IME. `WIXEL_URL`, `WIXEL_BROWSER`, `WIXEL_CLIPBOARD_OUT`을 사용한다. OS 클립보드를 모의하며 실물 iPhone 홈 이탈 검사가 아니다.
- `node tools/xlsx-safety-fixtures.mjs`: 표준 로그축 5종과 TABLE 값-only 합성 자료 생성. 기본 결과는 `D:/Codex/Temp/wixel-xlsx-safety`다.
- `node tools/xlsx-preflight.mjs`: XLSX 저장 전 한도·미지원 안내 취소/승인·문서 변경과 시트 추가 경고 이력. 시스템 저장 창은 메모리 핸들로 모의한다.

연구·점수·설계는 [63번 문서](../docs/codex/63_개선기획과설계.md), 실행 범위는 [64번 문서](../docs/codex/64_개선실행과검증.md)를 참고한다.


## iPad Bluetooth 보조키·단축키 확인

- `node tools/keyboard-check.mjs`: 모바일 작업 도구 및 바로 가기 키에서 기기 내 확인 창을 연다. Ctrl/Alt/Meta 수신, 한글/legacy 대체 판정, IME 통과, Tab/Escape, 클립보드 접근 0, 문서 불변과 320px 배치를 검사한다. `WIXEL_URL`은 로컬 서버, `WIXEL_BROWSER=chromium|webkit`, `WIXEL_KEYBOARD_CHECK_OUT`은 D: 결과 경로다.
- `tools/mobile-clipboard.mjs`에 iPad 데스크톱 UA·모바일 OFF, code 없는 한글+keyCode, Unidentified, 229, Alt 유지, 실제 리본/셀 마우스 복귀 검사를 추가했다. API의 선택 변경을 실제 DOM 포커스 복원으로 간주하지 않는다.
- [65번 문서](../docs/codex/65_iPad키보드단축키.md)에 기준 릴리스 재현과 최종 번들 결과·실기기 한계를 기록한다.


## 모바일 기기 최적화 OFF의 포인터 이동

`node tools/apple-pointer-work.mjs`는 iPhone, Mac 플랫폼으로 표시되는 iPad 및 Android를 모의하고 모바일 최적화 OFF에서 리본/상단 메뉴/빠른 실행/시트 탭의 드래그·휠과 손바닥 도구·양버튼 이동을 검사한다. 문서·선택·Undo 보존, 일반 클릭 복원, ON→OFF 유지, 넓은 iPad, 일반 데스크톱의 기존 제한도 확인한다. `WIXEL_URL`, `WIXEL_BROWSER=chromium|webkit`, `WIXEL_APPLE_POINTER_OUT`(D: 경로), `WIXEL_APPLE_POINTER_FILTER`를 지원한다. OS/Bluetooth 하드웨어 검사가 아니며 [66번 문서](../docs/codex/66_모바일기기마우스이동.md)에 범위와 결과를 기록한다.


## 사진 위치별 가시성

`node tools/picture-visibility.mjs`는 합성 그림의 실제 PNG 픽셀로 위치별 소실·셀 글자 겹침, 효과·자르기·크기·배율, 틀 고정과 스크롤, 그림 서식 미리보기와 온라인 이미지 로드를 검사한다. `WIXEL_URL`, `WIXEL_BROWSER=chromium|webkit`, `WIXEL_PICTURE_VISIBILITY_OUT`(D: 경로)을 지원한다. 실제 사용자 파일·외부 API·서버 저장을 사용하지 않는다. [67번 문서](../docs/codex/67_사진위치별표시오류.md)에 재현과 최종 검증 범위를 기록한다.


## iPad Ctrl 포인터·개체 메뉴·사진 검색

`ipad-ctrl-wheel.mjs`, `slicer-modifier-selection.mjs`, `chart-controls-layer.mjs`, `online-gallery-layout.mjs`는 로컬 `WIXEL_URL`에서 격리 합성 문서로 동작한다. `WIXEL_BROWSER=chromium|webkit`로 브라우저를 선택한다. 각각 배율·modifier 해제, 슬라이서 추가 선택/개체별 메뉴, 차트 단추의 겹침과 실제 클릭, 화면 크기별 사진 미리보기와 선택 보존을 확인한다. 실제 Bluetooth 기기 검사가 아니며 최종 결과와 범위는 [68번 문서](../docs/codex/68_iPad포인터와개체메뉴.md)에 기록한다.

### 차트 항목 축 레이블

`node tools/chart-axis-labels.mjs`는 실제 축 글자 수, SVG 경계와 회전 사각형의 충돌, 3D/다단계/수동 간격·각도를 검사합니다. `node tools/chart-axis-settings.mjs`는 축 직접 클릭과 전체 서식 패널의 설정·실행 취소·XLSX 저장을 검사합니다. `WIXEL_URL`은 로컬 서버, `WIXEL_BROWSER=chromium|webkit`을 사용하며 Playwright 경로는 위 공통 안내를 따릅니다. 원격 쓰기를 차단한 합성 문서만 사용합니다. 실제 Excel 앱의 자동 배치와 픽셀 동등성을 검사하는 도구는 아닙니다.


### 표 빠른 스타일 전체 적용

`node tools/table-quick-style.mjs`는 실제 테이블 디자인 리본에서 스타일을 골라 표 전체 채우기·글자색·강조·선 변경을 확인합니다. 직접 강조색과 기존 스타일, 행/열 서식, 기본 셀 서식, 사용자 지정 빈 셀, 모바일 390px, 한 번의 실행 취소·다시 실행과 XLSX 재열기를 검사합니다. `WIXEL_URL`은 로컬 서버, `WIXEL_BROWSER=chromium|webkit`, `WIXEL_TABLE_STYLE_OUT`은 D: 결과 경로이며 외부 쓰기를 차단한 합성 자료를 사용합니다.


### 행·열 머리글 경계 조절

`node tools/header-resize.mjs`는 마우스 hover·드래그·더블클릭, 배율·틀 고정·스크롤, 다중 선택, Esc/blur 취소, 보호·최종본 및 XLSX 크기 왕복을 검사합니다. `WIXEL_URL`은 로컬 서버이며 `WIXEL_BROWSER=chromium,webkit`, `WIXEL_HEADER_RESIZE_OUT`은 D: 결과 경로, `WIXEL_HEADER_RESIZE_FILTER`는 사례 이름 필터입니다. 실제 iPad 하드웨어 검사가 아닌 격리 합성 문서 검사입니다.


### 문서 전환·늦은 응답 회귀

`node tools/document-switch.mjs`는 합성 XLSX/WIXEL의 순차·지연 열기, 늦은 오류·셀 준비, 새 문서, 같은 파일명, 편집·Undo/Redo·자동 저장·재열기와 이전 진행창 제거를 검사합니다. `WIXEL_DOCUMENT_SWITCH_OUT`은 D: 결과 경로, `WIXEL_DOCUMENT_SWITCH_FILTER`는 사례 이름 필터입니다.

`node tools/document-published-switch.mjs`는 합성 API로 게시본·시작 URL의 늦은 응답, 보기 종료, 현재 게시본 갱신과 사본, 실패한 열기 후 갱신 재개, 피벗 준비 취소·서식 보존을 검사합니다. `WIXEL_DOCUMENT_PUBLISHED_OUT`과 `WIXEL_DOCUMENT_PUBLISHED_FILTER`로 결과 경로와 사례를 지정합니다.

두 도구 모두 `WIXEL_URL`은 로컬 서버만 허용하며 `WIXEL_BROWSER=chromium,webkit`와 `PLAYWRIGHT_MODULE`을 사용합니다. 실제 iPad의 Safari 세션을 제어하는 검사는 아닙니다.


### 대형 XLSB 저장·복구 안정성

`node tools/document-recovery.mjs`는 첫 저장 전 종료, 같은 문서의 미저장 편집, 완료 저장 재시작, 다른 탭의 포인터 변경, 동일 문서 동시 편집 분기, 저장 실패 및 모바일 복구 창을 합성 문서로 검사합니다. `WIXEL_RECOVERY_OUT`과 `WIXEL_RECOVERY_FILTER`를 지원합니다.

`node tools/xlsb-recovery-stress.mjs`는 비공개 JSON 목록(`WIXEL_REAL_AUDIT_MANIFEST`, 항목별 `id`, `path`)의 실제 파일을 로컬 브라우저에서 순서대로 열고 첫 저장 완료, UI 편집, 편집 저장 완료, 재열기 뒤 문서·시트·수식 셀 수와 편집 보존을 확인합니다. `WIXEL_REAL_AUDIT_IDS`, `WIXEL_ROLLBACK_OUT`, `WIXEL_ROLLBACK_TIMEOUT`을 지원합니다. 선택 항목 `WIXEL_CHROMIUM_HEAP_MB`는 Chromium의 old-space 예산(MiB)을 제한하며 결과에 기록합니다. `WIXEL_ISOLATED_FILES=1`은 파일별 새 브라우저 프로필로 개별 복구 검사를 하며, 기본 연속 검사와 구분하여 기록합니다. 명시적 GC는 사용하지 않습니다. 호스트 여유 메모리가 1GiB 아래이면 보호를 위해 중단하며 이를 통과로 계산하지 않습니다. 원본을 수정하지 않으며 결과·브라우저 프로필·화면은 D:의 비공개 검사 폴더에만 둡니다. 실제 업무 파일 및 결과를 커밋하지 마세요.

두 도구의 `WIXEL_URL`은 로컬 서버만 허용하고 `WIXEL_BROWSER=chromium,webkit`를 지원합니다. Windows WebKit의 실제 IndexedDB Blob 검사는 D:의 격리 persistent profile을 사용합니다. 실제 iPad의 OS 메모리 종료 인증과는 다릅니다. 범위와 결과는 [74번 문서](../docs/codex/74_XLSB복구안정성.md)에 기록합니다.

- `tools/pivot-document-release.mjs`: 합성 피벗 필드 창을 열거나 숨긴 뒤 다른 문서를 열어, 이전 창·드래그 리스너·통합문서가 Chromium 강제 GC 후 해제되는지 검사합니다. 동일 문서에서 창을 숨길 때는 상태를 유지합니다. `WIXEL_URL`, `WIXEL_PIVOT_RELEASE_OUT` 사용. 물리 iPad 종료 재현 검사가 아닙니다.

- `tools/reloaded-document-release.mjs`: 합성 문서 저장→재시작→다른 문서 저장 뒤 이전 Workbook의 도달 가능성을 Chromium 강제 GC로 확인합니다. 강제 GC는 참조 해제 진단 전용이며 실제 대형 파일의 자연 메모리 사용량 판정에 대체하지 않습니다. 준비 확인은 Boolean을 반환해 검사 도구가 문서 JSHandle을 보유하지 않습니다.

## 브라우저 호환성 회귀

- `browser-quality.mjs`: Chromium·WebKit·Firefox 합성 문서 입력/이동/저장/다운로드 재열기. WIXEL_BROWSERS로 엔진 목록, WIXEL_BROWSER_QUALITY_OUT으로 결과 경로를 지정한다. 소스 서버에서는 IndexedDB/보관함 직접 모듈 검사도 한다. 정적 번들은 WIXEL_BROWSER_QUALITY_EXCLUDE=IndexedDB로 이 직접 모듈 사례를 제외한다.
- `clipboard-feedback.mjs`: 클립보드 API 거절·부재·지연과 수동 복사 재시도. WIXEL_BROWSER로 엔진, WIXEL_CLIPBOARD_FEEDBACK_OUT으로 결과 위치를 지정한다.
- `popup-viewport-compat.mjs`: 모바일 밀도 OFF/ON의 축소된 visualViewport, 하위 메뉴, 안전 영역, 크기 복원. WIXEL_BROWSER로 엔진, WIXEL_POPUP_VIEWPORT_OUT으로 결과 위치를 지정한다.

실제 장치 OS/Bluetooth 검사와 합성 브라우저 검사를 구분한다. 최신 범위와 한계는 [79번 검증 문서](../docs/codex/79_브라우저호환성과입력저장보강.md)를 따른다.

## 차트 예측선 영역 회귀

`tools/chart-forecast-ui.mjs`는 격리 합성 차트로 실제 선형 예측 메뉴, 실행 취소/다시 실행, 자동·고정·보조·역방향 축, 가로 막대, 지수/이동 평균, 작은 차트, 기존 XY 표시를 검사한다. SVG를 래스터화하여 추세선이 그림 영역 밖에 표시되는지도 측정한다. WIXEL_BROWSER는 chromium/webkit/firefox, WIXEL_URL은 로컬 서버, WIXEL_FORECAST_OUT은 D드라이브 결과 위치다. WIXEL_FORECAST_BASELINE에 이전 chart.js 파일을 지정하면 소스 서버의 해당 응답만 이전 코드로 바꿔 재현한다. WIXEL_FORECAST_PUBLIC=1은 정확한 위셀 공개 호스트의 합성 검사만 허용하며 API/외부 요청과 쓰기는 계속 차단한다. 범위·결과는 [80번 문서](../docs/codex/80_차트예측선영역.md)를 따른다.

## 차트 색과 스타일 회귀

`tools/chart-colors-ui.mjs`는 가져온 그라데이션·개별 점 색이 있는 합성 차트에서 실제 리본/옆 버튼/서식 창의 색·스타일·미리보기·Undo·보호 차단을 검사한다. `WIXEL_BROWSER`는 chromium/webkit/firefox, `WIXEL_COLORS_OUT`은 D 결과 경로, `WIXEL_URL`은 로컬 소스/번들 서버이다. `WIXEL_COLORS_BASELINE`은 기준 app.js/chart.js/chart-edit.js/chart-selection-ui.js가 있는 폴더를 지정한다. `WIXEL_COLORS_PUBLIC=1`은 정확한 위셀 공개 호스트에서 합성 검사만 허용한다. API와 외부 요청·쓰기는 계속 차단한다. 상세 범위는 [81번 문서](../docs/codex/81_차트색과스타일명령.md)를 따른다.

## 색상표·메뉴 갤러리 키보드 회귀

- `palette-keyboard.mjs`: HFC/HH 실제 Alt 경로, Alt 유지·해제, 모바일/데스크톱, 화살표·Home/End·Tab·Enter/Esc, 적용/Undo와 고정 선택 범위, 글꼴·차트 색 검색. `WIXEL_PALETTE_OUT`, `WIXEL_PALETTE_FILTER`, `WIXEL_PALETTE_BASELINE` 지원.
- `menu-gallery-keyboard.mjs`: 무늬·조건부 서식·표·피벗·슬라이서·차트·테마·도형의 실제 키팁 진입, 이동, 적용/Undo와 기존 Alt+↓/←. `WIXEL_GALLERY_OUT`, `WIXEL_GALLERY_FILTER` 지원.

둘 다 `WIXEL_URL`, `WIXEL_BROWSER=chromium|firefox|webkit` 사용. 기본은 로컬 주소이며 `WIXEL_PALETTE_PUBLIC=1`/`WIXEL_GALLERY_PUBLIC=1`일 때 정확한 위셀 공개 호스트에 한해 합성 검사한다. API·외부 요청·원격 쓰기는 차단한다. 모든 결과는 D: 임시 폴더에 둔다.

## 수식 가져오기 브라우저 회귀 검사

`tools/import-formula-diagnostics.mjs`는 합성 XLSX를 실제 파일 입력으로 열어 원본 `#NAME?` 오류와 미지원 함수의 저장값을 구분하고, LET/LAMBDA 재계산·실행 취소·저장 후 재열기 및 계산 상태 창의 모바일 배치를 검사합니다. 사용자 파일을 사용하지 않으며 외부 요청과 쓰기 요청을 차단합니다.

개발 서버를 실행한 뒤 설치된 Playwright와 브라우저 런타임을 사용해 실행합니다. `PLAYWRIGHT_MODULE`로 모듈 경로, `PLAYWRIGHT_BROWSERS_PATH`로 브라우저 캐시 경로를 지정할 수 있습니다.

```powershell
$env:WIXEL_URL = 'http://localhost:5178/'
$env:WIXEL_BROWSER = 'chromium' # firefox 또는 webkit
$env:WIXEL_OUTPUT = 'D:/Codex/Temp/wixel-import-formula-diagnostics-chromium'
node tools/import-formula-diagnostics.mjs
```

URL과 출력 경로의 기본값은 위와 같습니다(출력 폴더 끝은 선택한 브라우저명). `WIXEL_IMPORT_FILTER`에 시나리오 이름 일부를 지정하면 해당 검사만 실행하며, 여러 조건은 `|`로 구분합니다. 결과 JSON, 합성 XLSX와 스크린샷은 출력 폴더에 기록합니다.

## 도형 내부 텍스트 편집

`node tools/shape-inline-text.mjs`는 15개 합성 사례로 도형에서 F2·더블클릭·우클릭·문자 입력으로 본문 편집에 진입하고, 마우스 커서/선택 치환, Ctrl+B/I/U 및 리본 글꼴·크기·색의 부분 적용, 줄바꿈·조합 이벤트·일반 텍스트 붙여넣기, 로컬 Undo/Redo와 한 번의 통합 문서 Undo를 검사합니다. Alt H F C·F10 색상 키팁, 리본의 선택 텍스트 복사·잘라내기·붙여넣기, 클립보드 권한 거절과 종료 후 지연 응답도 확인합니다. 빈 편집 여백 클릭 뒤 초점 유지와 이전 조합 종료의 예약 프레임을 다음 조합 중 강제로 실행하는 커서 경합도 포함합니다. 아래 셀 보존, 보호·문서/시트 전환 시 대상 격리, 스크롤·배율·틀 고정 추적, 기존 서식 패널의 리치텍스트 보존, 합성 XLSX 다운로드→재열기, 좁은 모바일 배치도 포함합니다.

`WIXEL_URL`로 소스 또는 번들을 지정하고 `WIXEL_BROWSER=chromium|firefox|webkit`, `WIXEL_SHAPE_INLINE_OUT`(D: 출력 폴더), `WIXEL_SHAPE_INLINE_FILTER`(사례명 부분 문자열, 여러 조건은 `|`로 구분)를 사용할 수 있습니다. 한글 자판의 물리 키 코드와 Ctrl·Command 조합도 서식·실행 취소 사례에서 검사합니다. 수정 전 소스 서버에는 `WIXEL_SHAPE_INLINE_BASELINE=1`을 지정해 F2·더블클릭이 본문 커서 대신 서식 패널을 여는 상태를 기록합니다. 공개 주소의 합성 검증은 `WIXEL_SHAPE_INLINE_PUBLIC=1`과 승인된 위셀 공개 주소를 함께 지정할 때만 허용합니다. 모든 외부/API 요청과 쓰기를 차단하며 실제 업무 문서는 도구에 포함하지 않습니다. IME·붙여넣기는 합성 DOM 이벤트, 모바일은 화면 크기 시뮬레이션으로 실제 기기의 입력기·클립보드 검증을 대신하지 않습니다.


## 팝업 접근키 충돌 회귀

`node tools/popup-key-conflicts.mjs`는 공통 UI 소스 모듈을 호출하여 팝업 접근키와 별칭의 단일 할당, 표시된 괄호 키와 실행 대상 일치, 창별 명시 키와 공통 취소 기본값, 숨김·비활성·동적 항목, 부모/자식 범위, 한글 IME 및 일반 입력, 스크롤 밖 항목, 36개 초과 옵션의 Tab 접근을 확인합니다. 기존 `dialog-access-keys.mjs`도 중복 키를 재배정한 뒤 해당 표시 키가 즉시 한 명령을 실행하는 기준을 사용합니다.

```powershell
$env:WIXEL_URL = 'http://localhost:5178/'
$env:WIXEL_BROWSER = 'chromium' # firefox, webkit도 지원
$env:WIXEL_OUTPUT = 'D:/Codex/Temp/wixel-popup-key-conflicts-chromium'
node tools/popup-key-conflicts.mjs
```

소스 서버 전용이며 사용자 파일을 열지 않습니다. 외부 요청과 서버 쓰기를 차단하고 사례별 PNG, 오류 및 검사 수를 `result.json`에 저장합니다. `WIXEL_FILTER`로 사례 이름 일부를 지정할 수 있습니다. 실제 앱 및 최종 번들 단축키는 `excel-shortcut-audit.mjs`를 함께 사용합니다. 합성 한글 이벤트는 실제 iOS/Windows IME 장치 검증과 구분합니다.

### Excel Alt 호환·직접 단축키 회귀 (2026-10-05)

`excel-shortcut-audit.mjs`는 합성 문서에서 구/현대 필터 경로, 조건 보존·Undo, 물리 한글 키, 취소/입력 격리, Shift 조합, 슬라이서 Ctrl+1, End 이동, 포커스별 저장을 검증합니다. `WIXEL_BROWSER=chromium|firefox|webkit`, `WIXEL_URL`, `WIXEL_SHORTCUT_OUT`(D: 폴더), 필요시 `WIXEL_SHORTCUT_FILTER`를 지정합니다. `PLAYWRIGHT_MODULE`은 Windows 경로와 file URL을 지원합니다. 공개 합성 검사에는 승인된 workers.dev URL과 `WIXEL_SHORTCUT_PUBLIC=1`을 함께 지정합니다. 모든 API·외부 요청·원격 쓰기는 차단합니다.

`shortcut-review-regressions.mjs`는 추가 검수에서 발견한 시트 이름 입력 후 Enter 누수, 보호 문서의 저장/구조 변경, End→Alt 전환과 키 안내 표시를 검사합니다. 환경 변수는 도구 상단을 참고합니다.

`keytips.mjs`에는 새 호환 등록표의 전체 접두 경로도 포함됩니다. `ribbon-keytip-coverage.mjs`는 구 경로 별칭이 현대 리본 배지를 대체하지 않는지 실제 16탭/364조작을 검사합니다. `keytips-ime.mjs`는 활성 Alt 순서의 Process/229와 일반 한글 입력을 구분합니다. UI 테스트는 실제 iPad Bluetooth 장치나 Windows IME 드라이버 검증과 다릅니다. 검수 분모·미지원 목록은 `docs/codex/87_Excel단축키와접근키충돌.md`에 있습니다.


## 기본 행수와 확장 옵션 회귀

row-limit-mode.mjs는 기본 1,048,576행/확장 20,000,000행의 12가지 UI 사례를 검사한다. row-limit-safety.mjs는 마지막 행의 편집·메타데이터·인쇄/HTML 보존 9사례를 검사한다. WIXEL_URL로 소스/번들을, WIXEL_BROWSER=chromium|firefox|webkit으로 엔진을 지정한다. 출력과 사례 필터 환경 변수는 각 도구 첫 부분을 따른다. 승인된 공개주소의 합성 검증에만 WIXEL_ROW_LIMIT_PUBLIC=1을 추가한다. 외부 요청·API·원격 쓰기는 차단한다.

row-limit-storage.mjs는 로컬 Chromium에서 CSV 확장 행 가져오기 및 실제 WIXEL 다운로드/재열기를 확인한다. WIXEL_ROW_STORAGE_OUT은 D: 출력 폴더다. 실제 사용자 파일을 사용하지 않는다.


### 무료 웹 글꼴 회귀

`node tools/web-fonts.mjs`: 한/영 검색·400종 이후 더 보기·분류·실제 Google/CDN 바이너리 로드·모바일·셀/도형 서식·기본 글꼴·XLSX 이름 왕복·늦은 로드와 Undo/문서 전환·실패 재시도·연속 미니 서식 명령을 확인한다. `WIXEL_URL`, `WIXEL_BROWSER=chromium|firefox|webkit`, `WIXEL_OUTPUT`(D: 경로), 선택 `WIXEL_FILTER`로 실행한다. 공개검사는 `WIXEL_FONTS_PUBLIC=1`과 정확한 공개주소를 지정하고 문서 API 및 관계없는 외부 요청을 차단한다. `PLAYWRIGHT_MODULE`과 `PLAYWRIGHT_BROWSERS_PATH`는 기존 번들 런타임을 사용할 수 있다.

`node tools/update-font-catalog.mjs --cache-dir D:/Codex/Temp/wixel-font-catalog`: 공식 목록과 라이선스 확인 후 카탈로그 갱신. 자세한 검증/재생성 절차와 제외 기준은 [무료 글꼴 라이선스](../docs/font-licenses.md)를 참조한다.

### 피벗 정렬과 합계 경계 회귀

`node tools/pivot-sort-boundaries.mjs`는 실제 리본·우클릭·필터·정렬 창을 조작하여 총합계·부분합·일반 데이터 경계, Undo/Redo, 슬라이서 후 정렬 유지, XLSX 저장 왕복을 검사한다. `WIXEL_URL`, `WIXEL_BROWSER=chromium|firefox|webkit`, `WIXEL_PIVOT_SORT_OUT`(D: 폴더), 선택 `WIXEL_PIVOT_SORT_FILTER`를 지정한다. 승인된 공개 주소의 합성 검사는 `WIXEL_PIVOT_SORT_PUBLIC=1`을 추가한다. API 쓰기·외부 요청 차단과 개별 스크린샷/결과 JSON을 포함한다. 기준 HEAD 및 실제 파일 비교 결과는 [피벗 정렬 검증](../docs/codex/90_피벗정렬합계보존.md)을 참조한다.

### 슬라이서 위치와 셀 앵커 회귀

`node tools/slicer-anchor-layout.mjs`는 합성 XLSX에서 파일 열기·셀 입력의 자동 높이/너비, 수동 크기, 붙여넣기·필터·기본 너비, 실행 취소, 배율·창 크기·스크롤·틀 고정을 검사한다. `SLICER_ANCHOR_BASELINE=HEAD`는 수정 전 app.js를 메모리에서 제공하고 실패를 별도 `.local/slicer-anchor-layout/baseline`에 보관한다.

`node tools/slicer-real-layout.mjs "D:/원본 파일.xlsb"`는 명시한 실제 파일의 슬라이서와 표·셀 경계를 배율별로 비교한다. `WIXEL_URL`로 소스/번들의 로컬 서버를, `SLICER_ANCHOR_OUTPUT` 또는 `SLICER_REAL_OUTPUT`으로 D: 출력 폴더를 지정한다. 두 도구 모두 별도 브라우저에서 API·외부 요청·쓰기 요청을 차단하며 실제 파일을 수정하지 않는다. 검사 범위와 Excel COM 미검증 한계는 [94. 슬라이서 위치와 셀 앵커](../docs/codex/94_슬라이서위치와셀앵커.md)에 기록한다.
합성 공개 검증에는 `SLICER_ANCHOR_PUBLIC=1`과 정확한 위셀 공개 주소를 함께 지정한다. 실제 업무 파일 검사 도구는 로컬 전용을 유지한다. 수정 전 회귀 기준은 `SLICER_ANCHOR_BASELINE=79a0f97`이다.

### 아주 좁은 열의 숫자 표시 회귀

`node tools/narrow-column-layout.mjs`는 숫자·날짜·백분율·General 소수/지수와 텍스트·오류를 폭·배율별로 검사한다. 숫자가 들어가지 않는 폭의 `#` 개수, 한 글자도 안 들어갈 때의 빈 표시, 원본 숫자 보존과 열 확대·실행 취소를 확인한다. `WIXEL_URL`, `NARROW_COLUMN_OUTPUT`(D: 폴더), `NARROW_COLUMN_BASELINE=ff60c83`으로 서버·출력·수정 전 표시를 지정한다. 실제 원본은 `REAL_WORKBOOK`으로 지정하며 로컬 서버만 허용한다. 공개 합성 검사는 정확한 URL을 `WIXEL_ALLOWED_TEST_URL`에도 지정한다. API·외부 요청·쓰기 요청을 차단하며 결과와 제한은 [95. 아주 좁은 열의 숫자 표시](../docs/codex/95_아주좁은열의숫자표시.md)에 기록한다.

### 피벗 옵션과 결과 겹침 회귀

`node tools/pivot-options-overlap.mjs`는 native XLSX 합성 파일을 격리 브라우저에서 열어 자동 맞춤 옵션과 정렬·필터·새로 고침의 너비, 행·열·값 확장 충돌, 슬라이서 일괄 변경, 거절된 슬라이서 새로 고침의 stale 캐시 보존, Undo/Redo 및 자동 갱신 재시도를 검사한다. `WIXEL_URL`, `WIXEL_BROWSER=chromium|firefox|webkit`, `PIVOT_LAYOUT_OUTPUT`(D: 폴더), 선택 `PIVOT_LAYOUT_FILTER`를 지정한다. `PIVOT_LAYOUT_BASELINE=c95cb90`은 수정 전 app/xlsx/xlsb를 메모리에서 제공하며 예상 실패의 결과를 별도 폴더에 보관한다.

사용자가 지정한 이번 XLSB는 `REAL_WORKBOOK`으로 로컬 서버에서만 검증한다. 14개 피벗의 옵션·native XLSX 왕복과 검색어 시트 정렬·Undo 후 열 너비를 확인한다. 공개 합성 검사에는 정확한 URL을 `WIXEL_ALLOWED_TEST_URL`에도 지정한다. 모든 브라우저 검사에서 API·외부 요청·쓰기 요청을 차단한다. `TEMP`·`TMP`와 Playwright 브라우저 캐시도 D:를 사용한다. 범위와 한계는 [96. 피벗 옵션과 겹침 방지](../docs/codex/96_피벗옵션과겹침방지.md)에 기록한다.

### 전체 행 삽입·피벗 버튼·그림 앵커 회귀

`node tools/pivot-row-insert.mjs`는 합성 문서에서 실제 삽입 탭의 시트 행 삽입 버튼과 홈 메뉴의 연속 삽입, 격자·리본 버튼·탭·상태 표시줄 버튼 초점의 F4, 선택 유지 및 Undo/Redo를 검사한다. 두 피벗 위/사이에 20회 전체 행을 삽입한 뒤 페이지·행·열 필터 버튼 개수와 셀 좌표, 소스 정의 및 확대·스크롤·틀 고정을 확인한다. 작은 행 머리글의 25/50/75% 부근 선택점은 실제 MouseEvent 정수 CSS 좌표와 가운데 60% 범위를 함께 고려하고 입력 좌표를 결과에 기록한다. 보호·마지막 행 제한의 원자적 거절, 수식 편집 F4의 참조 전환, 메뉴/입력 초점 격리 및 슬라이서 충돌 경고의 8초 표시도 포함한다.

합성 PNG의 oneCell/twoCell/absolute 배치와 혼합 그룹·슬라이서의 행 삽입 및 화면 좌표, 20회 삽입+피벗 새로 고침의 oneCell 크기 보존, 행 삽입→수동 피벗 서식→새로 고침→Undo/Redo의 셀·정의·캐시 상태와 autoRefresh가 원본 편집 이력에 합쳐지는 원자성을 확인한다. 자동 갱신 사례는 처음 원본 버전을 등록하는 350ms debounce가 끝난 뒤 편집하도록 준비 단계에서 500ms 대기한다. 값 필드 확장→Undo→앞 행 삽입→새로 고침에서 이전 역할 캐시가 부활해 외부 셀과 수동 서식을 지우지 않는지도 검사한다. 부분 셀 밀기는 전체 행 삽입과 다른 기능으로 이 도구의 필수 범위에 포함하지 않는다.

`WIXEL_URL`, `WIXEL_BROWSER=chromium|firefox|webkit`, `WIXEL_ROW_INSERT_OUT`(D: 폴더), `WIXEL_ROW_INSERT_FILTER`(사례명 부분 문자열, `|`로 여러 조건)를 지정할 수 있다. 로컬 소스 서버에서 `WIXEL_ROW_INSERT_BASELINE=316020f`를 지정하면 기준 app/view/header-resize/workbook/ui 응답만 메모리에서 제공한다. 정확한 공개 합성 URL은 `WIXEL_ALLOWED_TEST_URL`에도 같은 값으로 지정해야 한다. 실제 사용자 파일을 열지 않으며 API·외부 요청·원격 쓰기를 차단한다. `PLAYWRIGHT_MODULE`, `PLAYWRIGHT_BROWSERS_PATH`, `TEMP`·`TMP`는 기존 런타임과 D: 캐시·임시 폴더를 사용한다. 결과 JSON은 before/after 초점·선택·이력 및 실제 입력 이벤트 trace를 포함하고 스크린샷은 출력 폴더에 저장한다. 확대와 고정은 브라우저 시뮬레이션이며 물리 장치 입력 검증과 구분한다.

### 행 삽입 F4 입력 경로와 합성 성능 회귀

`node tools/row-insert-performance-ui.mjs`는 19개 합성 사례에서 idle 셀 편집기·격자·리본·시트 탭 초점의 F4, 물리 `code=F4`와 Process/229 표시, 실제 조합 중 입력 보존, 반복 키·빠른 리본/F4·키팁 및 편집 경계를 검사한다. 브라우저 프로토콜의 연속 입력과 고의적 긴 작업은 물리 키보드 검증과 구분하며 키 누락 원인으로 단정하지 않는다. 기본 대량 합성 시트는 2,500행·15,000셀·10,000수식이며 동작 시간은 기계별 고정 통과 기준 없이 기록한다.

`WIXEL_URL`, `WIXEL_ROW_PERF_OUT`(D: 출력), `WIXEL_ROW_PERF_FILTER`(`|`로 사례명 부분 문자열), `WIXEL_ROW_PERF_ROWS`(10~20,000행)를 지정한다. `WIXEL_ROW_PERF_BASELINE=b7ced1f`는 로컬 서버에서 app/workbook/view/header-resize/keyboard-shortcuts/ui를 메모리 제공하여 변경 전 행동을 비교한다. API·외부 요청·원격 쓰기를 차단하고 공개 합성 검사는 정확한 `WIXEL_ALLOWED_TEST_URL`이 필요하다. 이 도구는 개인 업무 파일을 열거나 업로드하지 않는다. `PLAYWRIGHT_MODULE`, `PLAYWRIGHT_BROWSERS_PATH`, `TEMP`·`TMP`는 기존 런타임과 D: 캐시·임시 폴더를 사용한다. JSON은 입력 전달·실제 처리 후 초점/편집/선택/Undo 상태, 단계별 호출 시간 및 긴 작업을 기록한다.
### 복사한 전체 행 붙여넣기·삽입 회귀

`node tools/copied-rows.mjs`는 별도 Source 시트의 10개 전체 행(혼합 값·수식·꼬리 빈 행 또는 모두 빈 행)을 복사하여 Ctrl+V·리본 붙여넣기로 덮어쓰기하고, 우클릭 ‘복사한 셀 삽입’·Ctrl+Shift+=·홈 삽입 메뉴로 기존 행을 아래로 이동하는 실제 UI 경로를 검사한다. 전체 행 수·수식 참조·먼 열의 기존 값 제거·행 높이와 서식·원본 보존·단일 Undo/Redo와 전체 결과 행 선택 복원, 연속 삽입과 F4, 대상=복사 원본, 상대/절대참조 및 Undo 후 F4, 보호와 마지막 행 거절을 확인한다. 200만 셀 사전 제한은 덮어쓰기의 원자적 거절과 넓은 기존 열이 삽입 뒤 복사 영역 밖으로 이동하는 정상 사례를 나눠 검증하고, 병합 일부 교차 거절과 내용 있는 셀만 붙여넣기의 먼 열 값·조건부 서식·유효성 검사·높이 보존을 확인한다. 아래 피벗의 시작·필터 버튼 좌표와 oneCell 그림의 이동 및 새로 고침·55/100/200% 화면 좌표도 포함한다. 삽입 탭의 일반 시트 행 삽입은 빈 행 삽입을 유지하는 대조 사례다. 4열 폭의 전체 행 사본에서 열 너비만 붙여넣기하면 앞 4열만 변경되고 먼 열의 너비가 반복 적용되지 않는지도 확인한다.

`WIXEL_URL`, `WIXEL_BROWSER=chromium|firefox|webkit`, `WIXEL_COPIED_ROWS_OUT`(D: 폴더), `WIXEL_COPIED_ROWS_FILTER`(사례명 부분 문자열, 여러 조건은 `|`)를 지정한다. 로컬 소스 서버에서 `WIXEL_COPIED_ROWS_BASELINE=45745d9`를 지정하면 기준 app/paste-special 응답을 메모리에서 제공하여 변경 전 행동을 기록한다. 모든 문서는 합성이며 개인 파일과 OS 클립보드는 사용하지 않는다. 키 입력은 Playwright이고 native 복사/붙여넣기 데이터 전달은 모의 이벤트다. API·외부 요청·원격 쓰기를 차단하며 공개 합성 검사는 정확한 `WIXEL_ALLOWED_TEST_URL`을 함께 지정해야 한다. `PLAYWRIGHT_MODULE`, `PLAYWRIGHT_BROWSERS_PATH`, `TEMP`·`TMP`는 기존 런타임과 D: 캐시·임시 경로를 사용한다. JSON과 PNG는 출력 폴더에 저장한다.

### 슬라이서 첫 선택·반복 선택 성능과 정확성 회귀

`node tools/slicer-selection-performance-ui.mjs`는 격리 Chromium에서 24만 행 열 블록 원본과 1·4·6개 연결 피벗을 만들고 첫 선택과 반복 선택의 mousedown 처리 시간, 선택 DOM 갱신 관측 시간, 두 번의 animation frame 후 상태, 긴 작업과 레이아웃·개체 렌더 호출을 기록한다. 초기 피벗·슬라이서 표시가 원본 캐시를 이미 준비하므로 ‘첫 선택’은 새 문서의 첫 상호작용을 뜻한다. 12만 행의 rollup 아래 경로, 800개 행 항목의 큰 피벗 출력, 주차×캠페인 다중 필터·Ctrl 추가 선택, clear·Alt+C와 Undo/Redo, 같은 작업의 8연속 선택과 긴 작업 중 브라우저 프로토콜 8클릭, 중복 연결, 보호·연결 오류·변경 없는 선택·피벗 겹침의 원자적 거절과 개체 선택 유지도 검사한다. 변경된 클릭은 레이아웃 한 번을 요구하되 실행 시간에는 기계별 고정 통과 기준을 두지 않는다. 프로토콜 입력은 물리 장치의 입력 손실 검증으로 단정하지 않는다.

`WIXEL_URL`, `WIXEL_SLICER_PERF_OUT`(D: 폴더), `WIXEL_SLICER_PERF_FILTER`(사례명 부분 문자열, 여러 조건은 `|`), `WIXEL_SLICER_PERF_ROWS`(60~500,000행, 기본 240,000)를 지정한다. `WIXEL_SLICER_PERF_BASELINE=4457259`는 로컬 서버에서 app/workbook/cellmap/pivot/cube/pivot-field-items/view를 Git 기준의 메모리 응답으로 제공하고 중복 레이아웃은 통과 조건 대신 비교 수치로 기록한다. API·외부 요청·원격 쓰기를 차단하며 공개 합성 검사에는 정확한 `WIXEL_ALLOWED_TEST_URL`을 함께 지정해야 한다. 개인 업무 파일이나 OS 클립보드는 사용하지 않는다. `PLAYWRIGHT_MODULE`, `PLAYWRIGHT_BROWSERS_PATH`, `TEMP`·`TMP`는 기존 런타임과 D: 캐시·임시 경로를 사용한다. JSON·PNG는 지정 출력 폴더에 저장한다.

### 필터 버튼 셀 Alt+↓와 메뉴 키보드 회귀

`node tools/filter-alt-down.mjs`는 합성 문서의 일반 필터·표 필터·피벗 행/열/페이지/Σ 버튼 셀에서 Alt 누름→아래 방향키→Alt 해제의 실제 브라우저 입력을 검사한다. 메뉴의 초기 초점과 E 검색 접근키, 방향키·Tab·Space·검색·Enter·Escape와 방향키→Enter 정렬 명령, 가상 체크리스트, 적용 후 단일 Undo/Redo, 마우스 검색 초점 보존, 배율·스크롤·틀 고정, 리본/빠른 실행 버튼 초점 전달 및 편집·조합·대화상자·다른 입력의 격리를 확인한다. 본문 셀의 기존 데이터 입력 목록과 숨긴 피벗 머리글/실제 필드가 없는 Σ의 경계, 보호·최종 표시 문서의 비활성 명령과 열린 메뉴 뒤 보호/문서 버전 변경의 거절도 포함한다.

`WIXEL_URL`, `WIXEL_BROWSER=chromium|firefox|webkit`, `WIXEL_FILTER_ALT_OUT`(D: 출력), `WIXEL_FILTER_ALT_FILTER`(사례명 부분 문자열, 여러 조건은 `|`)를 지정한다. `WIXEL_FILTER_ALT_BASELINE=7cd961c`는 로컬 소스 서버의 app/ui/view/keyboard-shortcuts/app-filter-checklist를 Git 기준의 메모리 응답으로 제공한다. 공개 합성 검사에는 `WIXEL_ALLOWED_TEST_URL`을 `WIXEL_URL`과 정확히 같은 값으로 지정해야 한다. 개인 파일과 OS 클립보드를 사용하지 않고 API·외부 요청·쓰기 요청을 차단한다. 기존 `PLAYWRIGHT_MODULE`과 D:의 `PLAYWRIGHT_BROWSERS_PATH`, `TEMP`·`TMP`를 사용하며 결과 JSON과 실패·배율별 PNG를 D: 출력에 저장한다. 키 이벤트 trace는 누름/해제와 초점·선택·편집·Undo 상태를 포함한다. Process/229·조합 이벤트는 합성이며 물리 키보드나 OS 입력기 검증과 구분한다.


### 제작자 공식 필터 글의 기능 감사

`node tools/excel-filter-video-features.mjs`는 제작자 공식 글에서 확인한 선택 셀 값·채우기색·글꼴색 필터, 현재 열 필터 해제, Ctrl+Shift+L, Alt→D→F→F, 필터 머리글 Alt+↓→E→검색→Enter와 빠른 실행 도구 모음 등록·Alt 번호를 작은 합성 문서에서 검사한다. 실제 셀 우클릭·리본 우클릭·시작 화면 설정 경로, 중복 등록 방지·제거 후 번호 재배치·재열기 보존, Undo/Redo, 원본 값·서식·다른 시트 보존과 본문 Alt+↓의 기존 데이터 입력 목록도 확인한다. 원본 영상의 재생이나 전체 내용·타임스탬프를 검증했다고 단정하지 않는다. 정렬·병합·자동 합계·PDF는 공식 글의 추가 언급으로 분리하여 빠른 실행 명령 목록만 조사하며 기능 실행 검증으로 계산하지 않는다.

`WIXEL_URL`로 소스 또는 최종 번들을, `WIXEL_VIDEO_FEATURE_OUT`으로 D: 출력 폴더를, `WIXEL_VIDEO_FEATURE_FILTER`로 사례명 일부를 지정한다(여러 조건은 `|`로 구분). 로컬 소스 서버의 `WIXEL_VIDEO_FEATURE_BASELINE=7cd961c`는 기준 버전의 모든 src JavaScript와 HTML·CSS를 메모리 응답으로 제공하며 수정 전 결과를 별도 폴더에 보관한다. 현재 소스의 JavaScript도 실행 시작 시 메모리에서 고정하고 실제 최종 번들은 제공된 서버 자산을 검사한다. 공개 합성 검사에는 `WIXEL_ALLOWED_TEST_URL`을 `WIXEL_URL`과 정확히 같은 값으로 지정해야 한다. 개인 파일과 OS 클립보드를 사용하지 않고 외부·API·쓰기 요청을 차단한다. 기존 `PLAYWRIGHT_MODULE`과 D:의 `PLAYWRIGHT_BROWSERS_PATH`, `TEMP`·`TMP`를 사용한다. 검사 수·소스 해시·라우팅 여부·제한과 오류를 summary.json에, 화면을 출력 폴더에 저장하며 브라우저는 종료한다.

### 브라우저 확대와 행·열 구조 단축키 회귀

`node tools/browser-structure-shortcuts.mjs`는 합성 문서에서 실제 머리글의 한 개·여러 행/열 선택 후 Ctrl+Shift+Equal/Minus 및 숫자 키패드 +/-로 삽입·삭제하고, 수식·서식·높이/너비·단일 Undo/Redo와 재선택 없는 F4 반복을 검사한다. 한 개·세 개 행/열은 Ctrl+Shift+Minus를 연속 세 번 누른 뒤 F4로 한 번 더 삭제하고, 매번 전체 행/열 선택 종류·범위·활성 셀, 정확한 삭제 수와 수식 보정, 네 개의 독립 이력 및 전체 Undo/Redo의 선택 복원을 확인한다. 선택 아래·오른쪽의 5행/5열 병합이 삭제 후 3행/3열 선택으로 들어오는 경계에서는 Undo/Redo가 병합 범위로 선택을 확장하지 않는지, 이어지는 단축키/F4가 정확히 세 줄만 삭제하는지도 확인한다. 셀 선택은 삽입/삭제 대화상자의 취소와 방향별 밀기, 복사한 전체 10행은 삽입과 F4 반복을 확인한다. Ctrl 단독 Equal/Minus·숫자 키패드 +/- 및 논리 plus/Meta minus는 앱의 defaultPrevented=false, 모델·이력·선택·격자 배율 불변을 요구한다. Playwright 키 이벤트로 실제 브라우저나 OS 배율이 변했다는 주장은 하지 않는다. 리본·빠른 실행 버튼 초점 전달, 보호 거절과 F4 재검사, 편집기·수식/검색 입력·대화상자·메뉴·IME·선택 도형의 격리도 포함한다.

`WIXEL_URL`, `WIXEL_BROWSER=chromium|firefox|webkit`, `WIXEL_STRUCTURE_OUT`(D: 출력), `WIXEL_STRUCTURE_FILTER`(사례명 부분 문자열, 여러 조건은 `|`)를 지정한다. `WIXEL_STRUCTURE_BASELINE=05a6acf`는 로컬 소스 서버의 app/keyboard-shortcuts를 Git 기준 메모리 응답으로 제공해 기존 Ctrl+- 삭제와 Ctrl+Shift+- 테두리 제거 충돌을 기록한다. `WIXEL_STRUCTURE_BASELINE=8ceab535`와 `WIXEL_STRUCTURE_FILTER=consecutive-CtrlShift-delete-selection`을 함께 지정하면 첫 삭제 후 단일 셀로 바뀌어 두 번째 키가 삭제 대화상자를 여는 결함을 비교한다. 이 검사의 기준/소스/컴파일/공개 결과는 `WIXEL_STRUCTURE_OUT=.local/structure-delete-selection/<단계>`처럼 별도 D: 폴더에 저장한다. 공개 합성 검사는 `WIXEL_ALLOWED_TEST_URL`을 `WIXEL_URL`과 정확히 같은 값으로 설정해야 한다. API·외부 요청·원격 쓰기를 차단하며 개인 파일과 OS 클립보드는 사용하지 않는다. 복사는 모의 native 이벤트로 전달하고 조합 이벤트는 물리 IME 장치 검증과 구분한다. 기존 `PLAYWRIGHT_MODULE`과 D:의 `PLAYWRIGHT_BROWSERS_PATH`, `TEMP`·`TMP`를 사용한다. JSON은 실제 키 누름·해제의 전달/차단, 선택·초점·편집·Undo·격자 배율 상태를 기록하며 실패 PNG도 D: 출력에 저장한다.

## 대형 XLSX 모델 보존 검사

`large-xlsx-fidelity.mjs`는 로컬 XLSX/XLSM/XLSB를 읽어 전체 논리 값·수식 원문·직접 서식, 피벗 캐시 값·항목, 시트 메타데이터와 그림 배치를 범주별 해시로 기록한다. `--roundtrip`을 지정하면 새 XLSX/XLSM을 저장하고 다시 읽어 보존 여부를 비교한다. 원본과 출력 파일·결과 폴더는 아래처럼 D:에 두며, 매 실행마다 새 출력 경로를 사용한다.

```powershell
node tools/large-xlsx-fidelity.mjs D:/Codex/Temp/wixel-samples/sample.xlsx --out D:/Codex/Temp/wixel-fidelity/run-01 --roundtrip D:/Codex/Temp/wixel-fidelity/run-01/roundtrip.xlsx --async
```

기본값은 `--async`다. `--sync`는 비교용 동기 읽기, `--source SRC_DIR`는 검사할 별도 소스 디렉터리, `--max-diffs 8`은 시트별 제한된 차이 진단 수를 지정한다. `--roundtrip` 없이 실행하면 원본 digest만 기록하므로 왕복 보존 합격으로 세지 않는다. 왕복 검사의 성공 기준은 `report.json`의 `pass: true`와 종료 코드 0이다. 코드 1은 비교·보존 실패, 2는 실행 오류 또는 미완료다.

원본 가져오기·저장과 재가져오기는 **각 자식 프로세스가 종료된 뒤 순차 실행**한다. 여러 실제 파일도 한 파일씩 끝내고 다음 파일을 실행하며, 대형 워크북 두 개를 동시에 가져오지 않는다. 원본 덮어쓰기와 기존 저장 대상 재사용을 거절하고, 전후 SHA-256·크기·수정시각, 소스 JS 해시 및 왕복 단계 간 검사 도구 해시를 확인한다. 결과에는 시트명·메타데이터·그림 내용이 포함될 수 있으므로 원본·결과·진단 로그는 Git과 공개 업로드에서 제외한다.

생성된 drawing/table ID와 피벗 캐시 경로는 의미가 같은 참조 순서로 정규화한다. 일반 그룹의 `groupItems` 안에 있는 중첩 자식 ID와 부모·자식 참조도 포함하며, 위치·텍스트·참조 대상의 실제 변화는 계속 차이로 기록한다. 초기 도구는 `groupItems` 재귀를 빠뜨려 같은 그룹도 무작위 내부 ID 때문에 도면 해시가 달라질 수 있었다. 해당 버전의 기존 결과를 덮어쓰거나 통과로 바꾸지 않고, 새 도구로 새 출력 폴더에서 재검사한다. 외부 캐시의 `sourceReference.expectedCacheItemsId`도 이미 등록된 `snapshotId`/`cacheItemsId`의 같은 연결 순서로 정규화한다. 다른 캐시를 가리키거나 실제 원본·외부 관계·범위·binding이 달라지면 계속 차이로 기록한다. `scope.idNormalization`과 단계 간 도구 해시를 함께 확인한다.

이 도구는 Node 모델과 저장 정의를 검사한다. 수식 재계산은 `check.mjs`, 피벗 계산은 `pvcmp.mjs`로 별도 확인한다. 조건부 서식은 정의 보존 범위이며 Excel의 실제 표시·이미지 픽셀·브라우저 Blob 메모리 경로를 검증하지 않는다. **물리 iPad/Safari에서 대형 파일이 열리는지는 미검증**이며 Node 통과를 기기 안정성 합격으로 표현하지 않는다.

## Native Excel 표본과 모델 레이아웃 비교

`--native-targets TARGETS.json`을 대형 파일 검사에 추가하면 이미 가져온 모델 하나에서 지정한 셀의 직접 서식·논리 좌표, 행·열 크기와 그림 배치를 `original/native-model-samples.json`에 저장한다. 왕복 검사는 `reimport/`에도 같은 표본을 기록한다. 이 옵션은 Excel을 실행하지 않는다. 대상 JSON은 시트 index가 0부터, 행 번호가 1부터이며 셀·열 주소는 A1/A 형식이다. `name`으로 시트를 지정할 수도 있고 index와 함께 지정하면 둘의 일치 여부를 확인한다.

```json
{"sheets":[{"index":0,"cells":["A1","C8"],"rows":[1,8],"columns":["A","C"]}]}
```

```powershell
node tools/large-xlsx-fidelity.mjs D:/Codex/Temp/wixel-samples/sample.xlsx --out D:/Codex/Temp/wixel-fidelity/run-02 --native-targets D:/Codex/Temp/wixel-fidelity/targets.json
powershell -NoProfile -File tools/native-layout-oracle.ps1 -InputPath D:/Codex/Temp/wixel-samples/sample.xlsx -TargetsJson D:/Codex/Temp/wixel-fidelity/targets.json -OutputPath D:/Codex/Temp/wixel-fidelity/native-01.json
node tools/native-layout-compare.mjs --native D:/Codex/Temp/wixel-fidelity/native-01.json --model D:/Codex/Temp/wixel-fidelity/run-02/original/native-model-samples.json --out D:/Codex/Temp/wixel-fidelity/compare-01.json
```

위 명령도 하나씩 종료를 기다린다. oracle은 Windows 데스크톱 Excel과 COM이 필요하며, 새 Excel 인스턴스의 PID·시작시각·빈 통합 문서 상태를 확인한 뒤 원본을 읽기 전용으로 연다. 링크 갱신·매크로 실행·자동 계산을 억제하고 새로 고침·저장을 호출하지 않으며, 기존 Excel에는 연결하거나 설정·종료하지 않는다. 전후 원본 해시·크기·수정시각을 검사하고 소유 인스턴스만 정리한다. COM 격리 확인 실패·미지원 파일·XLM 매크로 사전검사 실패는 미완료로 기록한다.

기본 표본 한도는 시트당 셀 64개·행 32개·열 64개·그림 512개다. oracle의 `-MaxCellsPerSheet`, `-MaxRowsPerSheet`, `-MaxColumnsPerSheet`, `-MaxShapesPerSheet`로 늘릴 수 있지만 모델 표본의 기본 한도도 고려해야 한다. 요청이 한도를 초과하면 제외된 표본은 미검증이므로 결과의 제한 플래그도 확인한다. 비교에서 누락·지원되지 않는 매핑 등으로 `unavailable`이 남으면 합격 처리하지 않는다. 비교 종료 코드 0은 지정한 표본의 `mismatches 0`·`unavailable 0`, 1은 차이 또는 미검증 항목, 2는 실행 오류다.

비교는 Excel 포인트를 96dpi CSS 픽셀로 환산하며 기본 좌표 허용 오차는 0.51px다. 열 너비의 문자 단위는 픽셀과 직접 비교하지 않는다. oracle은 DisplayFormat도 기록하지만 모델 비교는 직접 서식과 논리 배치를 대상으로 한다. 표본 통과는 전체 Excel 시각 동등성, 조건부 서식의 최종 화면, 차트·이미지 픽셀이나 물리 iPad 화면의 합격을 뜻하지 않는다.

## 함수 이름의 명시적 수정 검사

`node tools/formula-name-repair-browser.mjs`는 아이폰 13 크기의 합성 문서로 가져오기 안내, 함수 이름 검토·취소·적용, Undo/Redo, XLSX 다운로드·재열기, 오래된 계획·시트 보호·작성 중 입력 보존을 확인합니다. `FORMULA_REPAIR_BROWSER=webkit|chromium`, `WIXEL_URL`, `FORMULA_REPAIR_OUTPUT`(D:의 출력 폴더), 기존 `PLAYWRIGHT_MODULE`·`PLAYWRIGHT_BROWSERS_PATH`를 지정합니다. 파일 교체가 끝난 실제 통합 문서 객체를 기다립니다.

`--file`과 `--corrected`는 별도의 비공개 실제 파일 검사입니다. localhost에서만 허용하고 원본의 크기·수정 시각·SHA256을 전후 비교합니다. 화면은 비공개 출력 폴더에만 보관하고 콘솔은 검사 횟수만 기록합니다. 실제 파일·출력·화면은 커밋하지 않습니다. 공개 사이트 검사에는 파일 인수를 사용하지 않습니다. 브라우저는 항상 종료하며 모바일 크기의 WebKit 검사는 물리 아이폰 Safari 검증과 구분합니다.

## 가져온 피벗의 저장 화면 보존 회귀

`node tools/pivot-import-presentation.mjs`는 작은 합성 XLSX를 실제 파일 선택 입력으로 열어 저장된 지난달 상대 날짜 필터 결과와 13개 값 필드의 부분합 캡션을 최초 열기에서 보존하는지 검사한다. 셀 값·직접 서식·병합·열 너비·행 높이·조건부 서식 정의와 실제 글꼴/채우기색, 행·페이지 필터의 Alt+↓와 E 검색, 다중 행 필드 접기/펼치기, 첫 변경의 잔상 제거, Undo/Redo 뒤 다른 필터 적용 및 두 번의 저장·재열기를 확인한다. 1만 행 합성 캐시는 로드된 실제 `computePivot` 함수를 계측하여 최초 열기의 호출 수 0을 요구하고, 사용자 필터·접기는 호출 수가 증가해야 통과한다.

`WIXEL_URL`은 로컬 서버만 허용하며 `WIXEL_BROWSER=chromium|firefox|webkit`, `WIXEL_PIVOT_PRESENTATION_OUT`(D: 폴더), `WIXEL_PIVOT_PRESENTATION_FILTER`(사례명 부분 문자열, 여러 조건은 `|`)를 지정한다. 개인 업무 파일을 사용하지 않고 외부·API·쓰기 요청을 차단한다. 기존 `PLAYWRIGHT_MODULE`과 D:의 `PLAYWRIGHT_BROWSERS_PATH`, `TEMP`·`TMP`를 사용한다. JSON과 PNG는 지정 출력 폴더에 보관하고 브라우저는 종료한다. 합성 브라우저 검사는 Native Excel 표시 동등성이나 물리 iPad의 대형 파일 안정성을 뜻하지 않는다.
