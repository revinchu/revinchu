# 검증 도구 (tools/)

`npm test` 외에 브라우저 동작과 엑셀 파일 회귀를 확인하는 스크립트입니다.
실제 파일을 비교할 때는 **수정 전에 기준 결과를 저장하고, 수정 후 같은 파일과 명령으로 다시 측정**하세요.
최신 실행 결과와 미검증 범위는 [`09_WIXEL3_구현과검증.md`](../docs/codex/09_WIXEL3_구현과검증.md), 원본 인계 검증은 [`06_검증결과.md`](../docs/codex/06_검증결과.md)에 기록합니다.

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
네 브라우저 도구는 초기화 스크립트에서 `window.TABULA_STATIC = true`를 설정해
서버 문서 자동 저장·복원으로 검사 데이터가 바뀌는 것을 막습니다. 도구의 성공은 서버 API 인증·보안 검증을 의미하지 않습니다.
`ui-regressions.mjs`도 별도 브라우저에서 합성 문서를 사용합니다. 인증 화면 시나리오는 API 응답을 모의하며, 실제 서버 API는 `npm test`의 서버 테스트에서 검증합니다.

### Git Bash 사용 시

Windows Git Bash에서도 같은 D: 경로를 사용합니다. 아래 환경 변수는 명령을 실행하는 각 터미널에 적용합니다.

```bash
cd /d/Codex/Workspaces/wixel/tabula
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
| `keys.mjs` | 단축키 66개 확인. 3개는 실행 오류만 확인한다고 출력 | `node tools/keys.mjs` | 요약의 `bad`·`pageErrors`가 모두 0, 종료 코드 0 |
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
| `node tools/format-regressions.mjs` | 차트 축/계열 연속 변경, 3D 회전·깊이, 도형 채우기·효과·텍스트와 페이지 오류 |
| `node tools/security-regressions.mjs` | CSP 우회 테스트 환경에서 HTML/SVG 공격 차단, 인쇄·링크·정상 차트/도형 표현 |
| `node tools/vault-ui-integration.mjs` | 로컬 workerd의 보관함 연결·저장·다른 기기 복원·게시 관리 UI |
| `node tools/wixel3-ui.mjs` | 시작·저장 화면, Ctrl+Space, QAT, 개인 보관함 충돌, Google Sheets, 피벗과 셀 서식 |
| `node --test cloudflare/backend.test.js` | 개인 보관함 격리·버전·청크·게시 API |
| `node cloudflare/integration.mjs` | 실제 로컬 workerd API; 별도 Cloudflare dev 서버 필요 |

`format-regressions.mjs`는 저장소 `.local/`에 차트와 도형 화면을 캡처합니다. Cloudflare 설정·실행 환경과 한도 검사 명령은 `cloudflare/README.md`를 확인하세요. 실제 공개 배포 후의 동작은 로컬 통과와 별도로 확인해야 합니다.
