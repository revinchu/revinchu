# WIPOINT — Cloudflare Workers 배포

WIXEL 3 (`tabula/cloudflare`) 와 같은 계정 · 같은 Wrangler 로 배포합니다. WIPOINT 는 **정적 자산만** 씁니다 (Worker 코드 · Durable Object · KV 없음). 문서는 사용자 브라우저(IndexedDB)와 사용자가 내려받은 .pptx 파일에만 있으므로 서버에 저장되는 데이터가 없습니다.

## 배포 (PowerShell, 저장소의 wipoint 폴더)

~~~powershell
$env:XDG_CONFIG_HOME = 'D:\Codex\Config'          # WIXEL 배포 때와 같은 인증 폴더
$wrangler = 'D:\Codex\Temp\wixel3-tools\node_modules\wrangler\bin\wrangler.js'

node --test test/*.test.js
node build.mjs --cloud                                  # → dist-cloudflare/ (index.html + wipoint-<해시>.js + _headers + assets/)
node $wrangler deploy --dry-run --outdir ..\.local\wipoint-build --config cloudflare/wrangler.jsonc
node $wrangler deploy --config cloudflare/wrangler.jsonc
~~~

macOS/Linux: `npx --yes wrangler@4 deploy --config cloudflare/wrangler.jsonc` (저장소 package.json 에 wrangler 를 추가하지 않습니다).

배포 주소는 `https://wipoint.<계정 서브도메인>.workers.dev` 입니다 (WIXEL 3 이 `wixel-3.wizx.workers.dev` 이므로 `wipoint.wizx.workers.dev`). 이름을 바꾸려면 `wrangler.jsonc` 의 `name` 을 바꿉니다.

## 필요한 권한

WIXEL 3 배포에 쓴 OAuth 범위(`workers_scripts:write`, `account:read`, `user:read`)로 충분합니다. 정적 자산 업로드도 Workers Scripts Write 권한을 씁니다.

## 보안 헤더

`build.mjs --cloud` 가 `_headers` 를 만듭니다: `script-src 'self'`(인라인 스크립트 없음), `object-src 'none'`, `frame-ancestors 'self'`, `nosniff`, 해시 이름 JS 는 1년 캐시, `index.html` 은 `no-cache`. 슬라이드는 `style` 속성으로 그리므로 `style-src` 는 제한하지 않습니다.

## 같은 주소에서 WIXEL 과 함께 쓰려면 (선택)

WIXEL Worker 의 정적 자산 폴더 아래 `wipoint/` 로 `dist-cloudflare` 내용을 복사하면 `https://wixel-3.wizx.workers.dev/wipoint/` 에서 열 수 있습니다. 이때는 `_headers` 의 경로 앞에 `/wipoint` 를 붙여야 합니다. 별도 Worker 로 두는 쪽이 관리가 쉽습니다.

## 배포 전 확인

1. `node --test test/*.test.js` 통과
2. `node server.js` 로 띄운 뒤 `PLAYWRIGHT_MODULE=… node tools/smoke.mjs` → `bad 0`
3. `node build.mjs --cloud` 후 `dist-cloudflare/_headers` 와 해시 JS 확인
