# 65. iPad Bluetooth 키보드 단축키

2026-10-03. 기준 릴리스 `97601ee55e04caf911528c7bbf8bb20a8f976428`.

## 사용자 보고와 구분

Keys-To-Go 2를 연결한 iPad Safari에서 셀 우클릭 복사·붙여넣기는 동작하지만 Ctrl+C는 동작하지 않고, 왼쪽 Alt+V는 동작한다는 보고다. 키 각인과 브라우저에 전달된 보조키는 같다고 가정하지 않는다. 실제 기기 이벤트를 받지 않았으므로 키보드 고장, OS 설정, 앱 종료 원인을 확정하지 않는다.

[Apple](https://support.apple.com/en-us/102393)의 iPad 기본 단축키는 Command+C/V다. [보조키 설정](https://support.apple.com/en-gb/guide/ipad/ipaddf61a0c2/ipados)은 설정 → 일반 → 키보드 → 하드웨어 키보드 → 보조 키에서 변경한다. [Logitech 범용 모델](https://support.logi.com/hc/en-us/articles/22241277704215-Getting-Started-Keys-To-Go-2)은 OS에 따라 보조키 역할을 자동 변경하고, Fn+O 3초로 Apple 배열을 선택한다. 이 Fn 조합을 Fn 없는 iPad 전용판에 일괄 안내하지 않는다. 사용자의 왼쪽 Alt가 Command로 전달될 가능성은 가설이다.

## 수정 설계

- **레이아웃과 동작 분리:** 셀 Ctrl/Command C·X·V의 내부 사본 처리는 모바일 작업 모드 또는 Apple 터치 기기에 적용한다. iPadOS의 Mac 사용자 에이전트와 넓은 화면·마우스·모바일 OFF에도 동작한다. 일반 데스크톱의 native 클립보드 경로는 유지한다.
- **키 식별:** `keyboard-shortcuts.js`의 `shortcutCode`는 물리 code 우선, code 미제공/Unidentified에서 영문 key, 마지막으로 영문·숫자의 legacy keyCode/which를 사용한다. 한글 key 자체를 무조건 영문으로 치환하지 않는다. 조합 중·229·Dead·Process는 대체 판단하지 않는다.
- 격자 명령과 Alt 리본 키, 개체 C/X/D와 편집 중 저장에 대체 코드 처리를 연결한다. 실제 Alt를 Ctrl로 바꾸지 않는다. 보조키의 OS 매핑을 앱에서 추측하여 덮어쓰지 않는다.
- **사본 신뢰 경계:** 넓은 iPad에서도 blur/hidden 뒤 외부 자료에 이전 내부 사본 서식을 재사용하지 않는다. 중복 native 이벤트, 보호·Undo, 편집 및 IME 경로는 그대로 검증한다.
- **기기 내 확인:** 모바일 작업 도구 → 키보드 단축키 확인, 또는 도움말 → 바로 가기 키 → 키보드 단축키 확인. 입력란이 아닌 시험 영역을 사용하여 화면 키보드를 요청하지 않는다. 현재 보조키·허용된 물리 키 이름·C/X/V/Z 수신 상태만 표시한다. 문서와 클립보드를 변경하지 않고 입력 글자·이력을 저장하거나 전송하지 않는다.
- 확인 창의 시험 영역에서만 팝업 접근키 가로채기를 제외한다. Tab·Escape와 한글 조합은 기본 처리한다. OS가 브라우저 전에 소비한 키는 진단할 수 없다고 명시한다.

## 검증

브라우저 합성/자동 입력은 실제 Keys-To-Go 2 Bluetooth 또는 iPadOS의 OS 단축키 배정 검사가 아니다. 실기기에서 사용자가 보는 실제 보조키 수신 확인과 셀 복사·붙여넣기를 별도로 확인해야 한다.

최종 JavaScript `wixel-5791fdb46af4ceb3.js` (4,634,567 bytes):

| 검사 | 결과 |
|---|---|
| 전체 Node | 1,170/1,170 |
| 단일 HTML·Cloudflare 빌드 | 모두 성공 |
| 기존 단축키 | 114/114, 페이지 오류 0 |
| 리본 명령 smoke | 387개, bad 0 |
| 셀 복사 Chromium·WebKit | 각각 34/34·243검사 |
| 키 수신 확인 창 Chromium·WebKit | 각각 9/9·103검사 |
| 기존 외부 키보드 Chromium·WebKit | 각각 14/14·164검사 |

최종 키보드 UI는 총 114시나리오·1,020검사, 페이지 오류·원격 쓰기 0이다. 확인 창은 클립보드 접근 0과 문서·Undo 불변을 확인했고, 양 엔진의 320px PNG를 직접 확인하여 잘림/겹침 없음을 검사했다. 모든 OS 클립보드는 모의한다.

기준 e7ef 번들에서 유효 기능 경계 9조건 중 3실패(한글 code 누락 Ctrl/Meta 2개, iPad 모바일 OFF 1개)를 재현했고 수정본에서는 모두 통과한다. 강제 body blur 뒤 API `selectCell` 호출은 실제 마우스 선택처럼 포커스를 복구하지 않으므로, 그 초기 도구 실패를 제품 결함으로 주장하지 않는다. 최종 도구는 실제 격자 클릭 뒤 복사 복귀를 검사한다.

로그: `.local/ipad-keys-{tests,regression-final,smoke-final}.log`, `.local/mobile-clipboard-ipad-compiled-{chromium,webkit}.log`, `.local/ipad-keys-{check,hardware}-compiled-{chromium,webkit}.log`. 기기 내 확인 창·입력 도구는 `tools/keyboard-check.mjs`, `tools/mobile-clipboard.mjs`, `tools/mobile-hardware-keyboard.mjs`다. 실제 iPad Bluetooth 검증은 완료하지 않았다.
