# 공개 배포 전 독립 점검 (2026-10-01)

이 문서는 로컬 검증과 공식 플랫폼 제한 검토입니다. Cloudflare 운영 계정에 배포한 뒤 측정한 결과는 아닙니다. 계정 플랜 변경이나 유료 전환은 하지 않았습니다.

2026-10-04 변경: 현재 문서 한도는 32MiB입니다. 아래 20MiB 수치와 CPU 표는 당시 측정 기록입니다. 새 한도에서 합성 packed JSON 27,536,555바이트와 정확한 32MiB의 로컬 workerd 저장·읽기 SHA-256 일치, 1바이트 초과 413 및 기존 버전 유지 검사를 통과했습니다. 현재 문서와 과거본은 각각 100MiB를 유지하며, 게시 스냅숏은 20개 × 32MiB로 최대 640MiB입니다. 운영 성능이나 요금제 한도 통과를 뜻하지 않습니다.

## 확인한 동작

- 로컬 workerd의 정적 앱을 CSP 우회 없이 Chromium에서 열었습니다. 콘텐츠 해시가 붙은 외부 모듈 1개만 로드하고 인라인 스크립트는 없으며, 앱 부팅 중 pageerror는 0개였습니다.
- 응답의 script-src 'self' 및 script-src-attr 'none' 아래에서 합성 인라인 script와 img onerror를 추가했습니다. 두 실행은 모두 차단됐고 script-src-elem/script-src-attr 위반 이벤트를 각각 확인했습니다.
- 별도 Miniflare/workerd 인스턴스에서 API_RATE, FETCH_RATE, VAULT_RATE 바인딩을 하나씩 누락해 모두 503/RATE_UNAVAILABLE을 확인했습니다. 요청 제한을 1회로 설정한 별도 인스턴스에서는 첫 요청 200, 다음 요청 429 및 Retry-After: 60을 확인했습니다.
- Worker/DO/프록시와 브라우저 저장 모듈에는 복구키·문서 본문을 출력하는 console 호출이 없습니다. 복구키는 요청 헤더에 있고 문서 및 게시 ID와 분리됩니다. 원문 키 대신 SHA-256 해시로 DO를 고릅니다.
- config는 invocation 로그와 trace를 명시적으로 비활성화하고 로그 query string 제거를 켰습니다. 앱 코드가 기록하지 않는다는 것과 플랫폼의 모든 내부 운영 기록이 없다는 것은 다릅니다.
- 앞선 실제 workerd 25개 HTTP 검사와 정확한 20MiB 문서 저장/다운로드 SHA-256 일치, 1바이트 초과 413 결과는 유효합니다. 새 배포/수정 후에는 해당 integration 스크립트를 다시 실행해야 합니다.

## CPU 해석

공식 [Workers 제한](https://developers.cloudflare.com/workers/platform/limits/)은 일반 Workers Free HTTP 요청 CPU를 10ms로 명시합니다. 네트워크·스토리지 대기 시간은 CPU에 포함되지 않습니다. [SQLite Durable Objects 제한](https://developers.cloudflare.com/durable-objects/platform/limits/)은 DO 요청의 기본 CPU를 30초로 별도 명시합니다. Workers 제한 문서도 무거운 처리를 DO로 옮기는 방법을 안내합니다.

현재 Worker는 인증 해시·요청 제한·라우팅·RPC 스트림 전달을 처리합니다. 큰 문서의 UTF-8/JSON 문법 검사, 1MiB 청크 복사, SQLite 저장, 백업 청크 생성은 DO에서 수행합니다. 따라서 큰 문서 처리의 CPU를 일반 Worker의 10ms 한도와 동일시해서 20MiB 한도를 낮추지는 않았습니다.

이 PC의 Node 24 + 실제 메모리 SQLite에서 동일한 VaultStore.put을 측정한 값은 다음과 같습니다.

| 합성 문서 | 바이트 | 로컬 프로세스 CPU | 경과 시간 |
| --- | ---: | ---: | ---: |
| 긴 JSON 문자열 | 20,971,520 | 203ms | 194ms |
| 셀 객체 60만 개와 문자열 | 20,971,520 | 437ms | 415ms |

이는 로컬 처리 비용의 참고치이며 Cloudflare 운영 CPU나 무료 계정 성공률을 증명하지 않습니다. 로컬 workerd에서 20MiB HTTP 왕복은 약 5.7초였으며 이 값 역시 CPU 시간이 아닌 전체 경과 시간입니다. 현재 증거로는 20MiB를 즉시 낮출 사유가 확인되지 않았습니다.

무료 운영 배포 후에는 합성 1MiB/5MiB/20MiB 문서를 점진적으로 확인하고 Metrics의 1102/exceededCpu를 확인해야 합니다. 계정 제한으로 반복 실패한다면 먼저 Worker 쪽 스트림 전달/외부 가져오기 비용인지 DO 쪽 JSON 복잡도인지 구분해야 합니다. 임시로 문서 한도를 5MiB로 낮추는 것은 운영 데이터가 있을 때 선택할 수 있는 완화책이며, 그 경우 API health·공유 한도·프런트 설명·테스트를 함께 바꿔야 합니다. 가입 수나 파일 수에 따른 총 스토리지 제한도 별도로 적용됩니다.

## 남는 운영 범위

Rate Limit binding은 지역별 제한이며 전 세계 단일 비용 상한이 아닙니다. 무작위 복구키로 새 보관함을 만드는 쓰기 요청은 현재 허용됩니다. 광범위한 익명 공개 서비스에는 가입/봇 방어·모니터링을 별도로 고려해야 합니다. 이 판단은 CSP나 문서 크기 조정만으로 해결되지 않습니다.

Cloudflare 공식 [Workers Logs](https://developers.cloudflare.com/workers/observability/logs/workers-logs/)의 invocation 로그 설정을 참고했습니다. 실데이터를 로깅해서 CPU를 측정하지 마세요. 운영 진단을 위해 일시적인 로그나 trace를 켤 때는 합성 보관함과 합성 문서만 사용하고 원래 설정으로 되돌려야 합니다.
