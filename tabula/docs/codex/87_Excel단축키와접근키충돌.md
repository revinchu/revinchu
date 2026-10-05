# 87. Excel 단축키 호환과 팝업 접근키 충돌

2026-10-05. 사용자 보고: Alt → D부터 인식되지 않아 Alt D F F 필터를 실행할 수 없음, 팝업 C 접근키 중복. 기준 커밋 1bd7bd89291993923426c4d310a880e7bfaa84d9. 작업 경로 D:/위셀/wixel-3, 임시 결과는 D:/Codex/Temp에만 저장했다.

## 변경

- 실제 설치된 한국어 Excel의 메뉴 명령 식별자·접근키를 대조해 이전 메뉴 경로 91개, 공식 현대 리본 경로 3개를 연결했다. 기존 ES·AY3 중복을 정리하며 실제 실행 등록표에서 고유 경로를 유지한다.
- D/E/I/O/T/V와 중간 단계에서 다음 글자를 안내한다. 예: D → F → F 필터 전환, D → F → S 조건 지우기, D → F → A 고급 필터. 불완전한 접두어에서는 실행하지 않고 Escape/Backspace로 되돌아간다.
- AY를 조기 실행하던 연결을 AY3로 정정했다. HSU 정렬, HOR 시트 이름 변경을 추가했다. 기존 ES는 유지한다. 메뉴/입력/명령을 구분하고 구 경로 별칭이 현대 리본 대표 배지를 대체하지 않게 했다.
- Alt+Shift+F1 새 시트, Ctrl+Shift+F1 전체 화면, Ctrl+Shift+G 통계, Ctrl+Shift+S 다른 이름 저장, Ctrl+F12 열기, Shift+F12 저장, Ctrl+Shift+F12 인쇄, Ctrl+Alt+=/- 배율을 연결했다. 미등록 Shift 조합이 Ctrl만 누른 다른 기능으로 떨어지는 것을 막았다.
- End → 방향키 데이터 경계 이동, Shift와의 범위 확장 및 Escape/마우스/시트 전환 취소를 추가했다.
- 슬라이서 Ctrl+1은 크기 및 속성을 연다. 도형 Ctrl+1도 개체 서식으로 명시 연결했다. 도형은 수정 전에도 run의 후속 분기를 통해 정상적으로 열렸으므로 신규 결함으로 집계하지 않는다.
- 리본 포커스의 Ctrl+S/Shift+S는 물리 키를 판별하고 이미 처리한 이벤트를 재실행하지 않는다. 활성 Alt 순서의 한글 Process/229 키는 물리 위치로 읽되 일반 조합 입력은 명령으로 처리하지 않는다.
- 팝업 주 접근키·별칭을 같은 범위에서 단일 대상에 배정한다. 명시 키 충돌은 다른 키로 표시·실행하며 창별 명시 키가 공통 기본 취소 키보다 우선한다. 36개를 넘으면 모호하게 재사용하지 않고 Tab/화살표 탐색을 사용한다. 숨김·비활성·중첩 창, 동적 옵션, 실제 표시 괄호와 실행 키를 함께 처리한다.
- 새 시트 명령도 읽기 전용/최종본/통합 문서 구조 보호를 따른다. 셀 편집·모달 입력·검색 중에는 리본 순서가 텍스트를 가로채지 않도록 기존 범위를 유지한다.

## 검증과 범위

최종 검증 수치·배포 식별자는 아래 후속 기록에 기재한다. 초기 실패 로그도 삭제하지 않았다. 팝업 C 복사는 기존 실제 셀 메뉴에서 정상인 경우도 있었고, 범용 중복 명시 키/별칭의 모호한 순환 및 충돌이 확인된 부분을 수정했다.

검증 도구:
- test/excel-keytip-compat.test.js: 연결 대상, 출처, 실제 리본 병합, 중복·접두어·대표 배지 검사.
- test/excel-direct-shortcuts.test.js: 보조키·물리 키·IME·AltGraph 격리.
- tools/excel-shortcut-audit.mjs: 필터 조건/Undo, 순서 입력/취소, 입력 격리, 직접 조합, 슬라이서, End와 포커스별 저장.
- tools/popup-key-conflicts.mjs: 소스 공통 팝업의 명시 키/별칭 충돌·숨김·중첩·동적 표시.
- tools/keytips.mjs, ribbon-keytip-coverage.mjs, keytips-ime.mjs 및 기존 팝업 도구도 회귀 실행.

Microsoft 문서와 현재 코드의 직접 조합을 추가 대조했으나 아래는 미구현 또는 별도 검증 범위다. Shift+F8 비연속 추가 선택, Ctrl+6 개체 표시 모드, Ctrl+8 윤곽 기호, F6 영역 초점 순환, Ctrl+Shift+PageUp/Down 시트 그룹, 스레드 댓글, VBA IDE·Power Query·Office 추가 기능. 최신 Microsoft 365의 Ctrl+F2 설명과 이전 Excel 인쇄 미리보기 차이도 있으며 위셀의 기존 Ctrl+F2 인쇄 경로를 유지했다. OS/브라우저가 앱에 전달하지 않는 예약 키를 보장하지 않는다. 실제 iPad/iPhone Bluetooth 키보드 및 Windows 한글 IME 드라이버는 이번 자동화의 검증 대상이 아니다.

자동 생성 Zxx를 Excel의 기본 단축키로 계산하지 않는다. 전체 Excel 기능·단축키 100% 동등성을 주장하지 않는다. 아래 전수 인벤토리는 명시된 설치 버전·구 메뉴 163개 경로라는 분모에 한정한다.

---

# Excel Alt 경로 호환 검사

검사일: 2026-10-05. 대상 소스: D:/위셀/wixel-3/tabula. 원본 업무 파일을 이 인벤토리에 포함하지 않았다.

## 증거와 분모

- 설치된 한국어 Microsoft Excel 16.0, build 14334, LCID 1042에서 새 비표시 Excel 인스턴스의 Worksheet Menu Bar를 읽었다. Workbook은 열지 않았으며 읽은 정보는 Caption, Id, Type, 부모 경로뿐이다.
- 원시 증거: excel-legacy-commandbars.json. 메뉴 캡션의 & 다음 접근키를 부모 경로와 연결했다. 이 정보는 해당 Excel 버전 메뉴의 실제 명령 식별 근거이며 모든 플랫폼에서 실제 키 입력을 실행한 결과는 아니다.
- 전체 280개 control record, 실행 leaf 247개. E/O/I/D/V/T의 영숫자 경로 범위는 **169개 leaf record / 163개 고유 path**이다. Record와 path 수를 혼용하지 않는다.
- 169와 163의 차이는 중복 path 6개이다. EAM은 구 메모 지우기 ID 874만 연결했다. IA/IPO/TB/TK/DDP는 서로 다른 기능이 같은 경로를 가지므로 임의 선택하지 않았다.
- 구현 연결: **구 경로 91/163개 (55.8%)**, 별도 공식 현대 리본 경로 3개. 합계 **94개 호환 경로**. 이것은 단축키 연결 비율이며 Excel 전체 기능 동등성 점수가 아니다.
- 범위 밖 78개 leaf record는 F/H/W의 63개와 구 개체 동적 verb 자리표시자 EO< 15개이다. 현재 리본·운영체제 명령 충돌 때문에 이를 구 경로 실행 분모에 넣지 않았다.
- 현재 Excel 리본의 모든 명령, 대화상자 내부 단축키, 개체 상황별 단축키, 브라우저·OS 접근성 단축키는 CommandBars 분모와 별개이다. 자동 생성한 Zxx는 위셀 조작 접근성을 제공하며 Excel 기본 키 동등성을 뜻하지 않는다.

## 범주별 남은 경로

| 범주 | 고유 path 수 |
|---|---:|
| Office·시스템·외부 통합 | 43 |
| 기능 미구현·부분 구현 | 24 |
| 중복 접근키 | 5 |
| 합계 | 72 |

Office·시스템 통합으로 분류한 항목도 위셀에서 지원한다는 의미가 아니다. 기능 미구현·부분 구현과 구분한 이유는 데스크톱 Office 실행 환경, 폐지된 서비스 또는 외부 연결 모델을 필요로 하기 때문이다.

| 경로 | 원래 메뉴 (control ID) | 분류와 미연결 이유 |
|---|---|---|
| Alt D D A | 데이터 범위 속성(&A)... (1951) | Office·시스템·외부 통합: QueryTable 데이터 범위 속성. |
| Alt D D D | 데이터 가져오기(&D)... (6262) | Office·시스템·외부 통합: Excel 외부 데이터 연결 파일/공급자 선택. |
| Alt D D E | 쿼리 편집(&E)... (1950) | Office·시스템·외부 통합: 외부 QueryTable 쿼리 편집. |
| Alt D D M | 매개 변수(&M)... (537) | Office·시스템·외부 통합: QueryTable 매개변수. |
| Alt D D N | 새 쿼리 만들기(&N)... (2054) | Office·시스템·외부 통합: Microsoft Query/ODBC 질의. |
| Alt D D P | 파일의 사진(&P)... (33996) / 클립보드의 사진(&P) (34002) | 중복 접근키: 그림 파일(33996) / 클립보드에서 데이터(34002) |
| Alt D D W | 새 웹 쿼리(&W)... (3829) | 기능 미구현·부분 구현: 웹에서 읽기는 있으나 Excel QueryTable 연결·갱신 옵션 전부 동등하지 않음. |
| Alt D G A | 자동 개요(&A) (904) | 기능 미구현·부분 구현: 수식 구조 기반 자동 윤곽 설정 없음. |
| Alt D G E | 설정​​(&E)... (906) | 기능 미구현·부분 구현: 윤곽 설정의 요약 방향·스타일 설정 대화상자 없음. |
| Alt D I D | 변경 내용 취소 후 새로 고침(&D) (7762) | Office·시스템·외부 통합: SharePoint 변경 취소·새로 고침. |
| Alt D I L | 서버의 목록 보기(&L) (7763) | Office·시스템·외부 통합: SharePoint 목록 보기. |
| Alt D I P | 표 게시(&P)... (7684) | Office·시스템·외부 통합: 구 SharePoint 목록 게시. |
| Alt D I U | 목록 연결 끊기(&U) (7683) | Office·시스템·외부 통합: SharePoint 목록 연결 해제. |
| Alt D I Y | 목록 동기화(&Y) (7761) | Office·시스템·외부 통합: SharePoint 목록 동기화. |
| Alt D O | 레코드 관리(&O)... (860) | 기능 미구현·부분 구현: 행 단위 레코드 입력 데이터 폼 없음. |
| Alt D P | 피벗 테이블/피벗 차트 마법사(&P) (2915) | 기능 미구현·부분 구현: 피벗 삽입은 있으나 구 피벗 마법사의 다중 통합 범위·원본 선택 단계 미구현. |
| Alt D R | 새로 고침(&R) (459) | 기능 미구현·부분 구현: 현재 외부 데이터 범위만 새로 고침하는 QueryTable 동작 없음. |
| Alt D X A | XML 확장 팩(&A)... (7784) | Office·시스템·외부 통합: 구 XML 확장 팩. |
| Alt D X E | 내보내기​​(&E)... (7432) | Office·시스템·외부 통합: Excel XML Map 기반 내보내기. |
| Alt D X I | 가져오기​​(&I)... (7433) | Office·시스템·외부 통합: Excel XML Map 기반 가져오기. |
| Alt D X P | XML 맵 속성(&P)... (7813) | Office·시스템·외부 통합: XML Map 속성. |
| Alt D X R | XML 데이터 새로 고침(&R) (7812) | Office·시스템·외부 통합: XML 연결 새로 고침. |
| Alt D X X | XML 원본(&X)... (7693) | Office·시스템·외부 통합: XML 스키마 원본 작업창. |
| Alt E A W | 통합 문서 정리(&W) (33882) | 기능 미구현·부분 구현: 통합 문서 정리 기능 없음. |
| Alt E B | Office 클립보드(&B)... (809) | Office·시스템·외부 통합: Office 공용 다중 항목 클립보드 작업창. |
| Alt E H | 하이퍼링크로 붙여넣기(&H) (2787) | 기능 미구현·부분 구현: 하이퍼링크로 붙여넣기 없음. pasteLink는 셀 수식 연결이므로 대체하지 않음. |
| Alt E I A | 시트 그룹(&A)... (869) | 기능 미구현·부분 구현: 여러 워크시트에 채우기 전용 동작 없음. |
| Alt E I J | 양쪽 맞춤(&J) (871) | 기능 미구현·부분 구현: 선택 셀에 분산 맞춤 채우기 없음. |
| Alt E I S | 계열​​(&S)... (870) | 기능 미구현·부분 구현: 연속 데이터 대화상자 없음. 차트 계열 편집은 다른 기능. |
| Alt E K | 연결​​(&K)... (759) | Office·시스템·외부 통합: Excel 외부 통합 문서 연결 관리·경로 변경·원본 열기. |
| Alt E O | 개체(&O) (961) | Office·시스템·외부 통합: OLE 개체 편집/실행. |
| Alt E O N | Office 수식으로 변환(&N) (33423) | Office·시스템·외부 통합: Office 수식 OLE 변환. |
| Alt E O V | 변환​​(&V)... (1967) | Office·시스템·외부 통합: OLE 개체 형식 변환. |
| Alt I A | 페이지 나누기 모두 원래대로(&A) (1585) / 잉크 주석(&A) (9071) | 중복 접근키: 페이지 나누기 모두 원래대로(1585) / 잉크 주석(9071) |
| Alt I G | 다이어그램(&G)... (1032) | 기능 미구현·부분 구현: SmartArt는 존재하나 구 Diagram 선택 대화상자와 기능/모델이 다름. |
| Alt I N A | 적용​​(&A)... (881) | 기능 미구현·부분 구현: 기존 수식에 정의된 이름 적용 동작 없음. |
| Alt I O | 개체(&O)... (546) | Office·시스템·외부 통합: OLE 개체 삽입. |
| Alt I P C | 클립 아트(&C)... (682) | Office·시스템·외부 통합: Office 클립 아트 카탈로그. |
| Alt I P D | 잉크 그리기 및 쓰기(&D) (9405) | Office·시스템·외부 통합: Office 잉크 개체 입력. |
| Alt I P O | 온라인 그림(&O)... (20868) / 조직도​​(&O) (702) | 중복 접근키: 온라인 그림(20868) / 조직도(702) |
| Alt I P S | 스캐너 또는 카메라(&S)... (1764) | Office·시스템·외부 통합: OS 스캐너/카메라 수집. |
| Alt I P W | WordArt(&W)... (1031) | 기능 미구현·부분 구현: WordArt 전용 갤러리·곡선 텍스트 없음. |
| Alt O A | 자동 서식(&A)... (786) | 기능 미구현·부분 구현: 구 자동 서식 대화상자 없음. 표 스타일 적용과 구별. |
| Alt O C S | 표준 너비(&S)... (888) | 기능 미구현·부분 구현: 선택 열 너비는 있으나 시트의 표준/default 열 너비 지정 기능 없음. |
| Alt O H B | 배경​​(&B)... (952) | 기능 미구현·부분 구현: 워크시트 반복 배경 이미지 기능 없음. |
| Alt O S | 스타일​​(&S)... (254) | 기능 미구현·부분 구현: 셀 스타일 메뉴는 있으나 구 스타일 정의·수정 대화상자와 다름. |
| Alt O T T | 윗주 설정(&T)... (1613) | 기능 미구현·부분 구현: 윗주 편집·표시는 있으나 전용 글꼴/맞춤 설정 대화상자 없음. |
| Alt T B | 통합 문서 통계(&B) (33780) / 통합 문서 ​​공유(&B)... (2040) | 중복 접근키: 통합 문서 통계(33780) / 통합 문서 공유(2040) |
| Alt T C | 사용자 지정(&C)... (797) | Office·시스템·외부 통합: 구 Office 명령 모음 사용자 지정. 리본 사용자 지정과 구별. |
| Alt T D | 공유 작업 영역(&D)... (7710) | Office·시스템·외부 통합: Office/SharePoint 공유 작업영역. |
| Alt T H | 텍스트 읽어주기 도구 모음 표시(&H) (7011) | Office·시스템·외부 통합: Office 텍스트 음성 변환 도구모음. |
| Alt T I | 추가 기능(&I)... (943) | Office·시스템·외부 통합: Excel COM/XLL 추가 기능 관리. |
| Alt T K | 오류 검사(&K)... (6122) / 키보드 사용자 지정(&K)... (818) | 중복 접근키: 오류 검사(6122) / 키보드 사용자 지정(818) |
| Alt T M R | 새 매크로 기록(&R)... (184) | Office·시스템·외부 통합: VBA 매크로 기록기. |
| Alt T M S | 보안​​(&S)... (3627) | Office·시스템·외부 통합: VBA 보안/신뢰 센터. |
| Alt T M V | Visual Basic Editor(&V) (1695) | Office·시스템·외부 통합: VBA IDE. 매크로 목록·원문 뷰어는 동일하지 않음. |
| Alt T N M | 지금 모임 시작(&M) (3727) | Office·시스템·외부 통합: 폐지된 NetMeeting 통합. |
| Alt T N N | 검토 끝내기(&N)... (6141) | Office·시스템·외부 통합: 구 Office 검토 라우팅 종료. |
| Alt T N S | 모임 일정(&S)... (4179) | Office·시스템·외부 통합: Outlook/NetMeeting 회의 예약. |
| Alt T N W | 웹 토론(&W) (4177) | Office·시스템·외부 통합: 구 Office 웹 토론 서비스. |
| Alt T P S | 통합 문서 보호 및 공유(레거시)(&S) (3059) | Office·시스템·외부 통합: 구 공유 통합 문서 보호. |
| Alt T R | 리서치​​(&R)... (7343) | Office·시스템·외부 통합: Office 리서치 서비스. |
| Alt T T A | 변경 내용 적용/취소(&A)... (305) | Office·시스템·외부 통합: 구 공유 통합 문서 변경 내용 적용/취소. |
| Alt T T H | 변경 내용 표시(&H)... (2042) | Office·시스템·외부 통합: 구 공유 통합 문서 변경 내용 추적/강조. |
| Alt T U M | 수식 분석 모드(&M) (6059) | 기능 미구현·부분 구현: 구 수식 분석 모드 표시 전환 없음. 수식 표시와 구별. |
| Alt T U S | 수식 분석 도구 모음 표시(&S) (892) | Office·시스템·외부 통합: 구 Office 수식 분석 도구모음 표시. 위셀은 리본에 개별 명령. |
| Alt T W | 통합 문서 비교 및 병합(&W)... (2044) | Office·시스템·외부 통합: 구 공유 통합 문서 비교·병합. |
| Alt V C | 메모​​(&C) (1594) | 기능 미구현·부분 구현: 메모 전체 표시/숨기기는 있으나 구형 메모 창 표시 전환과 다른 의미. |
| Alt V H | 머리글/바닥글(&H)... (762) | 기능 미구현·부분 구현: 페이지 설정에 머리글/바닥글 필드는 있으나 전용 진입·초기 포커스 없음. |
| Alt V K | 작업창​​(&K) (5746) | Office·시스템·외부 통합: Office 구형 작업창. |
| Alt V S | 상태 표시줄(&S) (850) | 기능 미구현·부분 구현: 상태 표시줄 표시 전환 기능 없음. |
| Alt V V | 사용자 지정 ​​보기(&V)... (950) | 기능 미구현·부분 구현: 사용자 지정 보기 저장·복원 없음. |

## 구현 매핑

| 경로 | 원래 control ID | 위셀 명령/메뉴 |
|---|---:|---|
| Alt D B | 861 | command: subtotal (부분합) |
| Alt D E | 806 | command: textToColumns (텍스트 나누기) |
| Alt D F A | 901 | command: advancedFilter (고급 필터) |
| Alt D F F | 899 | command: toggleFilter (자동 필터) |
| Alt D F S | 900 | command: clearFilter (필터 모두 표시) |
| Alt D G C | 905 | command: outlineClear (개요 지우기) |
| Alt D G G | 3159 | command: outlineGroup (그룹) |
| Alt D G H | 464 | command: outlineHide (하위 수준 숨기기) |
| Alt D G S | 462 | command: outlineShow (하위 수준 표시) |
| Alt D G U | 3160 | command: outlineUngroup (그룹 해제) |
| Alt D I C | 7193 | command: createTable (표 만들기) |
| Alt D I R | 7765 | command: resizeTable (표 크기 조정) |
| Alt D I T | 7372 | command: tblTotals (표 요약 행) |
| Alt D I V | 7375 | command: convertToRange (표를 범위로 변환) |
| Alt D L | 2034 | command: dataValidation (데이터 유효성 검사) |
| Alt D N | 863 | command: consolidate (통합) |
| Alt D S | 928 | command: sortDialog (정렬) |
| Alt D T | 862 | command: dataTable (가상 분석 데이터 표) |
| Alt E A A | 1964 | command: clearAll (모두 지우기) |
| Alt E A C | 873 | command: clearContents (내용 지우기) |
| Alt E A F | 872 | command: clearFormats (서식 지우기) |
| Alt E A H | 2955 | command: clearHyperlinks (하이퍼링크 지우기) |
| Alt E A M | 874 | command: clearComments (메모 지우기) |
| Alt E A R | 16486 | command: removeHyperlink (하이퍼링크 제거) |
| Alt E C | 19 | command: copy (복사) |
| Alt E D | 478 | command: deleteMenuKey (셀 삭제) |
| Alt E E | 313 | command: replace (바꾸기) |
| Alt E F | 1849 | command: find (찾기) |
| Alt E G | 757 | command: goto (이동) |
| Alt E I D | 372 | command: fillDown (아래쪽 채우기) |
| Alt E I L | 868 | command: fillLeft (왼쪽 채우기) |
| Alt E I R | 371 | command: fillRight (오른쪽 채우기) |
| Alt E I U | 867 | command: fillUp (위쪽 채우기) |
| Alt E L | 847 | command: deleteSheet (시트 삭제) |
| Alt E M | 848 | command: moveCopySheet (시트 이동/복사) |
| Alt E P | 22 | command: paste (붙여넣기) |
| Alt E R | 37 | command: redoOrRepeat (반복) |
| Alt E S | 755 | command: pasteSpecial (선택하여 붙여넣기) |
| Alt E T | 21 | command: cut (잘라내기) |
| Alt E U | 128 | command: undo (실행 취소) |
| Alt I B | 509 | command: insertPageBreak (페이지 나누기 삽입) |
| Alt I C | 297 | command: insertCols (열 삽입) |
| Alt I E | 295 | command: insertMenuKey (셀 삽입) |
| Alt I F | 385 | command: insertFunction (함수 삽입) |
| Alt I H | 1957 | command: insertChartAll (차트 삽입) |
| Alt I I | 1576 | command: hyperlink (하이퍼링크) |
| Alt I M | 1589 | command: editComment (메모 편집) |
| Alt I N C | 880 | command: createNamesFromSel (선택 영역에서 이름 만들기) |
| Alt I N D | 878 | command: defineName (이름 정의) |
| Alt I N P | 879 | command: pasteName (이름 붙여넣기) |
| Alt I P A | 2630 | menu: shapes (도형 메뉴) |
| Alt I P F | 2619 | command: insertPicture (파일에서 그림 삽입) |
| Alt I R | 296 | command: insertRows (행 삽입) |
| Alt I S | 308 | command: insertSymbol (기호 삽입) |
| Alt I W | 852 | command: addSheet (워크시트 삽입) |
| Alt O C A | 885 | command: autofitSel (열 너비 자동 맞춤) |
| Alt O C H | 886 | command: hideCols (열 숨기기) |
| Alt O C U | 887 | command: unhideCols (열 숨기기 취소) |
| Alt O C W | 542 | command: colWidth (열 너비) |
| Alt O D | 3058 | command: condManager (조건부 서식 규칙 관리) |
| Alt O E | 855 | command: formatCells (셀 서식) |
| Alt O H H | 890 | command: hideSheet (시트 숨기기) |
| Alt O H R | 889 | command: renameSheet (시트 이름 바꾸기) |
| Alt O H T | 5747 | command: sheetTabColor (시트 탭 색) |
| Alt O H U | 891 | command: unhideSheet (시트 숨기기 취소) |
| Alt O R A | 882 | command: autofitRowsSel (행 높이 자동 맞춤) |
| Alt O R E | 541 | command: rowHeight (행 높이) |
| Alt O R H | 883 | command: hideRows (행 숨기기) |
| Alt O R U | 884 | command: unhideRows (행 숨기기 취소) |
| Alt O T E | 1611 | command: editPhonetic (윗주 편집) |
| Alt O T S | 1614 | command: togglePhonetic (윗주 필드 표시) |
| Alt T A | 793 | command: autoCorrectOptions (자동 고침 옵션) |
| Alt T E | 857 | command: scenarioManager (시나리오 관리자) |
| Alt T G | 856 | command: goalSeek (목표값 찾기) |
| Alt T M M | 186 | command: macros (매크로) |
| Alt T O | 522 | command: options (옵션) |
| Alt T P A | 6997 | command: allowEditRanges (범위 편집 허용) |
| Alt T P P | 893 | command: protectSheet (시트 보호) |
| Alt T P W | 894 | command: protectWorkbook (통합 문서 보호) |
| Alt T S | 2 | command: spellCheck (맞춤법 검사) |
| Alt T U A | 453 | command: removeArrows (연결선 모두 제거) |
| Alt T U D | 451 | command: traceDependents (참조하는 셀 추적) |
| Alt T U E | 463 | command: traceError (오류 추적) |
| Alt T U F | 5687 | command: evaluateFormula (수식 계산) |
| Alt T U T | 486 | command: tracePrecedents (참조되는 셀 추적) |
| Alt T U W | 5686 | command: watchWindow (조사식 창) |
| Alt V F | 849 | command: toggleFormulaBar (수식 입력줄) |
| Alt V N | 723 | command: viewNormal (기본 보기) |
| Alt V P | 724 | command: viewPageBreakPreview (페이지 나누기 미리 보기) |
| Alt V U | 178 | command: fullScreen (전체 화면) |
| Alt V Z | 925 | command: zoomDialog (화면 배율) |

추가 현대 경로: Alt H S U → sortDialog, Alt A Y 3 → reapplyFilter, Alt H O R → renameSheet.

기존 기능과 연결하였더라도 원래 기능의 세부 구현이 Excel과 완전히 같다고 보증하지 않는다. 예를 들어 spellCheck는 브라우저 맞춤법 검사 기능을 사용하며 macros는 기존 위셀 매크로 처리 범위 안에서 동작한다. 이번 범위에서 새 VBA 실행기·맞춤법 사전·외부 연결 공급자를 구현하지 않았다.

## 접두어·의미 충돌 조치

- AY는 중간 단계로 남겨야 AY3를 끝까지 입력할 수 있다. 기존 terminal AY를 제거하고 AY3를 공식 동작에 연결한다.
- HF는 글꼴 하위 키 HFC/HFF/HFS/HFP/HFG/HFK/HFN의 접두어다. HF 자체를 즉시 실행 명령으로 등록하지 않는다.
- WF는 틀 고정 메뉴 WFF/WFR/WFC의 접두어이므로 구 Window→Freeze Panes의 WF를 terminal로 등록하지 않는다.
- F는 위셀 파일 메뉴, WN은 현대 기본 보기, HH는 채우기 색, HC는 복사, HO는 셀 서식 메뉴 접두어다. 구 F/H/W 접근키와 무조건 병합하지 않는다.
- DT(가상 분석 데이터 표, 862)와 DIC(일반 표 만들기, 7193)는 다르다. DP(구 피벗 마법사)는 현재 피벗 삽입 명령으로 오연결하지 않는다.
- EAH(하이퍼링크만 지우기)와 EAR(하이퍼링크 서식까지 제거)를 구별한다. EH 하이퍼링크 붙여넣기는 셀 수식 연결 붙여넣기와 다르다.
- ER은 실행 취소 복구 가능 시 redo, 그렇지 않으면 마지막 작업 반복을 실행한다. 항상 repeatLast로 처리하지 않는다.
- OHT는 현재 시트 탭 색 메뉴, TA는 optionsDialog(3)의 언어 교정/자동 고침, VZ는 배율 대화상자를 연다.

## 자동 검사 결과

- 새 compatibility 단위 검사 9개 + 기존 ribbon-keytips 6개 + keyboard-shortcuts 4개, **19/19 통과**.
- 실제 TABS·KEYTIPS·EXCEL_KEYTIP_COMPAT로 공유 registry 생성: 모든 위셀 조작 대표 키 유지, compatibility 94개 전부 연결, 중복·terminal/prefix 충돌·다른 명령 덮어쓰기 0.
- 알 수 없는 긴 경로는 짧은 명령으로 fallback하지 않는다. D/DF 입력만으로 DFF를 조기 실행하지 않는다.
- 각 legacy 항목에 해당 설치 Excel의 control ID·캡션·버전·언어, 현대 경로에는 직접 Microsoft 지원 문서 URL이 포함된다.
- 이 문서는 정적 연결과 단위 검사를 기록한다. 실제 브라우저 다중 엔진 키 입력·사용자 데이터 회귀 결과는 root 작업의 별도 로그/보고서를 따른다.

## 공식 자료

- [Microsoft Excel 단축키](https://support.microsoft.com/en-us/accessibility/excel/keyboard-shortcuts-in-excel): 대부분의 이전 Alt 메뉴 경로 유지, 전체 경로 입력 필요, 현대 리본 및 플랫폼별 차이.
- [정렬/필터 접근성 단축키](https://support.microsoft.com/en-us/accessibility/excel/use-a-screen-reader-to-sort-or-filter-a-table-in-excel): HSU, AY3.
- [워크시트 이름 바꾸기](https://support.microsoft.com/en-gb/excel/rename-a-worksheet): HOR.
- [Application.CommandBars](https://learn.microsoft.com/en-us/office/vba/api/excel.application.commandbars): 설치 Excel의 실제 command bar 메타데이터 근거.
- [CommandBarControl.Caption](https://learn.microsoft.com/en-us/office/vba/api/office.commandbarcontrol.caption): 캡션 접근 정보. 리본 도입 후의 구 command bar라는 범위를 명시.

## 추가 검수 수정

- 시트 이름 입력의 Enter/Escape 기본 동작이 격자 편집기로 전달되어 셀 내용을 덮어쓸 수 있던 오류를 차단하고 IME 조합 Enter를 구분했다.
- 보호된 셀에서도 다른 이름 저장/Excel 내보내기를 허용한다. 시트 수정 단축키는 읽기 전용·최종본을 우회하지 않는다.
- End 모드에서 Alt/F10 메뉴로 들어가면 끝 이동 대기를 해제한다. 구 메뉴 단계 안내에는 실제 다음 키만 표시하고 공통 팝업 자동 접근키를 덧붙이지 않는다.

## 최종 검증 (2026-10-05)

- 전체 Node 자동 검사: 2,098/2,098 통과.
- 리본 검사: 22/22 통과, 16개 탭의 364개 조작과 494개 등록 경로. 정적/호환 Alt 경로 검사 184개 실패 0, IME 검사 9/9 통과.
- 새 단축키 브라우저 검사: 소스·번들 × Chromium/Firefox/WebKit, 18개 시나리오씩 총 108/108·576개 검증 통과. 대표 배지 보완 후 증분 18/18·126개 검증 통과.
- 팝업 접근키: 3엔진 총 24/24·249개 검증, 기존 대화상자 22/22, 키보드 동등성 12/12·83개 검증, 셀 문맥 메뉴 21/21 통과.
- 추가 독립 리뷰 회귀: 소스 8/8·109개, 번들 3엔진 24/24·327개 검증 통과. 이름 확정 Enter의 셀 손상, 보호 시트 저장, End 상태 잔류, 메뉴 안내 키, 검색창 저장 문제를 재현 후 수정했다.
- 실제 XLSX B4 / XLSB B2 수정 전·후 읽기·계산·XLSX 저장 후 재열기 및 피벗 비교: 8/8 실행 통과. 합계 2,429,849개 셀·2,333,483개 수식·26개 피벗, 수식 불일치·재열기 문제·피벗 차이 모두 0. 원본 파일 및 검사용 복사본 SHA-256 변경 0. 브라우저 모든 기능의 Excel 동등성을 의미하지는 않는다.
- 실제 파일 수정 후 읽기 시간: B4 8.044초, B2 9.239초. XLSX roundtrip: B4 14.975초, B2 43.384초. 해당 로컬 환경에서 측정한 값이며 모바일 성능 보장은 아니다. 최초 검사 두 건은 브라우저 캐시 경로 미설정으로 실행 실패했고 환경 설정 후 재실행으로 통과했다.
- 초기 실패 로그도 D:/Codex/Temp 아래에 보존했다. 업무 파일·출력 통합 문서는 GitHub에 포함하지 않았다.

시트 이름·색 변경의 직접 마우스 경로에도 기존 문서/권한 확인을 적용했다. 읽기 전용·최종본 및 팝업을 연 뒤 문서가 전환된 경우 잘못된 문서에 적용하지 않는다. 제목/이름 입력창의 IME Enter와 확정 키도 셀 편집기로 전달되지 않도록 처리했다.

- 최종 배포 번들 전체 명령 실행: 397개, 오류 0. 직접 시트 이름/색 변경의 정상 동작·읽기 전용·최종본·팝업 후 권한 변경도 3엔진에서 검증했다.
