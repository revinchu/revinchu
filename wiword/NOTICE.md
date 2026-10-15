# NOTICE — 참고한 프로젝트

WIWORD 의 코드는 이 저장소에서 작성했습니다. 아래 프로젝트는 구조와 동작을 연구하는 데 참고했고, MIT 라이선스인 kordoc · rhwp 에서는 일부 규칙과 절차를 JavaScript 로 옮겨 썼습니다 (아래 라이선스 전문). 예제 문서는 저장소에 넣지 않았습니다.

| 프로젝트 | 라이선스 | 참고한 것 |
|---|---|---|
| [chrisryugj/kordoc](https://github.com/chrisryugj/kordoc) | MIT | HWP 5.0 / HWPX / PDF → 마크다운 변환 방식 (레코드 해석, 표 셀 병합, 각주 · 머리글 처리 순서). HWP 3.0 의 한컴 사적 글자 영역 매핑 · 표 칸 기하 → 격자 규칙 · 그리기 개체 글상자 탐색 · 배포용/암호 문서 복호 절차를 옮겨 씀 (`src/hwp3.js`, `src/hwp.js`, `src/hwpx.js`) |
| [edwardkim/rhwp](https://github.com/edwardkim/rhwp) | MIT | HWP 3.0 레코드 구조(문서 정보 · 글자/문단 모양 · 표 · 그림) 와 조합형 매핑 검증, 암호 방식(HWP 3 DES · HWP 5 AES CFB-1 · HWPX PBKDF2/AES-CBC) 참고. 예제 문서는 로컬 시험에만 사용 (저장소에 포함하지 않음) |
| [HariFatherKR/hwp-parser](https://github.com/HariFatherKR/hwp-parser) | 저장소 표기 따름 | HWP 레코드 구조 참고, 예제 .hwp 파일은 로컬 시험에만 사용 (저장소에 포함하지 않음) |
| [storytold/wordcraft](https://github.com/storytold/wordcraft) | 저장소 표기 따름 | Rust 워드 프로세서의 화면 구성 · 기능 목록 · 사용 흐름 연구 (코드 사용 없음) |
| [Apache POI test-data](https://github.com/apache/poi) | Apache-2.0 | .docx / .doc 호환성 시험 문서 모음 (로컬 시험에만 사용) |

형식 문서: ECMA-376 (Office Open XML), [MS-DOC], 한글과컴퓨터 「HWP 5.0 문서 형식」 · 「OWPML(HWPX)」 공개 문서, PDF 1.7 (ISO 32000-1).

"한글", "HWP" 는 한글과컴퓨터의, "Microsoft Word" 는 Microsoft 의 상표입니다. WIWORD 는 이들과 관계가 없습니다.

## MIT 라이선스 전문

두 프로젝트의 라이선스 문구는 같고 저작권 표시만 다릅니다.

- kordoc — Copyright (c) 2026 chrisryugj
- rhwp — Copyright (c) 2025-2026 Edward Kim

```
Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.
```
