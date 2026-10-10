# 05. Codex 에 붙여 넣을 지시문

## A. 환경 확인
```
이 저장소의 wipoint/ 는 WIPOINT(PowerPoint 와 같은 웹 프레젠테이션)입니다.
wipoint/AGENTS.md, wipoint/CLAUDE.md, wipoint/docs/00~03 을 먼저 읽으세요.
cd wipoint && npm test (26개 통과) → npm start → playwright 로 node tools/smoke.mjs (bad 0) 를 확인하고 보고하세요.
규칙: npm 의존성 추가 금지, import/export 형식 제한, 모든 UI 글자는 한국어.
```

## B. 배포 (WIXEL 3 와 같은 Cloudflare 계정)
```
wipoint/cloudflare/README.md 대로 node build.mjs --cloud 후 wrangler deploy --dry-run, 이상 없으면 배포하세요.
배포 후 https://wipoint.<서브도메인>.workers.dev 에서 새 프레젠테이션 · pptx 열기/저장 · 슬라이드 쇼가 되는지 확인하세요.
```

## C. 실제 PowerPoint 파일 호환 (사용자 파일이 있을 때)
```
~/wipoint-samples/ 의 .pptx 마다 LibreOffice 로 PDF 를 만들고, WIPOINT 에서 같은 파일을 열어 슬라이드별 화면을 캡처해 비교하세요 (docs/03 §5).
차이가 큰 것부터 pptx.js 읽기를 고치고, 고칠 때마다 test/pptx.test.js 에 작은 재현 테스트를 추가하세요. 사용자 파일은 커밋하지 마세요.
```

## D. 다음 기능 (docs/02 P1)
```
docs/02_보완항목.md 의 '차트 내장 통합 문서'를 구현하세요: 차트를 pptx 로 저장할 때 ppt/embeddings/Microsoft_Excel_Worksheet{n}.xlsx 를 만들고 c:externalData 로 연결 (xlsx 는 직접 작성, 의존성 없이).
PowerPoint 에서 [데이터 편집]이 되는지 사용자에게 확인을 요청하세요.
```
