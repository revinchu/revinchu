// 모든 명령을 실제 브라우저에서 한 번씩 실행해 오류가 없는지 확인
//   (서버 실행 중) node tools/smoke.mjs        → "commands N, bad 0" 이 정상
// 환경 변수: PLAYWRIGHT_MODULE (playwright 경로), WIPOINT_URL (기본 http://127.0.0.1:5179/)
const { chromium } = await import(process.env.PLAYWRIGHT_MODULE || 'playwright');
const URL = process.env.WIPOINT_URL || 'http://127.0.0.1:5179/';
// 파일 고르기 · 인쇄 · 전체 화면 · 화면 캡처처럼 사람 손이 필요한 명령은 건너뜀
const SKIP = new Set(['open', 'save', 'saveAs', 'print', 'printNow', 'exportPng', 'insertPicture', 'changePicture', 'pickFillImage', 'screenshot', 'paste', 'pasteFiles', 'backstage', 'newPres', 'showFromStart', 'showFromCurrent', 'presenterView', 'rehearse', 'viewReading', 'insertIcons', 'share', 'saveAsPicture', 'exportImage', 'photoAlbum', 'smartLookup', 'translateSel', 'changePictureFromClipboard']);
const b = await chromium.launch();
const p = await b.newPage({ viewport: { width: 1400, height: 900 } });
const errs = [];
p.on('pageerror', (e) => errs.push(e.message));
await p.goto(URL);
await p.waitForTimeout(800);
await p.keyboard.press('Escape');
await p.evaluate(() => { window.wipoint.run('newBlank'); });
const names = await p.evaluate(() => Object.keys(window.wipoint.COMMANDS));
const bad = [];
for (const name of names) {
  if (SKIP.has(name)) continue;
  // 매번 같은 출발점: 슬라이드 3장(2번: 제목 및 내용), 본문 개체 틀 선택
  await p.evaluate(() => {
    const w = window.wipoint;
    w.S.formatPane = null;
    w.run('newBlank'); w.run('newSlide', 'titleContent'); w.run('newSlide', 'blank'); w.run('insertTable', 2, 2);
    w.goSlide(1);
    const o = w.S.pres.slides[1].objects[1];
    o.text.paras[0].runs = [{ t: '본문 글' }];
    w.S.sel = new Set([o.id]);
  });
  const before = errs.length;
  await p.evaluate((n) => { try { const r = window.wipoint.run(n, document.querySelector('.ribbon')); if (r?.catch) r.catch((e) => { window.__err = e.message; }); } catch (e) { window.__err = e.message; } }, name);
  await p.waitForTimeout(60);
  const e2 = await p.evaluate(() => { const e = window.__err; window.__err = null; return e; });
  // 열린 창 · 메뉴 · 쇼 닫기
  await p.keyboard.press('Escape');
  await p.evaluate(() => { for (const d of document.querySelectorAll('.dialog-backdrop, .menu, .show, .backstage')) d.remove(); });
  if (errs.length > before || e2) bad.push(`${name}: ${e2 ?? errs.slice(before).join(' / ')}`);
}
console.log(`commands ${names.length - SKIP.size}`);
console.log(`bad ${bad.length}`);
for (const x of bad) console.log('  ', x);
await b.close();
process.exit(bad.length ? 1 : 0);
