import { test } from 'node:test';
import assert from 'node:assert/strict';
import {
  newPresentation, newSlide, newShape, newTextBox, newTable, para, run, textBody, plainText, setPlainText, applyRunProps, formatRange, rangeRunProp,
  changeLayout, resetSlideLayout, duplicateSlide, cloneObjects, alignObjects, reorder, bbox, findAll, replaceAll, History, addAnim, animSteps, stepTimeline,
  resizePresentation, pruneMedia, validatePresentation, outline, LAYOUTS,
} from '../src/model.js';
import { resolveColor, applyMods, themePaletteRows, THEMES, colorRef } from '../src/themes.js';
import { shapePath, SHAPE_GALLERY, customPath, textRect } from '../src/shapes.js';
import { niceScale, chartSvg, chartToTsv, tsvToChart } from '../src/chart.js';
import { smartArt, SMART_KINDS } from '../src/smartart.js';
import { TEMPLATES, buildTemplate } from '../src/templates.js';
import { slideHtml, textHtml, numberLabel, themeDecor } from '../src/render.js';

test('새 문서: 제목 슬라이드 하나, 레이아웃 9종 모두 만들 수 있음', () => {
  const p = newPresentation();
  assert.equal(p.slides.length, 1);
  assert.equal(p.slides[0].layout, 'title');
  assert.deepEqual(p.slides[0].objects.map((o) => o.ph), ['ctrTitle', 'subTitle']);
  for (const [k] of LAYOUTS) assert.ok(newSlide(p, k));
  assert.equal(newSlide(p, 'comparison').objects.length, 5);
  assert.equal(newSlide(p, 'pictureCaption').objects.find((o) => o.ph === 'pic').type, 'image');
});

test('글: 일반 텍스트 · 범위 서식 · 공통 값', () => {
  const b = textBody([para('안녕하세요 세계'), para('둘째 줄')]);
  assert.equal(plainText(b), '안녕하세요 세계\n둘째 줄');
  formatRange(b, { p: 0, o: 3 }, { p: 1, o: 1 }, { b: true });
  assert.deepEqual(b.paras[0].runs, [{ t: '안녕하' }, { t: '세요 세계', b: true }]);
  assert.deepEqual(b.paras[1].runs, [{ t: '둘', b: true }, { t: '째 줄' }]);
  assert.equal(rangeRunProp(b, { p: 0, o: 4 }, { p: 0, o: 6 }, 'b'), true);
  assert.equal(rangeRunProp(b, { p: 0, o: 0 }, { p: 0, o: 6 }, 'b'), null);
  // 같은 서식 묶음은 합쳐짐
  formatRange(b, { p: 0, o: 0 }, { p: 0, o: 3 }, { b: true });
  assert.deepEqual(b.paras[0].runs, [{ t: '안녕하세요 세계', b: true }]);
  applyRunProps(b, { size: 20, b: undefined });
  assert.ok(b.paras.every((p) => p.runs.every((r) => r.size === 20 && !('b' in r))));
  setPlainText(b, '하나\n둘\n셋');
  assert.equal(b.paras.length, 3);
});

test('레이아웃 바꾸기: 내용은 옮기고 남는 내용 있는 개체 틀은 일반 개체로', () => {
  const p = newPresentation({ firstLayout: null });
  const s = newSlide(p, 'twoContent');
  s.objects[0].text.paras[0].runs = [run('제목')];
  s.objects[1].text.paras[0].runs = [run('왼쪽')];
  s.objects[2].text.paras[0].runs = [run('오른쪽')];
  changeLayout(p, s, 'titleContent');
  assert.equal(s.layout, 'titleContent');
  assert.equal(plainText(s.objects[0].text), '제목');
  assert.equal(plainText(s.objects[1].text), '왼쪽');
  assert.ok(s.objects.some((o) => o.phOrphan && plainText(o.text) === '오른쪽'));
  s.objects[0].x = 500;
  resetSlideLayout(p, s);
  assert.equal(s.objects[0].x, 88);
});

test('복제 · 그룹 복사는 새 id, 애니메이션 대상도 따라감', () => {
  const p = newPresentation();
  const s = p.slides[0];
  const a = newShape('rect', { x: 0, y: 0, w: 10, h: 10 });
  const b = newShape('ellipse', { x: 20, y: 0, w: 10, h: 10 });
  a.grp = b.grp = 'g1';
  s.objects.push(a, b);
  addAnim(s, a.id, 'entr', 'fade');
  const d = duplicateSlide(s);
  assert.notEqual(d.id, s.id);
  assert.ok(d.objects.every((o) => !s.objects.some((x) => x.id === o.id)));
  assert.equal(d.anims[0].obj, d.objects[2].id);
  assert.notEqual(d.objects[2].grp, 'g1');
  const c = cloneObjects([a, b], 5, 5);
  assert.equal(c[0].x, 5);
  assert.equal(c[0].grp, c[1].grp);
  assert.notEqual(c[0].grp, 'g1');
});

test('맞춤 · 배분 · 순서', () => {
  const objs = [0, 1, 2].map((i) => newShape('rect', { x: i * 100 + (i === 2 ? 50 : 0), y: i * 10, w: 50, h: 50 }));
  alignObjects(objs, 't', { w: 1280, h: 720 });
  assert.ok(objs.every((o) => o.y === 0));
  alignObjects(objs, 'dh', { w: 1280, h: 720 });
  assert.equal(objs[1].x, 125);
  assert.deepEqual(bbox(objs), { x: 0, y: 0, w: 300, h: 50 });
  const s = { objects: objs.slice() };
  reorder(s, [objs[0].id], 'front');
  assert.equal(s.objects[2], objs[0]);
  reorder(s, [objs[0].id], 'backward');
  assert.equal(s.objects[1], objs[0]);
});

test('찾기 · 모두 바꾸기 (표 칸과 메모 포함)', () => {
  const p = newPresentation();
  p.slides[0].objects[0].text.paras[0].runs = [run('매출 보고', { b: true }), run(' 매출')];
  const t = newTable(2, 2, { x: 0, y: 0, w: 200, h: 80 });
  t.rows[1].cells[1].text.paras[0].runs = [run('매출액')];
  p.slides[0].objects.push(t);
  p.slides[0].notes = '매출 메모';
  assert.equal(findAll(p, '매출').length, 4);
  assert.equal(findAll(p, '매출', { wholeWord: true }).length, 3);
  assert.equal(replaceAll(p, '매출', '수익'), 4);
  assert.equal(plainText(p.slides[0].objects[0].text), '수익 보고 수익');
  assert.equal(p.slides[0].objects[0].text.paras[0].runs[0].b, true);
  assert.equal(p.slides[0].notes, '수익 메모');
});

test('실행 취소: 연속 변경 묶기, 미디어는 유지', () => {
  const p = newPresentation();
  const h = new History();
  p.media.m1 = 'data:image/png;base64,AA==';
  h.record(p);
  p.slides.push(newSlide(p));
  h.record(p, 'k');
  p.slides[1].notes = 'a';
  h.record(p, 'k');
  p.slides[1].notes = 'ab';
  assert.equal(h.undo.length, 2);
  h.doUndo(p);
  assert.equal(p.slides.length, 2);
  assert.equal(p.slides[1].notes, '');
  h.doUndo(p);
  assert.equal(p.slides.length, 1);
  assert.equal(p.media.m1, 'data:image/png;base64,AA==');
  h.doRedo(p);
  assert.equal(p.slides.length, 2);
});

test('애니메이션 단계 · 시간표', () => {
  const p = newPresentation();
  const s = p.slides[0];
  const [a, b] = s.objects;
  addAnim(s, a.id, 'entr', 'fade', { dur: 1 });
  addAnim(s, b.id, 'entr', 'fly', { start: 'after', dur: 0.5 });
  addAnim(s, b.id, 'emph', 'pulse', { start: 'with', delay: 0.25 });
  addAnim(s, a.id, 'exit', 'fade');
  const steps = animSteps(s);
  assert.equal(steps.length, 2);
  assert.deepEqual(stepTimeline(steps[0]).map((x) => [x.at, x.end]), [[0, 1], [1, 1.5], [1.25, 1.75]]);
});

test('슬라이드 크기 바꾸기 (4:3 맞춤)', () => {
  const p = newPresentation();
  const t = p.slides[0].objects[0];
  t.text.paras[0].runs = [run('제목', { size: 40 })];
  resizePresentation(p, 960, 720, 'fit');
  assert.deepEqual(p.size, { w: 960, h: 720 });
  assert.ok(Math.abs(t.w - 960 * 0.75) < 1);
  assert.equal(t.text.paras[0].runs[0].size, 30);
});

test('미디어 정리 · 문서 검사 · 개요', () => {
  const p = newPresentation();
  p.media = { a: 'x', b: 'y' };
  p.slides[0].objects.push({ id: 'i', type: 'image', media: 'a', x: 0, y: 0, w: 1, h: 1 });
  pruneMedia(p);
  assert.deepEqual(Object.keys(p.media), ['a']);
  assert.throws(() => validatePresentation({}), /WIPOINT/);
  p.slides[0].objects[0].text.paras[0].runs = [run('표지')];
  assert.equal(outline(p)[0].title, '표지');
});

test('테마 색: 테마 슬롯 · 밝기 변형 · PowerPoint 선형 tint', () => {
  const t = THEMES[1];
  assert.equal(resolveColor(t, '@accent1'), '#4472C4');
  assert.equal(resolveColor(t, '@tx1'), '#000000');
  assert.equal(resolveColor({ colors: { dk1: '#000000' } }, '@dk1:t75'), '#898989');
  assert.equal(resolveColor(t, '#FF0000:a50'), 'rgba(255,0,0,0.5)');
  assert.equal(applyMods('#4472C4', [['lumMod', 0.2], ['lumOff', 0.8]]).hex, '#DAE3F3');
  assert.equal(colorRef('accent2', [['lumMod', 0.75]]), '@accent2:lm75');
  assert.equal(themePaletteRows().length, 6);
});

test('도형 경로: 갤러리의 모든 도형이 경로를 만듦', () => {
  for (const [, list] of SHAPE_GALLERY) for (const [k] of list) {
    const d = shapePath(k, 100, 60);
    assert.match(d, /^M/, k);
    assert.ok(!d.includes('NaN'), k);
  }
  assert.equal(customPath([{ w: 10, h: 10, cmds: [['M', 0, 0], ['L', 10, 10], ['Z']] }], 100, 50), 'M0 0 L100 50 Z');
  assert.deepEqual(textRect('rect', 10, 10), [0, 0, 10, 10]);
});

test('차트: 눈금 · SVG · 표 데이터 왕복', () => {
  assert.deepEqual(niceScale(0, 93), { lo: 0, hi: 100, step: 20 });
  const ch = { kind: 'col', title: 'T', cats: ['a', 'b'], series: [{ name: 's', vals: [1, 2] }], legend: 'b', labels: true };
  for (const kind of ['col', 'bar', 'line', 'area', 'pie', 'doughnut', 'scatter']) assert.match(chartSvg({ ...ch, kind }, 400, 300, THEMES[0]), /^<svg/);
  const tsv = chartToTsv(ch);
  assert.equal(tsv, '\ts\na\t1\nb\t2');
  const back = tsvToChart(ch, '\t매출\t비용\n1월\t1,000\t5\n2월\t20\tx');
  assert.deepEqual(back.series.map((s) => s.vals), [[1000, 20], [5, 0]]);
});

test('SmartArt 9종은 한 그룹의 도형들', () => {
  for (const [k] of SMART_KINDS) {
    const objs = smartArt(k, ['가', '나', '다'], { x: 0, y: 0, w: 800, h: 400 });
    assert.ok(objs.length >= 3, k);
    assert.equal(new Set(objs.map((o) => o.grp)).size, 1, k);
    for (const o of objs) assert.ok(Number.isFinite(o.x + o.y + o.w + o.h), k);
  }
});

test('서식 파일은 모두 만들어지고 그려짐', () => {
  for (const t of TEMPLATES) {
    const p = buildTemplate(t.id);
    assert.ok(p.slides.length >= 1, t.id);
    for (const [i, s] of p.slides.entries()) assert.match(slideHtml(p, s, { index: i }), /^<div class="sl/);
  }
});

test('렌더링: 글머리 번호 · 빈 개체 틀 안내 · 장식', () => {
  assert.equal(numberLabel('arabicPeriod', 3), '3.');
  assert.equal(numberLabel('romanUcPeriod', 4), 'IV.');
  assert.equal(numberLabel('circleNumDbPlain', 2), '②');
  assert.equal(numberLabel('ganada', 2), '나.');
  const p = newPresentation();
  assert.match(slideHtml(p, p.slides[0], { prompt: true }), /제목을 입력하십시오/);
  assert.doesNotMatch(slideHtml(p, p.slides[0], {}), /제목을 입력하십시오/);
  const o = newTextBox({ x: 0, y: 0, w: 100, h: 40 }, '<b>');
  assert.match(textHtml(THEMES[0], o, o.text), /&lt;b&gt;/);
  assert.equal(themeDecor({ ...THEMES[2] }, 'blank', { w: 1280, h: 720 }).length, 0);
  assert.ok(themeDecor({ ...THEMES[2] }, 'title', { w: 1280, h: 720 }).length > 0);
});

test('배경 제거: 가장자리 배경색은 투명, 가운데 피사체는 남김', async () => {
  const { backgroundMask, applyMask } = await import('../src/bgremove.js');
  const w = 40; const h = 30;
  const data = new Uint8ClampedArray(w * h * 4);
  for (let y = 0; y < h; y++) for (let x = 0; x < w; x++) {
    const i = (y * w + x) * 4;
    const inside = x >= 12 && x < 28 && y >= 8 && y < 22;
    data.set(inside ? [200, 30, 30, 255] : [250, 250, 248, 255], i);
  }
  // 피사체 안에 배경과 같은 색 구멍 (가장자리와 이어지지 않으므로 남아야 함)
  data.set([250, 250, 248, 255], (15 * w + 20) * 4);
  const a = backgroundMask(data, w, h);
  assert.equal(a[0], 0);
  assert.equal(a[15 * w + 15], 255);
  assert.equal(a[15 * w + 20], 255);
  // [제거할 영역 표시] 로 구멍도 지움
  const b = backgroundMask(data, w, h, { remove: [[20, 15]] });
  assert.equal(b[15 * w + 20], 0);
  // [보관할 영역 표시] 는 배경색이어도 남김
  const c = backgroundMask(data, w, h, { keep: [[2, 2]], brush: 1 });
  assert.equal(c[2 * w + 2], 255);
  assert.equal(applyMask(data, a)[3], 0);
});

test('개요 보기 편집: 줄 목록 ⇄ 슬라이드', async () => {
  const { outlineRows, applyOutline, newPresentation, newSlide, slideTitle } = await import('../src/model.js');
  const p = newPresentation({});
  const s2 = newSlide(p, 'titleContent');
  p.slides.push(s2);
  s2.objects[0].text.paras[0].runs = [{ t: '둘째', b: true }];
  s2.objects[1].text.paras = [{ runs: [{ t: '가', color: '#ff0000' }] }, { lvl: 1, runs: [{ t: '나' }] }];
  const rows = outlineRows(p);
  assert.deepEqual(rows.map((r) => [r.kind, r.text, r.lvl]), [['title', '', undefined], ['title', '둘째', undefined], ['body', '가', 0], ['body', '나', 1]]);
  // '나' 를 고치고, 새 슬라이드 '셋째' + 본문 추가
  rows[3].text = '나나';
  rows.push({ kind: 'title', text: '셋째' }, { kind: 'body', lvl: 0, text: '하나' });
  applyOutline(p, rows);
  assert.equal(p.slides.length, 3);
  assert.equal(p.slides[1], s2);
  assert.equal(s2.objects[1].text.paras[0].runs[0].color, '#ff0000');
  assert.equal(s2.objects[1].text.paras[1].runs[0].t, '나나');
  assert.equal(s2.objects[1].text.paras[1].lvl, 1);
  assert.equal(slideTitle(p.slides[2]), '셋째');
  assert.equal(outlineRows(p).at(-1).text, '하나');
  // 둘째 슬라이드 제목 줄을 지우면 (본문은 앞 슬라이드로) 슬라이드가 빠짐
  const r2 = outlineRows(p).filter((r) => r.text !== '둘째');
  applyOutline(p, r2);
  assert.equal(p.slides.length, 2);
});

test('공동 편집: 슬라이드 단위 patch 만들기 · 적용', async () => {
  const { baseline, diffPres, applyPatch, fullState, applyFull } = await import('../src/collabcore.js');
  const { newPresentation, newSlide } = await import('../src/model.js');
  const a = newPresentation({});
  a.slides.push(newSlide(a, 'titleContent'));
  const b = JSON.parse(JSON.stringify(a));
  const ba = baseline(a); const bb = baseline(b);
  assert.equal(diffPres(ba, a), null);
  // A: 2번 슬라이드 제목 고침 + 새 슬라이드 + 그림
  a.slides[1].objects[0].text.paras[0].runs = [{ t: '함께 고친 제목' }];
  const s3 = newSlide(a, 'blank'); a.slides.push(s3);
  a.media = { m1: 'data:image/png;base64,AA' };
  const p1 = diffPres(ba, a);
  assert.deepEqual(Object.keys(p1.slides).sort(), [a.slides[1].id, s3.id].sort());
  assert.ok(p1.order && p1.media.m1);
  // B: 동시에 1번 슬라이드를 고친 상태에서 A 의 patch 받기 (B 의 변경은 남음)
  b.slides[0].hidden = true;
  applyPatch(b, p1, bb);
  assert.equal(b.slides.length, 3);
  assert.equal(b.slides[1].objects[0].text.paras[0].runs[0].t, '함께 고친 제목');
  assert.equal(b.slides[0].hidden, true);
  const p2 = diffPres(bb, b);
  assert.deepEqual(Object.keys(p2.slides), [b.slides[0].id]);
  assert.equal(p2.media, undefined);
  // 슬라이드 지우기 · 순서 바꾸기
  b.slides = [b.slides[2], b.slides[0]];
  const p3 = diffPres(bb, b);
  applyPatch(a, p3, ba);
  assert.deepEqual(a.slides.map((s) => s.id), b.slides.map((s) => s.id));
  // 새로 들어온 사람: 전체 상태
  const c = newPresentation({});
  applyFull(c, fullState(a));
  assert.deepEqual(c.slides.map((s) => s.id), a.slides.map((s) => s.id));
  assert.equal(c.media.m1, 'data:image/png;base64,AA');
});
