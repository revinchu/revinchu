// 호환성 점검용 "모든 기능" 프레젠테이션 → pptx (PowerPoint 스키마 검사 · LibreOffice 렌더 비교용)
// 사용: node tools/compat-deck.mjs out.pptx
import { writeFileSync } from 'node:fs';
import { writePptx } from '../src/pptx.js';
import { newPresentation, newSlide, newShape, newTextBox, newTable, newChart, newImage, para, run, addAnim, MOTION_PATHS, ANIM_EFFECTS, TRANSITIONS, addCustomLayout, designForTheme } from '../src/model.js';
import { smartArt, SMART_KINDS } from '../src/smartart.js';
import { CHART_KINDS, CHART_SUBTYPES } from '../src/chart.js';
import { FX_PRESETS } from '../src/effects.js';
import { ART_EFFECTS } from '../src/arteffects.js';
import { SOUNDS } from '../src/sounds.js';
import { albumSlides } from '../src/review.js';
import { THEMES, cloneTheme } from '../src/themes.js';

const PNG = 'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==';
const SVG = `data:image/svg+xml;base64,${Buffer.from('<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 10 10"><circle cx="5" cy="5" r="4" fill="red"/></svg>').toString('base64')}`;
const WAV = 'data:audio/wav;base64,UklGRiQAAABXQVZFZm10IBAAAAABAAEAIlYAAESsAAACABAAZGF0YQAAAAA=';
const MP4 = 'data:video/mp4;base64,AAAAGGZ0eXBtcDQyAAAAAG1wNDJpc29t';

export function compatDeck() {
  const p = newPresentation({ theme: 'office' });
  p.media.png = PNG; p.media.svg = SVG; p.media.wav = WAV; p.media.mp4 = MP4;
  const add = (layout = 'blank') => { const s = newSlide(p, layout); p.slides.push(s); return s; };
  p.slides[0].objects[0].text.paras[0].runs = [run('호환성 점검', { b: true })];
  p.slides[0].objects[1].text.paras[0].runs = [run('모든 기능')];

  // 1. 글자 · 단락 서식
  const s1 = add('titleContent');
  s1.objects[0].text.paras[0].runs = [run('글자 서식')];
  s1.objects[1].text.paras = [
    para('', { bullet: { type: 'char', char: '■', color: '#FF0000', size: 1.2, font: 'Wingdings' } }),
    para('', { lvl: 1, bullet: { type: 'num', scheme: 'romanUcPeriod', start: 3 }, align: 'ctr', lineSpacing: 1.5, spcBef: 6, spcAft: 3 }),
    para('', { lvl: 2, bullet: { type: 'num', scheme: 'ganada' }, align: 'just', tabs: [{ pos: 100, algn: 'dec' }] }),
    para('', { align: 'r', lineSpacingPt: 30 }),
  ];
  s1.objects[1].text.paras[0].runs = [run('굵게', { b: true }), run(' 기울임', { i: true }), run(' 밑줄', { u: true }), run(' 취소선', { s: true }), run(' 위첨자', { base: 1 }), run(' 아래첨자', { base: -1 })];
  s1.objects[1].text.paras[1].runs = [run('색', { color: '@accent2:lm75' }), run(' 형광', { hl: '#FFFF00' }), run(' 간격', { spc: 3 }), run(' 대문자', { cap: 'all' }), run(' 작은 대문자', { cap: 'small' })];
  s1.objects[1].text.paras[2].runs = [run('윤곽선', { outline: { width: 1, color: '#0000FF' } }), run(' 그림자', { shadow: true }), run(' 글꼴', { font: '바탕', size: 20 }), run(' 링크', { link: 'https://example.com' })];
  s1.objects[1].text.paras[3].runs = [run('줄\v바꿈')];
  const tb = newTextBox({ x: 800, y: 500, w: 300, h: 100 }, '세로 글');
  tb.text.vert = 'eaVert'; tb.text.anchor = 'ctr'; tb.text.autofit = 'shrink'; tb.text.fontScale = 0.9; tb.text.lnSpcReduction = 0.1; tb.text.wrap = false; tb.text.insets = [0, 0, 0, 0];
  s1.objects.push(tb);
  const tb2 = newTextBox({ x: 100, y: 600, w: 300, h: 60 }, '자동 맞춤');
  tb2.text.autofit = 'resize';
  s1.objects.push(tb2);

  // 2. 도형 · 채우기 · 선 · 효과
  const s2 = add();
  const shapes = ['rect', 'roundRect', 'ellipse', 'triangle', 'rtTriangle', 'diamond', 'pentagon', 'hexagon', 'star5', 'rightArrow', 'chevron', 'cloud', 'heart', 'wedgeRectCallout', 'flowChartProcess', 'leftBrace', 'smileyFace', 'donut', 'blockArc', 'can'];
  shapes.forEach((k, i) => {
    const o = newShape(k, { x: 20 + (i % 10) * 125, y: 20 + Math.floor(i / 10) * 140, w: 110, h: 110 });
    o.rot = i * 7; o.flipH = i % 3 === 0; o.flipV = i % 5 === 0;
    o.fill = i % 4 === 0 ? { type: 'gradient', angle: 30, stops: [[0, '@accent1'], [0.5, '#FFFFFF'], [1, 'rgba(0,128,0,0.5)']] } : i % 4 === 1 ? { type: 'gradient', path: true, stops: [[0, '#FF0000'], [1, '#0000FF']] } : i % 4 === 2 ? { type: 'image', media: 'png', tile: i % 8 === 2 } : { type: 'solid', color: '@accent3:lm60:lo40' };
    o.line = { width: 1 + (i % 4), color: '@tx1', dash: ['solid', 'dash', 'sysDot', 'lgDashDot'][i % 4], head: i % 2 ? 'triangle' : 'none', tail: i % 3 ? 'oval' : 'none' };
    o.fx = FX_PRESETS[i % FX_PRESETS.length];
    if (i === 0) o.adj = { adj: 30000 };
    s2.objects.push(o);
  });
  const free = newShape('rect', { x: 20, y: 320, w: 200, h: 200 });
  free.path = [{ w: 100, h: 100, cmds: [['M', 0, 0], ['L', 100, 0], ['C', 100, 50, 50, 100, 0, 100], ['Q', 20, 50, 0, 0], ['Z']] }];
  s2.objects.push(free);
  const ln = newShape('line', { x: 300, y: 350, w: 300, h: 0 });
  ln.line = { width: 3, color: '#FF0000', dash: 'dash', head: 'stealth', tail: 'arrow' };
  ln.text = null;
  s2.objects.push(ln);
  const cx = newShape('straightConnector1', { x: 300, y: 400, w: 300, h: 100 });
  cx.text = null; cx.line = { width: 2, color: '@accent2' };
  s2.objects.push(cx);
  // 그룹
  const g1 = newShape('rect', { x: 700, y: 350, w: 100, h: 100 });
  const g2 = newShape('ellipse', { x: 820, y: 350, w: 100, h: 100 });
  g1.group = g2.group = 'grp1';
  s2.objects.push(g1, g2);

  // 3. 그림
  const s3 = add();
  const pics = [];
  for (let i = 0; i < 8; i++) {
    const im = newImage(i === 7 ? 'svg' : 'png', { x: 20 + i * 150, y: 20, w: 140, h: 140 });
    im.alt = `그림 ${i}`;
    if (i === 1) im.crop = { l: 0.1, t: 0.2, r: 0.1, b: 0 };
    if (i === 2) { im.alpha = 0.5; im.gray = true; }
    if (i === 3) { im.bright = 0.2; im.contrast = -0.2; }
    if (i === 4) { im.shape = 'ellipse'; im.line = { width: 3, color: '@accent1' }; }
    if (i === 5) { im.fx = FX_PRESETS[6]; im.link = 'https://example.com'; }
    if (i === 6) { im.art = { effect: ART_EFFECTS[1][0], amount: 50 }; im.artOrig = 'png'; im.decorative = true; }
    pics.push(im);
  }
  s3.objects.push(...pics);
  s3.objects.push({ id: 'vid1', type: 'media', kind: 'video', media: 'mp4', x: 20, y: 300, w: 320, h: 180, rot: 0, trim: { st: 0.5, end: 1 }, poster: 'png' });
  s3.objects.push({ id: 'aud1', type: 'media', kind: 'audio', media: 'wav', x: 400, y: 300, w: 48, h: 48, rot: 0 });
  s3.objects.push({ id: 'eq1', type: 'equation', x: 500, y: 400, w: 400, h: 100, rot: 0, latex: '\\frac{a}{b} + \\sqrt{x^2+y^2} = \\sum_{i=1}^{n} i', size: 28 });

  // 4. 표
  const s4 = add();
  const t = newTable(4, 4, { x: 50, y: 50, w: 800, h: 300 });
  t.rows[0].cells[0].text.paras[0].runs = [run('머리글')];
  t.rows[0].cells[0].span = 2; t.rows[0].cells[1].hMerge = true;
  t.rows[1].cells[0].rowSpan = 2; t.rows[2].cells[0].vMerge = true;
  t.rows[1].cells[1].fill = { type: 'solid', color: '#FFCC00' };
  t.rows[1].cells[2].borders = { l: { width: 2, color: '#FF0000' }, dd: { width: 1, color: '#000000' } };
  t.rows[3].cells[3].text.anchor = 'b';
  t.style = { ...t.style, firstRow: true, banded: true, lastRow: true, firstCol: true, accent: 'accent2' };
  s4.objects.push(t);

  // 5. 차트: 모든 종류 · 하위 형식
  for (const [kind] of CHART_KINDS) {
    const subs = CHART_SUBTYPES[kind]?.map((x) => x[0]) ?? [null];
    const s = add();
    subs.slice(0, 9).forEach((sub, i) => {
      const ch = newChart({ x: 10 + (i % 3) * 420, y: 10 + Math.floor(i / 3) * 235, w: 400, h: 225 });
      ch.chart.kind = kind;
      if (sub) ch.chart.subtype = sub;
      Object.assign(ch.chart, { title: `${kind} ${sub ?? ''}`, legend: 'r', labels: i % 2 ? { val: true, pos: 'ctr' } : null, dataTable: i === 2, grid: { major: true, minor: i === 3, cat: i === 4 }, axes: { x: { title: 'X' }, y: { title: 'Y', min: 0, fmt: '#,##0' } } });
      ch.chart.series[0].trend = i === 1 ? { type: 'linear', r2: true, eq: true } : null;
      ch.chart.series[0].err = i === 1 ? { type: 'percent', val: 5 } : null;
      ch.chart.series[0].pointColors = { 0: '#FF0000' };
      s.objects.push(ch);
    });
  }
  // 3차원 차트
  const s3d = add();
  ['clustered3d', 'stacked3d', 'percent3d', 'full3d'].forEach((sub, i) => {
    const ch = newChart({ x: 10 + (i % 2) * 620, y: 10 + Math.floor(i / 2) * 350, w: 600, h: 340 });
    ch.chart.subtype = sub; ch.chart.d3 = true; if (sub === 'full3d') ch.chart.variant3d = 'full';
    s3d.objects.push(ch);
  });
  for (const k of ['line', 'area', 'pie', 'bar']) {
    const ch = newChart({ x: 100, y: 100, w: 300, h: 200 }); ch.chart.kind = k; ch.chart.d3 = true; s3d.objects.push(ch);
  }
  // 콤보 · 보조 축
  const sc = add();
  const combo = newChart({ x: 50, y: 50, w: 600, h: 400 });
  Object.assign(combo.chart, { kind: 'combo' });
  combo.chart.series[1] = { ...combo.chart.series[1], type: 'line', axis: 'secondary' };
  sc.objects.push(combo);

  // 6. SmartArt 전부
  const ssa = add();
  SMART_KINDS.forEach(([k], i) => ssa.objects.push(...smartArt(k, ['가', '\t나', '다'], { x: 10 + (i % 8) * 158, y: 10 + Math.floor(i / 8) * 118, w: 150, h: 110 })));

  // 7. 애니메이션 전부 (효과 · 소리 · 텍스트 단위 · 트리거 · 단락별)
  for (const cls of ['entr', 'emph', 'exit', 'path']) {
    const s = add();
    ANIM_EFFECTS[cls].forEach(([eff], i) => {
      const o = newShape('rect', { x: 10 + (i % 10) * 125, y: 10 + Math.floor(i / 10) * 130, w: 110, h: 110 });
      o.text.paras[0].runs = [run(`${eff} 효과`)];
      s.objects.push(o);
      addAnim(s, o.id, cls, eff, { start: ['click', 'with', 'after'][i % 3], dir: ['b', 't', 'l', 'r'][i % 4], delay: i % 2 ? 0.25 : 0, ...(i === 1 ? { sound: SOUNDS[0][0] } : {}), ...(i === 2 ? { textUnit: 'letter', unitDelay: 20 } : {}), ...(i === 3 ? { byPara: true } : {}), ...(cls === 'path' ? { motion: MOTION_PATHS[eff] } : {}) });
    });
    const trig = newShape('rect', { x: 1100, y: 600, w: 100, h: 50 });
    s.objects.push(trig);
    addAnim(s, s.objects[0].id, 'entr', 'fade', { trigger: trig.id });
  }
  // 8. 전환 전부
  for (const [k] of TRANSITIONS) {
    const s = add('titleOnly');
    s.objects[0].text.paras[0].runs = [run(`전환 ${k}`)];
    s.transition = { type: k, dir: 'l', dur: k === 'fade' ? 1.3 : 0.75, advAfter: 2, advClick: true, ...(k === 'push' ? { sound: SOUNDS[2][0] } : {}) };
  }
  // 9. 배경 · 숨김 · 메모 · 메모(댓글) · 구역 · 확대/축소
  const s9 = add('section');
  s9.bg = { type: 'gradient', angle: 45, stops: [[0, '@bg1'], [1, '@accent1:lm20:lo80']] };
  s9.notes = '발표자 메모\n두 줄';
  s9.comments = [{ id: 'c1', author: '홍길동', text: '확인 바람', x: 10, y: 10, time: '2026-01-01T00:00:00Z', replies: [{ author: '김', text: '네' }] }];
  s9.section = '부록';
  s9.hidden = true;
  const s10 = add();
  s10.bg = { type: 'image', media: 'png' };
  s10.objects.push({ id: 'z1', type: 'zoom', x: 100, y: 100, w: 320, h: 180, rot: 0, zoom: { slide: p.slides[1].id, ret: true } });
  const s11 = add();
  s11.bg = { type: 'solid', color: '#FFEEDD' };
  s11.hideDecor = true;
  p.footer = { slideNum: true, date: true, text: '바닥글', hideOnTitle: true };
  p.customShows = [{ name: '짧게', slides: [p.slides[0].id, p.slides[2].id] }];
  // 사용자 지정 레이아웃
  addCustomLayout(p, '내 레이아웃');
  // 사진 앨범
  const albumPres = albumSlides(p, [{ media: 'png', caption: '사진 1' }, { media: 'png', caption: '사진 2' }], { layout: 'fit2', frame: 'roundRect', captions: true });
  if (Array.isArray(albumPres)) p.slides.push(...albumPres);
  // 여러 마스터: 다른 테마 디자인 1개
  const other = Object.keys(THEMES).find((k) => k !== 'office');
  const d = designForTheme(p, cloneTheme(THEMES[other]));
  const sd = add('titleContent');
  sd.design = d.id;
  sd.objects[0].text.paras[0].runs = [run('두 번째 마스터')];
  return p;
}

if (process.argv[2]) {
  writeFileSync(process.argv[2], writePptx(compatDeck()));
  console.log('wrote', process.argv[2]);
}
