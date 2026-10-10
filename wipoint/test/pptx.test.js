import { test } from 'node:test';
import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { dirname, join } from 'node:path';
import { writePptx, readPptx, pptxEntries, chartXml } from '../src/pptx.js';
import { unzip, textOf, zip } from '../src/zip.js';
import { parseXml } from '../src/xml.js';
import { newPresentation, newSlide, newShape, newTextBox, newTable, newChart, newImage, para, run, plainText, addAnim } from '../src/model.js';
import { smartArt } from '../src/smartart.js';
import { TEMPLATES, buildTemplate } from '../src/templates.js';

const here = dirname(fileURLToPath(import.meta.url));
const PNG = 'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==';

function sample() {
  const p = newPresentation({ theme: 'mint' });
  p.slides[0].objects[0].text.paras[0].runs = [run('표지 제목', { b: true })];
  p.slides[0].objects[1].text.paras[0].runs = [run('부제목')];
  const s = newSlide(p, 'titleContent');
  p.slides.push(s);
  s.objects[0].text.paras[0].runs = [run('핵심 성과')];
  s.objects[1].text.paras = [para('노출 120만', { bullet: { type: 'char', char: '•' } }), para('클릭률', { lvl: 1, bullet: { type: 'char', char: '•' } }), para('ROAS', { bullet: { type: 'num', scheme: 'arabicPeriod', start: 1 } })];
  const sh = newShape('roundRect', { x: 900, y: 450, w: 260, h: 120 });
  sh.text.paras[0].runs = [run('강조', { color: '#FF0000', size: 24, i: true })];
  sh.shadow = true;
  s.objects.push(sh);
  const tb = newTextBox({ x: 100, y: 600, w: 300, h: 40 }, '링크');
  tb.text.paras[0].runs[0].link = 'https://example.com';
  tb.link = '#slide1';
  s.objects.push(tb);
  s.notes = '메모 1줄\n2줄';
  s.transition = { type: 'push', dir: 'l', dur: 0.5, advAfter: 3 };
  s.section = '본문';
  addAnim(s, sh.id, 'entr', 'fly', { dir: 'l' });
  addAnim(s, s.objects[1].id, 'entr', 'fade', { start: 'after' });
  addAnim(s, sh.id, 'exit', 'zoom');
  const s2 = newSlide(p, 'blank');
  p.slides.push(s2);
  const t = newTable(3, 3, { x: 100, y: 100, w: 600, h: 150 });
  t.rows[0].cells[0].text.paras[0].runs = [run('항목')];
  t.rows[1].cells[1].fill = { type: 'solid', color: '@accent2' };
  t.style.accent = 'accent3';
  s2.objects.push(t);
  s2.objects.push(newChart({ x: 720, y: 100, w: 480, h: 360 }));
  p.media.m1 = PNG;
  const im = newImage('m1', { x: 100, y: 400, w: 200, h: 200 });
  im.crop = { l: 0.1, t: 0, r: 0.1, b: 0 };
  im.alt = '설명';
  s2.objects.push(im);
  s2.objects.push(...smartArt('chevron', ['가', '나'], { x: 400, y: 450, w: 400, h: 100 }));
  s2.hidden = true;
  s2.bg = { type: 'gradient', angle: 45, stops: [[0, '@bg1'], [1, '@accent1:lm20:lo80']] };
  p.footer = { slideNum: true, date: false, text: '회사명', hideOnTitle: true };
  return p;
}

test('pptx 쓰기: 필수 파트 · 관계 · 콘텐츠 형식 · 모든 XML 이 올바름', () => {
  const files = unzip(writePptx(sample()));
  for (const part of ['[Content_Types].xml', '_rels/.rels', 'ppt/presentation.xml', 'ppt/_rels/presentation.xml.rels', 'ppt/slideMasters/slideMaster1.xml', 'ppt/theme/theme1.xml', 'ppt/slides/slide1.xml', 'ppt/notesMasters/notesMaster1.xml', 'ppt/notesSlides/notesSlide2.xml', 'ppt/charts/chart1.xml', 'ppt/media/image1.png', 'docProps/core.xml']) assert.ok(files[part], part);
  const ct = textOf(files['[Content_Types].xml']);
  for (const name of Object.keys(files)) {
    if (!name.endsWith('.xml') && !name.endsWith('.rels')) continue;
    const x = textOf(files[name]);
    assert.ok(parseXml(x).name, name);
    // 태그 균형 (아주 단순한 검사)
    const opens = (x.match(/<[A-Za-z][^>]*[^/]>/g) ?? []).filter((t) => !t.startsWith('<?')).length;
    const closes = (x.match(/<\/[^>]+>/g) ?? []).length;
    assert.equal(opens, closes, `${name} 태그 수`);
    if (name.startsWith('ppt/') && !name.includes('_rels')) assert.ok(ct.includes(`/${name}"`), `${name} 콘텐츠 형식`);
  }
  // 관계 대상이 모두 있음
  for (const name of Object.keys(files).filter((n) => n.endsWith('.rels'))) {
    const base = name.replace('_rels/', '').replace(/\.rels$/, '');
    for (const m of textOf(files[name]).matchAll(/Target="([^"]+)"(?: TargetMode="External")?/g)) {
      if (m[0].includes('External')) continue;
      const dir = base.split('/').slice(0, -1);
      for (const seg of m[1].split('/')) { if (seg === '..') dir.pop(); else dir.push(seg); }
      assert.ok(files[dir.join('/')] || files[m[1]], `${name} → ${m[1]}`);
    }
  }
});

test('pptx 왕복: 글 · 서식 · 표 · 차트 · 그림 · 그룹 · 메모 · 전환 · 애니메이션 · 구역 · 바닥글', () => {
  const p = sample();
  const { pres: q, warnings } = readPptx(writePptx(p));
  assert.deepEqual(warnings, []);
  assert.equal(q.slides.length, 3);
  assert.deepEqual(q.size, p.size);
  assert.equal(q.theme.colors.accent1, p.theme.colors.accent1);
  assert.equal(q.theme.accentBand, true);
  const [s0, s1, s2] = q.slides;
  assert.equal(plainText(s0.objects[0].text), '표지 제목');
  assert.equal(s0.objects[0].text.paras[0].runs[0].b, true);
  assert.equal(s0.objects[0].ph, 'ctrTitle');
  assert.equal(s1.layout, 'titleContent');
  assert.deepEqual(s1.objects[1].text.paras.map((x) => [x.lvl, x.bullet?.type, x.bullet?.scheme ?? x.bullet?.char]), [[0, 'char', '•'], [1, 'char', '•'], [0, 'num', 'arabicPeriod']]);
  const sh = s1.objects[2];
  assert.equal(sh.shape, 'roundRect');
  assert.equal(sh.shadow, true);
  assert.deepEqual([sh.text.paras[0].runs[0].color, sh.text.paras[0].runs[0].size, sh.text.paras[0].runs[0].i], ['#FF0000', 24, true]);
  assert.equal(s1.objects[3].text.paras[0].runs[0].link, 'https://example.com');
  assert.equal(s1.objects[3].link, '#slide1');
  assert.equal(s1.notes, '메모 1줄\n2줄');
  assert.deepEqual(s1.transition, { type: 'push', dir: 'l', dur: 0.5, advAfter: 3 });
  assert.equal(s1.section, '본문');
  assert.deepEqual(s1.anims.map((a) => [a.cls, a.effect, a.start, a.obj === sh.id]), [['entr', 'fly', 'click', true], ['entr', 'fade', 'after', false], ['exit', 'zoom', 'click', true]]);
  assert.equal(s2.hidden, true);
  assert.equal(s2.bg.type, 'gradient');
  const t = s2.objects.find((o) => o.type === 'table');
  assert.equal(t.style.accent, 'accent3');
  assert.equal(t.rows[1].cells[1].fill.color, '@accent2');
  assert.equal(t.rows[0].cells[1].fill, undefined);
  assert.equal(plainText(t.rows[0].cells[0].text), '항목');
  assert.equal(s2.objects.find((o) => o.type === 'chart').chart.series.length, 3);
  const im = s2.objects.find((o) => o.type === 'image');
  assert.equal(im.alt, '설명');
  assert.ok(Math.abs(im.crop.l - 0.1) < 1e-6);
  assert.ok(q.media[im.media].startsWith('data:image/png;base64,'));
  const g = s2.objects.filter((o) => o.grp);
  assert.equal(g.length, 2);
  assert.equal(new Set(g.map((o) => o.grp)).size, 1);
  assert.deepEqual(q.footer, p.footer);
  // 바닥글 개체 틀이 일반 글 상자로 두 번 생기지 않음
  assert.ok(!s1.objects.some((o) => plainText(o.text) === '회사명'));
});

test('서식 파일 9종 모두 pptx 로 쓰고 다시 읽음', () => {
  for (const t of TEMPLATES) {
    const p = buildTemplate(t.id);
    const { pres: q } = readPptx(writePptx(p));
    assert.equal(q.slides.length, p.slides.length, t.id);
    assert.deepEqual(q.slides.map((s) => s.objects.length), p.slides.map((s) => s.objects.length), t.id);
  }
});

test('차트 파트: 막대/꺾은선/원형 모두 값이 들어감', () => {
  for (const kind of ['col', 'bar', 'line', 'area', 'pie', 'doughnut', 'scatter']) {
    const x = chartXml({ kind, title: 'T', cats: ['a', 'b'], series: [{ name: 's', vals: [1, 2] }], legend: 'r', labels: true });
    assert.ok(parseXml(x), kind);
    assert.match(x, /<c:v>2<\/c:v>/, kind);
  }
});

test('python-pptx (PowerPoint 기본 서식) 파일 읽기: 마스터 · 레이아웃 상속', () => {
  const { pres, warnings } = readPptx(new Uint8Array(readFileSync(join(here, 'fixtures/python-pptx.pptx'))));
  assert.deepEqual(warnings, []);
  assert.deepEqual(pres.size, { w: 960, h: 720 });
  const [s0, s1, s2] = pres.slides;
  assert.equal(plainText(s0.objects[0].text), 'python-pptx 제목');
  assert.equal(s0.objects[0].text.paras[0].runs[0].size, 44); // 마스터 titleStyle
  assert.equal(s0.objects[0].text.paras[0].align, 'ctr'); // 레이아웃 제목 맞춤
  assert.ok(s0.objects[0].x > 50 && s0.objects[0].w > 500); // 위치는 레이아웃에서
  assert.equal(s0.objects[1].text.paras[0].runs[0].color, '@dk1:t75');
  assert.deepEqual(s1.objects[1].text.paras.map((p) => [p.lvl, p.bullet?.char]), [[0, '•'], [1, '–'], [0, '•']]);
  assert.equal(s1.notes, '발표 메모 python');
  const shapes = s2.objects.filter((o) => o.type === 'shape' && !o.ph);
  assert.equal(shapes[0].shape, 'roundRect');
  // 도형 스타일 fillRef idx=3 → 테마의 세 번째 채우기(그라데이션), 색은 accent1
  assert.equal(shapes[0].fill.type, 'gradient');
  assert.ok(shapes[0].fill.stops.every(([, c]) => c.startsWith('@accent1')));
  assert.equal(shapes[0].line.color, '@accent1');
  assert.equal(shapes[1].fill.color, '#C00000');
  assert.equal(s2.objects.find((o) => o.type === 'table').rows[1].cells[1].text.paras[0].runs[0].t, '42');
  const ch = s2.objects.find((o) => o.type === 'chart').chart;
  assert.deepEqual([ch.kind, ch.cats, ch.series[0].vals, ch.title], ['col', ['가', '나', '다'], [3, 5, 2], '매출']);
  assert.equal(new Set(s2.objects.filter((o) => o.grp).map((o) => o.grp)).size, 1);
});

test('LibreOffice 가 다시 저장한 pptx 읽기', () => {
  const { pres, warnings } = readPptx(new Uint8Array(readFileSync(join(here, 'fixtures/libreoffice.pptx'))));
  assert.deepEqual(warnings, []);
  assert.equal(pres.slides.length, 3);
  assert.equal(plainText(pres.slides[0].objects[0].text), '위포인트 발표 자료');
  assert.ok(pres.slides[2].objects.some((o) => o.type === 'table'));
  assert.ok(pres.slides[2].objects.some((o) => o.type === 'chart'));
  assert.ok(pres.slides[2].objects.some((o) => o.type === 'image'));
});

test('암호 · 예전 형식 · ZIP 이 아닌 파일은 한국어 오류', () => {
  assert.throws(() => readPptx(new TextEncoder().encode('hello world, not a zip file at all')), /ZIP/);
});

test('파트 목록은 [Content_Types].xml 이 맨 앞', () => {
  assert.equal(Object.keys(pptxEntries(newPresentation()))[0], '[Content_Types].xml');
});

// embedded-font.pptx: DejaVu Serif Bold 부분 집합을 MTX 압축 EOT(.fntdata, XOR 포함)로 넣은 파일 (make-embedded-font-pptx.py)
test('포함된 글꼴: EOT + MTX(LZCOMP · CTF) 풀기, 다시 저장해도 유지', async () => {
  const { decodeEmbeddedFont, TRIPLETS } = await import('../src/fonts.js');
  const { pres, warnings } = readPptx(readFileSync(join(here, 'fixtures', 'embedded-font.pptx')));
  assert.deepEqual(warnings, []);
  assert.equal(pres.fonts.length, 1);
  assert.equal(pres.fonts[0].typeface, 'WP Embedded Serif');
  assert.equal(pres.fonts[0].style, 'regular');
  const bin = Buffer.from(pres.fonts[0].data.split(',')[1], 'base64');
  const { data, compressed } = decodeEmbeddedFont(new Uint8Array(bin));
  assert.equal(compressed, true);
  assert.deepEqual([...data.subarray(0, 4)], [0, 1, 0, 0]);
  // 표 목록에 glyf · loca 가 다시 생기고 maxp 의 글리프 수(73)만큼 loca 항목
  const tags = [];
  const n = (data[4] << 8) | data[5];
  let loca = 0;
  for (let i = 0; i < n; i++) { const o = 12 + i * 16; const tag = String.fromCharCode(...data.subarray(o, o + 4)); tags.push(tag); if (tag === 'loca') loca = (data[o + 12] << 24 | data[o + 13] << 16 | data[o + 14] << 8 | data[o + 15]); }
  assert.ok(tags.includes('glyf') && tags.includes('loca') && tags.includes('cmap'));
  assert.ok(loca === 74 * 2 || loca === 74 * 4);
  assert.equal(TRIPLETS.length, 128);
  // 다시 저장 → 같은 글꼴 데이터, embeddedFontLst
  const files = unzip(writePptx(pres));
  assert.match(textOf(files['ppt/presentation.xml']), /<p:embeddedFontLst><p:embeddedFont><p:font typeface="WP Embedded Serif"/);
  assert.match(textOf(files['[Content_Types].xml']), /Extension="fntdata"/);
  assert.deepEqual([...files['ppt/fonts/font1.fntdata']], [...bin]);
  assert.equal(readPptx(writePptx(pres)).pres.fonts.length, 1);
});

test('비디오 · 오디오: p:pic + videoFile/audioFile + p14:media 로 저장하고 다시 읽기', () => {
  const p = newPresentation();
  const s = p.slides[0];
  p.media.v1 = 'data:video/mp4;base64,AAAAIGZ0eXBpc29tAAACAGlzb21pc28y';
  p.media.a1 = 'data:audio/mpeg;base64,SUQzBAAAAAAAI1RTU0U=';
  p.media.p1 = PNG;
  s.objects.push({ id: 'vid', type: 'media', kind: 'video', media: 'v1', poster: 'p1', x: 100, y: 100, w: 640, h: 360, rot: 0, trim: { st: 1.5, end: 10 } });
  s.objects.push({ id: 'aud', type: 'media', kind: 'audio', media: 'a1', x: 900, y: 500, w: 64, h: 64, rot: 0 });
  const bytes = writePptx(p);
  const files = unzip(bytes);
  const xml = textOf(files['ppt/slides/slide1.xml']);
  assert.match(xml, /<a:videoFile r:link="rId\d+"\/>/);
  assert.match(xml, /<a:audioFile r:link="rId\d+"\/>/);
  assert.match(xml, /<p14:media [^>]*r:embed="rId\d+"><p14:trim st="1500" end="10000"\/>/);
  assert.ok(files['ppt/media/media1.mp4'] || Object.keys(files).some((k) => /media\d+\.mp4$/.test(k)));
  assert.match(textOf(files['[Content_Types].xml']), /Extension="mp4" ContentType="video\/mp4"/);
  assert.match(textOf(files['ppt/slides/_rels/slide1.xml.rels']), /relationships\/video"/);
  const back = readPptx(bytes).pres;
  const objs = back.slides[0].objects.filter((o) => o.type === 'media');
  assert.equal(objs.length, 2);
  const v = objs.find((o) => o.kind === 'video');
  assert.equal(back.media[v.media], p.media.v1);
  assert.ok(v.poster && back.media[v.poster]);
  assert.deepEqual(v.trim, { st: 1.5, end: 10 });
  assert.equal(back.media[objs.find((o) => o.kind === 'audio').media], p.media.a1);
});

test('차트: 데이터 통합 문서(embeddings/*.xlsx) + 셀 참조 캐시, 다시 읽기', async () => {
  const p = newPresentation();
  const ch = { kind: 'col', cats: ['1월', '2월', '3월'], series: [{ name: '클릭', vals: [10, 25, 18] }, { name: '전환', vals: [1, 2, 3] }], legend: 'b' };
  p.slides[0].objects.push(newChart({ x: 100, y: 100, w: 600, h: 400 }, ch));
  const files = unzip(writePptx(p));
  const cx = textOf(files['ppt/charts/chart1.xml']);
  assert.match(cx, /<c:externalData r:id="rId1">/);
  assert.match(cx, /<c:f>Sheet1!\$B\$2:\$B\$4<\/c:f>/);
  assert.match(cx, /<c:f>Sheet1!\$A\$2:\$A\$4<\/c:f>/);
  assert.match(textOf(files['ppt/charts/_rels/chart1.xml.rels']), /embeddings\/Microsoft_Excel_Worksheet1\.xlsx/);
  const wb = unzip(files['ppt/embeddings/Microsoft_Excel_Worksheet1.xlsx']);
  const sheet = textOf(wb['xl/worksheets/sheet1.xml']);
  assert.match(sheet, /<c r="B1" t="inlineStr"><is><t xml:space="preserve">클릭<\/t>/);
  assert.match(sheet, /<c r="C4"><v>3<\/v><\/c>/);
  assert.match(sheet, /<c r="A2" t="inlineStr"><is><t xml:space="preserve">1월/);
  const back = readPptx(writePptx(p)).pres.slides[0].objects.find((o) => o.type === 'chart');
  assert.deepEqual(back.chart.cats, ['1월', '2월', '3월']);
  assert.deepEqual(back.chart.series[1].vals, [1, 2, 3]);
});

test('모핑 전환 · 단락별 애니메이션 · 이동 경로 저장/읽기', () => {
  const p = sample();
  const s = p.slides[1];
  const t = s.objects.find((o) => o.text && o.text.paras.length > 1) ?? s.objects[1];
  t.text.paras = [para('첫째'), para('둘째'), para('셋째')];
  s.anims = [];
  addAnim(s, t.id, 'entr', 'fade', { byPara: true });
  addAnim(s, s.objects[0].id, 'path', 'pathArc');
  s.transition = { type: 'morph', dur: 2 };
  const bytes = writePptx(p);
  const xml = textOf(unzip(bytes)['ppt/slides/slide2.xml']);
  assert.match(xml, /<p159:morph option="byObject"\/>/);
  assert.match(xml, /<mc:Fallback><p:transition spd="slow"><p:fade\/>/);
  assert.match(xml, /<p:pRg st="2" end="2"\/>/);
  assert.match(xml, /build="p"/);
  assert.match(xml, /<p:animMotion origin="layout" path="M 0 0 C [^"]+ E"/);
  // WIPOINT 확장 없이 읽기 (PowerPoint 파일처럼)
  const stripped = new Map(Object.entries(unzip(bytes)));
  const sx = textOf(stripped.get('ppt/slides/slide2.xml')).replace(/<p:ext uri="\{6F1B2D57[^]*?<\/p:ext>/, '');
  stripped.set('ppt/slides/slide2.xml', new TextEncoder().encode(sx));
  const back = readPptx(zip(Object.fromEntries(stripped))).pres.slides[1];
  assert.equal(back.transition.type, 'morph');
  const para3 = back.anims.find((a) => a.byPara);
  assert.ok(para3 && para3.effect === 'fade');
  const path = back.anims.find((a) => a.cls === 'path');
  assert.match(path.motion, /^M 0 0 C/);
});

test('animSteps: 단락별은 단락마다 한 단계, 모핑 짝 찾기', async () => {
  const { animSteps, morphPairs, motionPoints } = await import('../src/model.js');
  const p = sample();
  const s = p.slides[1];
  const t = s.objects[1];
  t.text.paras = [para('a'), para(''), para('b')];
  s.anims = [];
  addAnim(s, t.id, 'entr', 'fade', { byPara: true });
  const st = animSteps(s);
  assert.equal(st.length, 2);
  assert.deepEqual(st.map((x) => x[0].para), [0, 2]);
  const a = { objects: [{ id: 'x', type: 'shape', shape: 'rect', name: '로고', x: 0, y: 0, w: 10, h: 10 }] };
  const b = { objects: [{ id: 'y', type: 'shape', shape: 'ellipse', name: '로고', x: 50, y: 0, w: 20, h: 20 }] };
  assert.equal(morphPairs(a, b).length, 1);
  const pts = motionPoints('M 0 0 C 0 0 1 1 1 1', 4);
  assert.deepEqual(pts[pts.length - 1], [1, 1]);
});

test('메모 · 사용자 지정 쇼 저장/읽기 (WIPOINT 확장 없이도)', () => {
  const p = sample();
  p.slides[0].comments = [{ id: 'c1', author: '김대리', text: '제목 확인 부탁', at: '2026-10-10T01:00:00Z', x: 100, y: 50, replies: [{ author: '박과장', text: '확인했습니다', at: '2026-10-10T02:00:00Z' }] }];
  p.customShows = [{ name: '요약본', slides: [p.slides[1].id, p.slides[0].id] }];
  const bytes = writePptx(p);
  const files = unzip(bytes);
  assert.match(textOf(files['ppt/comments/comment1.xml']), /<p:cm authorId="0" dt="2026-10-10T01:00:00" idx="1"><p:pos x="600" y="300"\/><p:text>제목 확인 부탁<\/p:text>/);
  assert.match(textOf(files['ppt/commentAuthors.xml']), /name="김대리"/);
  assert.match(textOf(files['ppt/presentation.xml']), /<p:custShowLst><p:custShow name="요약본" id="0"><p:sldLst><p:sld r:id="rId\d+"\/><p:sld r:id="rId\d+"\/>/);
  const back = readPptx(bytes).pres;
  assert.equal(back.slides[0].comments[0].replies[0].text, '확인했습니다');
  assert.deepEqual(back.customShows[0].slides, [back.slides[1].id, back.slides[0].id]);
  // WIPOINT 확장을 지운 PowerPoint 형식만으로
  const m = new Map(Object.entries(files));
  m.set('ppt/slides/slide1.xml', new TextEncoder().encode(textOf(files['ppt/slides/slide1.xml']).replace(/<p:ext uri="\{6F1B2D57[^]*?<\/p:ext>/, '')));
  const b2 = readPptx(zip(Object.fromEntries(m))).pres;
  const c = b2.slides[0].comments[0];
  assert.equal(c.author, '김대리');
  assert.equal(c.text, '제목 확인 부탁');
  assert.deepEqual([c.x, c.y], [100, 50]);
  assert.equal(c.replies[0].author, '박과장');
});

test('트리거 애니메이션 · 확대/축소 · 슬라이드 링크: pptx 왕복', async () => {
  const { triggerSteps, animSteps } = await import('../src/model.js');
  const p = sample();
  const s = p.slides[1];
  const btn = newShape('rect', { x: 10, y: 10, w: 100, h: 40 });
  btn.name = '단추';
  btn.link = '#slide1';
  const box = newShape('ellipse', { x: 200, y: 200, w: 100, h: 100 });
  box.name = '공';
  s.objects.push(btn, box);
  s.anims = [];
  addAnim(s, box.id, 'entr', 'fade', { trigger: btn.id });
  assert.equal(animSteps(s).length, 0);
  assert.equal(triggerSteps(s).get(btn.id).length, 1);
  const target = p.slides[p.slides.length - 1];
  p.slides[0].objects.push({ id: 'z1', type: 'zoom', x: 100, y: 100, w: 256, h: 144, rot: 0, zoom: { slide: target.id, ret: true } });
  const bytes = writePptx(p);
  const files = unzip(bytes);
  const xml = textOf(files['ppt/slides/slide2.xml']);
  assert.match(xml, /nodeType="interactiveSeq"/);
  assert.match(xml, /action="ppaction:\/\/hlinksldjump"/);
  assert.match(textOf(files['ppt/slides/_rels/slide2.xml.rels']), /relationships\/slide" Target="slide1.xml"/);
  // 우리 확장으로 읽기
  const r1 = readPptx(bytes).pres;
  const z = r1.slides[0].objects.find((o) => o.type === 'zoom');
  assert.equal(z.zoom.slide, r1.slides[r1.slides.length - 1].id);
  const tr = r1.slides[1].anims[0];
  assert.equal(r1.slides[1].objects.find((o) => o.id === tr.trigger)?.link, '#slide1');
  // 확장 없이 (PowerPoint 파일처럼)
  const stripped = new Map(Object.entries(files));
  const sx = xml.replace(/<p:ext uri="\{6F1B2D57[^]*?<\/p:ext>/, '');
  stripped.set('ppt/slides/slide2.xml', new TextEncoder().encode(sx));
  const back = readPptx(zip(Object.fromEntries(stripped))).pres.slides[1];
  const a = back.anims.find((x) => x.trigger);
  assert.ok(a, '트리거 읽기');
  assert.equal(back.objects.find((o) => o.id === a.trigger).link, '#slide1');
});

test('슬라이드 마스터: 마스터 개체 · 제목 서식이 슬라이드에 적용되고 pptx 왕복', async () => {
  const { masterSlide } = await import('../src/model.js');
  const p = sample();
  const m = masterSlide(p, 'master');
  m.objects.push(newShape('ellipse', { x: 1150, y: 20, w: 100, h: 100 }));
  m.objects.find((o) => o.ph === 'title').text.paras[0].runs[0].color = '#00AA00';
  masterSlide(p, 'title').hideMaster = true;
  const files = unzip(writePptx(p));
  const layouts = Object.keys(files).filter((k) => /slideLayouts\/slideLayout\d+\.xml$/.test(k)).map((k) => textOf(files[k]));
  assert.ok(layouts.some((x) => /prst="ellipse"/.test(x)), '마스터 도형이 레이아웃에');
  const s2 = textOf(files['ppt/slides/slide2.xml']);
  assert.match(s2, /<a:srgbClr val="00AA00"\/>/);
  const back = readPptx(zip(files)).pres;
  assert.equal(back.masters.master.objects.filter((o) => !o.ph).length, 1);
  assert.equal(back.masters.title.hideMaster, true);
});

test('수식: LaTeX ⇄ OMML (a14:m) pptx 왕복', async () => {
  const { parseLatex, toMathML } = await import('../src/math.js');
  assert.match(toMathML(parseLatex('\\frac{a}{b}')), /<mfrac><mi>a<\/mi><mi>b<\/mi><\/mfrac>/);
  const p = sample();
  const latex = 'x=\\frac{-b\\pm \\sqrt{b^2-4ac}}{2a}';
  p.slides[1].objects.push({ id: 'eq1', type: 'equation', x: 100, y: 300, w: 500, h: 120, rot: 0, latex, size: 32 });
  const files = unzip(writePptx(p));
  const xml = textOf(files['ppt/slides/slide2.xml']);
  assert.match(xml, /<a14:m [^>]*><m:oMathPara/);
  assert.match(xml, /<m:f><m:num>/);
  const eq = readPptx(zip(files)).pres.slides[1].objects.find((o) => o.type === 'equation');
  assert.equal(eq.latex, latex);
  assert.equal(eq.size, 32);
});
