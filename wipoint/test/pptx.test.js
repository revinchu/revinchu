import { test } from 'node:test';
import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { dirname, join } from 'node:path';
import { writePptx, readPptx, pptxEntries, chartXml } from '../src/pptx.js';
import { unzip, textOf } from '../src/zip.js';
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
