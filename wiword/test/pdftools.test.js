// PDF 도구 (pdftools.js) — 손으로 만든 작은 PDF 와 직접 만든 암호 PDF 로 확인
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import { readPdf, pngDataUrl } from '../src/pdf.js';
import * as T from '../src/pdftools.js';

const enc = new TextEncoder();
/** 쪽마다 글 하나 + (선택) 메모 · 링크 주석, 빈 쪽 */
function makePdf(pages) {
  const objs = []; const add = (s) => { objs.push(s); return objs.length; };
  const font = add('<</Type/Font/Subtype/Type1/BaseFont/Helvetica>>');
  const kids = [];
  const pagesIdx = add(null);
  for (const p of pages) {
    const c = p.text ? `BT /F1 20 Tf 30 60 Td (${p.text}) Tj ET` : '';
    const cs = add(`<</Length ${c.length}>>stream\n${c}\nendstream`);
    const annots = (p.annots ?? []).map((t) => add(t === 'Link' ? '<</Type/Annot/Subtype/Link/Rect[0 0 50 20]/A<</S/URI/URI(https://example.com/)>>>>' : `<</Type/Annot/Subtype/${t}/Rect[0 0 20 20]/Contents(memo)>>`));
    kids.push(add(`<</Type/Page/Parent ${pagesIdx} 0 R/MediaBox[0 0 300 120]${p.rotate ? `/Rotate ${p.rotate}` : ''}/Resources<</Font<</F1 ${font} 0 R>>>>/Contents ${cs} 0 R${annots.length ? `/Annots[${annots.map((a) => `${a} 0 R`).join(' ')}]` : ''}>>`));
  }
  objs[pagesIdx - 1] = `<</Type/Pages/Kids[${kids.map((k) => `${k} 0 R`).join(' ')}]/Count ${kids.length}>>`;
  const first = kids[0];
  const ol1 = objs.length + 2; const ol2 = objs.length + 3;
  const root = add(`<</Type/Catalog/Pages ${pagesIdx} 0 R/Outlines ${objs.length + 2} 0 R>>`);
  add(`<</Type/Outlines/First ${ol1 + 1} 0 R/Last ${ol2 + 1} 0 R/Count 2>>`);
  add(`<</Title(First)/Parent ${ol1} 0 R/Next ${ol2 + 1} 0 R/Dest[${first} 0 R/XYZ 0 100 0]>>`);
  add(`<</Title(Last)/Parent ${ol1} 0 R/Prev ${ol1 + 1} 0 R/Dest[${kids[kids.length - 1]} 0 R/Fit]>>`);
  let out = '%PDF-1.4\n'; const offs = [];
  objs.forEach((o, i) => { offs.push(out.length); out += `${i + 1} 0 obj\n${o}\nendobj\n`; });
  const x = out.length;
  out += `xref\n0 ${objs.length + 1}\n0000000000 65535 f \n${offs.map((o) => `${String(o).padStart(10, '0')} 00000 n \n`).join('')}trailer\n<</Size ${objs.length + 1}/Root ${root} 0 R>>\nstartxref\n${x}\n%%EOF\n`;
  return enc.encode(out);
}
const texts = (bytes, opts) => readPdf(bytes, opts).pres.slides.map((s) => s.objects.flatMap((o) => o.text?.paras ?? []).flatMap((p) => p.runs).map((r) => r.t).join('').trim());
const SRC = makePdf([{ text: 'Page1', annots: ['Text', 'Link'] }, { text: 'Page2', rotate: 90, annots: ['Highlight'] }, { text: 'Page3' }, {}]);
const fx = (f) => new Uint8Array(readFileSync(new URL(`./fixtures/${f}`, import.meta.url)));

test('PDF 도구: 정보 · 페이지 지정 · 책갈피', () => {
  const info = T.pdfInfo(SRC);
  assert.equal(info.pages.length, 4);
  assert.deepEqual(info.pages[1], { w: 120, h: 300, rotate: 90 });
  assert.equal(info.outline, 2);
  assert.deepEqual(info.annots, { Text: 1, Link: 1, Highlight: 1 });
  assert.deepEqual(T.parsePages('1-3, 5, 8-', 9), [0, 1, 2, 4, 7, 8]);
  assert.deepEqual(T.parsePages('3-1', 5), [2, 1, 0]);
  assert.deepEqual(T.parsePages('', 3), [0, 1, 2]);
  assert.throws(() => T.parsePages('a-b', 3), /페이지 지정/);
});

test('PDF 도구: 추출 · 삭제 · 역순 · 회전 · 분할', async () => {
  assert.deepEqual(texts(await T.extractPages(SRC, '3,1')), ['Page3', 'Page1']);
  assert.deepEqual(texts(await T.deletePages(SRC, '2,4')), ['Page1', 'Page3']);
  assert.deepEqual(texts(await T.reversePages(SRC)), ['', 'Page3', 'Page2', 'Page1']);
  const rot = T.pdfInfo(await T.rotatePages(SRC, 90, '1-2'));
  assert.deepEqual(rot.pages.map((p) => p.rotate), [90, 180, 0, 0]);
  // 책갈피는 남은 페이지로 옮겨짐 (지운 페이지를 가리키는 것은 빠짐)
  assert.equal(T.pdfInfo(await T.deletePages(SRC, '4')).outline, 1);
  const parts = await T.splitPdf(SRC, { every: 3, base: 'x' });
  assert.deepEqual(parts.map((p) => p.name), ['x_1-3.pdf', 'x_4.pdf']);
  assert.deepEqual(parts.map((p) => T.pdfInfo(p.bytes).pages.length), [3, 1]);
});

test('PDF 도구: 병합 (암호 문서 포함) · 양면 인쇄 맞추기 · 모아찍기', async () => {
  const merged = await T.mergePdfs([{ bytes: SRC, name: '원본' }, { bytes: fx('hello-aes128.pdf'), name: '암호', password: '1234' }, { bytes: fx('hello-aes256.pdf') }]);
  const t = texts(merged);
  assert.equal(t.length, 6);
  assert.match(t[4], /Hello WIWORD/);
  assert.equal(T.pdfInfo(merged).encrypted, false);
  assert.equal(T.pdfInfo(merged).outline, 5); // 파일 3개 + 원래 책갈피 2개
  const duplex = await T.duplexMerge([{ bytes: fx('hello-rc4.pdf') }, { bytes: SRC, pages: '1-3' }], { bookmarks: false });
  assert.deepEqual(texts(duplex).map(Boolean), [true, false, true, true, true, false]);
  const nup = await T.nUp(SRC, { n: 4 });
  const ni = T.pdfInfo(nup);
  assert.equal(ni.pages.length, 1);
  // 칸이 더 커지는 용지 방향을 고름: 가로로 긴 페이지 4쪽 → 가로 용지, 2쪽 → 세로 용지 (위아래로)
  assert.ok(ni.pages[0].w > ni.pages[0].h);
  assert.ok(T.pdfInfo(await T.nUp(SRC, { n: 2 })).pages[0].w < 600);
  assert.ok(T.pdfInfo(await T.nUp(fx('hello-rc4.pdf'), { n: 4 })).pages[0].w > 800);
});

test('PDF 도구: 주석 지우기 · 빈 페이지 · 도장 · 배경 · 정보 · 암호 해제', async () => {
  const r = await T.removeAnnotations(SRC);
  assert.equal(r.removed, 2);
  assert.deepEqual(T.pdfInfo(r).annots, { Link: 1 });
  assert.deepEqual(T.pdfInfo(await T.removeAnnotations(SRC, { types: ['Link'] })).annots, { Text: 1, Highlight: 1 });
  const an = T.analyzePages(SRC);
  assert.deepEqual(an.map((p) => p.blank), [false, false, false, true]);
  const st = await T.stampPdf(SRC, { numbers: { format: '- {n} / {total} -' }, watermark: { text: 'DRAFT' } });
  const words = texts(st);
  assert.match(words[0], /Page1/);
  assert.match(words.join(' '), /- 3 \/ 4 -/);
  assert.match(words.join(' '), /DRAFT/);
  const ov = await T.overlayPdf(SRC, fx('hello-rc4.pdf'), { behind: true, pages: '1' });
  assert.match(texts(ov)[0], /Hello WIWORD/);
  assert.equal(texts(ov)[1], 'Page2');
  const meta = await T.setPdfInfo(SRC, { title: '새 제목', author: 'WIWORD' });
  assert.deepEqual(T.pdfInfo(meta).info.title, '새 제목');
  const dec = await T.decryptPdf(fx('hello-aes128.pdf'), '1234');
  assert.equal(T.pdfInfo(dec).encrypted, false);
  assert.match(texts(dec)[0], /Hello WIWORD/);
  await assert.rejects(T.decryptPdf(fx('hello-aes128.pdf'), 'x'), (e) => e.code === 'PASSWORD');
  assert.ok((await T.compactPdf(SRC)).length > 0);
});

test('PDF 도구: 이미지로 PDF (PNG 투명도 · 크기)', async () => {
  const rgba = new Uint8Array(40 * 20 * 4);
  for (let i = 0; i < 40 * 20; i++) { rgba[i * 4] = 255; rgba[i * 4 + 3] = i % 40 < 20 ? 255 : 0; }
  const png = Uint8Array.from(atob(pngDataUrl(40, 20, rgba).split(',')[1]), (c) => c.charCodeAt(0));
  const pdf = await T.imagesToPdf([{ bytes: png, name: 'a.png' }], { paper: 'fit' });
  const info = T.pdfInfo(pdf);
  assert.deepEqual(info.pages[0], { w: 30, h: 15, rotate: 0 }); // 96dpi → pt
  assert.ok(new TextDecoder('latin1').decode(pdf).includes('/SMask'));
  const a4 = T.pdfInfo(await T.imagesToPdf([{ bytes: png }], { paper: 'A4' }));
  assert.deepEqual([a4.pages[0].w, a4.pages[0].h], [842, 595]); // 가로 그림 → 가로 용지
  await assert.rejects(T.imagesToPdf([{ bytes: enc.encode('GIF89a'), name: 'x.gif' }]), /JPEG · PNG/);
});

test('PDF 글자 고치기: 상자 바꾸기 · 지우기 · 새 글 (원래 글은 파일에서 사라짐)', async () => {
  const two = makePdf([{ text: 'Page1' }, { text: 'Keep' }]);
  const v = T.editablePage(T.openPdf(two), 0);
  assert.equal(v.blocks.length, 1);
  assert.equal(v.blocks[0].text, 'Page1');
  assert.equal(v.blocks[0].size, 20);
  const out = await T.applyTextEdits(two, [{ page: 0, block: 0, text: 'Hello' }, { page: 0, x: 20, y: 10, text: 'Added', size: 9 }, { page: 1, block: 0, text: '' }]);
  const t = texts(out);
  assert.match(t[0], /Hello/);
  assert.match(t[0], /Added/);
  assert.doesNotMatch(t[0], /Page1/);
  assert.equal(t[1], '');
  assert.ok(!new TextDecoder('latin1').decode(out).includes('(Page1)'));
  // 원래 글 위치 그대로 (기준선 · 왼쪽)
  const after = T.editablePage(T.openPdf(out), 0).blocks.find((b) => b.text === 'Hello');
  assert.ok(Math.abs(after.x - v.blocks[0].x) < 0.5 && Math.abs(after.lines[0].y - v.blocks[0].lines[0].y) < 0.5);
  // 암호 문서도 고침
  const enc = await T.applyTextEdits(fx('hello-aes128.pdf'), [{ page: 0, block: 0, text: 'Bye' }], { password: '1234' });
  assert.match(texts(enc)[0], /Bye/);
});

test('PDF → 문서: 옆으로 누운 쪽은 바로 세워 읽기 · docx/hwpx/md/epub 변환', async () => {
  const { readBytes, docBytes } = await import('../src/formats.js');
  const { docText } = await import('../src/model.js');
  const rot = await T.rotatePages(makePdf([{ text: 'Upright text here' }]), 90);
  const r = readBytes(rot, 'r.pdf');
  assert.match(docText(r.doc), /Upright text here/);
  for (const k of ['docx', 'hwpx', 'md', 'epub']) {
    const out = docBytes(r.doc, k, { name: 'r' });
    const back = readBytes(typeof out === 'string' ? new TextEncoder().encode(out) : out, `r.${k}`);
    assert.match(docText(back.doc), /Upright text here/, k);
  }
});
