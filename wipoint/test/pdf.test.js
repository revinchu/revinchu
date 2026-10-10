import { test } from 'node:test';
import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { dirname, join } from 'node:path';
import { readPdf, pngDataUrl } from '../src/pdf.js';
import { writePptx, readPptx } from '../src/pptx.js';
import { plainText } from '../src/model.js';

const here = dirname(fileURLToPath(import.meta.url));
const load = (n) => new Uint8Array(readFileSync(join(here, 'fixtures', n)));
const texts = (s) => s.objects.filter((o) => o.text).map((o) => plainText(o.text));

// libreoffice.pdf: make-pdf-source.py 로 만든 pptx 를 LibreOffice 로 PDF 변환
test('PDF (LibreOffice): 쪽 → 슬라이드, 글 · 도형 · 그림', () => {
  const { pres, warnings, jobs } = readPdf(load('libreoffice.pdf'));
  assert.equal(pres.slides.length, 3);
  assert.deepEqual(pres.size, { w: 1280, h: 720 });
  const t0 = texts(pres.slides[0]).join('\n');
  assert.match(t0, /네이버 검색광고 성장 전략/);
  assert.match(t0, /그리고 압도적 성장을 위한 로드맵/);
  // 여러 줄 글머리 목록은 한 텍스트 상자
  const list = pres.slides[1].objects.find((o) => o.text && plainText(o.text).includes('주차 관제 시스템'));
  assert.ok(list);
  assert.equal(list.text.paras.length, 4);
  assert.equal(list.text.wrap, false);
  // 채운 둥근 사각형 (채우기 + 선 한 도형)
  assert.ok(pres.slides[1].objects.some((o) => o.fill?.color === '#2E4BA0' && o.line));
  // 그림 (JPEG + 투명 가리개는 jobs 로)
  assert.ok(pres.slides[2].objects.some((o) => o.type === 'image'));
  assert.ok(jobs.length >= 1 && jobs[0].alpha.length === jobs[0].w * jobs[0].h);
  // 차트 막대 = 사각형 도형, 축 글자
  assert.equal(pres.slides[2].objects.filter((o) => o.fill?.color === '#4F81BD').length, 3);
  assert.ok(texts(pres.slides[2]).some((t) => t.replace(/\s/g, '') === '1월'));
  assert.deepEqual(warnings, []);
});

test('PDF (Chrome): CID 글꼴 · 한글 · 굵게 · 색', () => {
  const { pres } = readPdf(load('chrome.pdf'));
  assert.equal(pres.slides.length, 4);
  const s = pres.slides[1];
  const big = s.objects.find((o) => o.text && plainText(o.text).startsWith('1,284'));
  assert.ok(big);
  assert.equal(big.text.paras[0].runs[0].b, true);
  assert.equal(big.text.paras[0].runs[0].size, 36);
  assert.match(texts(s).join(' '), /핵심 지표 요약/);
  assert.match(texts(pres.slides[2]).join(' '), /네이버 검색/);
});

test('PDF → pptx 저장 후 다시 열기', () => {
  const { pres } = readPdf(load('libreoffice.pdf'));
  const back = readPptx(writePptx(pres)).pres;
  assert.equal(back.slides.length, 3);
  assert.match(texts(back.slides[0]).join(' '), /성장 전략/);
});

test('PDF: 잘못된 파일 · 암호', () => {
  assert.throws(() => readPdf(new TextEncoder().encode('hello')), /PDF 파일이 아닙니다/);
  const enc = new TextEncoder().encode('%PDF-1.4\n1 0 obj<</Type/Catalog/Pages 2 0 R>>endobj\n2 0 obj<</Type/Pages/Kids[3 0 R]/Count 1>>endobj\n3 0 obj<</Type/Page/MediaBox[0 0 100 100]>>endobj\ntrailer<</Root 1 0 R/Encrypt 4 0 R>>\n%%EOF');
  assert.throws(() => readPdf(enc), /암호/);
});

test('PDF: 손으로 쓴 최소 문서 (Flate 없음, Tj · re · 회전)', () => {
  const content = 'q 1 0 0 rg 10 10 30 20 re f Q BT /F1 12 Tf 20 60 Td (Hello) Tj ET';
  const pdf = `%PDF-1.4
1 0 obj<</Type/Catalog/Pages 2 0 R>>endobj
2 0 obj<</Type/Pages/Kids[3 0 R]/Count 1>>endobj
3 0 obj<</Type/Page/Parent 2 0 R/MediaBox[0 0 200 100]/Rotate 0/Contents 4 0 R/Resources<</Font<</F1 5 0 R>>>>>>endobj
4 0 obj<</Length ${content.length}>>stream
${content}
endstream endobj
5 0 obj<</Type/Font/Subtype/Type1/BaseFont/Helvetica-Bold>>endobj
trailer<</Root 1 0 R>>
%%EOF`;
  const { pres } = readPdf(new TextEncoder().encode(pdf));
  assert.deepEqual(pres.size, { w: 267, h: 133 });
  const s = pres.slides[0];
  const r = s.objects.find((o) => o.fill?.color === '#FF0000');
  assert.ok(r);
  assert.equal(r.shape, 'rect');
  assert.equal(Math.round(r.y), Math.round((100 - 30) * 4 / 3));
  const t = s.objects.find((o) => o.text);
  assert.equal(plainText(t.text), 'Hello');
  assert.equal(t.text.paras[0].runs[0].b, true);
  assert.equal(t.text.paras[0].runs[0].font, 'Arial');
});

test('PNG 인코더 (DOM 없음)', () => {
  const url = pngDataUrl(2, 1, new Uint8Array([255, 0, 0, 255, 0, 0, 255, 128]));
  assert.match(url, /^data:image\/png;base64,iVBORw0KGgo/);
});
