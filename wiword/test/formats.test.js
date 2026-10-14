// 형식 읽기 · 쓰기 왕복 (node --test)
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { readFileSync, readdirSync, existsSync } from 'node:fs';
import { readBytes, docBytes, detectKind, decodeText } from '../src/formats.js';
import { readDocx, writeDocx } from '../src/docx.js';
import { readHwpx, writeHwpx, hwpEqToLatex, latexToHwpEq } from '../src/hwpx.js';
import { mdToDoc, docToMd } from '../src/md.js';
import { readRtf } from '../src/rtf.js';
import { parseCsv, toCsv } from '../src/csv.js';
import { zip, unzip } from '../src/zip.js';
import { docText, newDoc, para, table, walkParas, headings } from '../src/model.js';
import { sampleDoc, TEMPLATES } from '../src/sample.js';

const enc = new TextEncoder();

test('마크다운 → 문서 → 마크다운', () => {
  const src = '# 제목\n\n본문 **굵게** *기울임* `코드` [링크](https://a.kr) snake_case\n\n- 하나\n- 둘\n  - 하위\n\n1. 첫째\n2. 둘째\n\n> 인용\n\n```js\nconst a = 1;\n```\n\n| 이름 | 값 |\n|:---|---:|\n| a | 1 |\n\n각주[^1]\n\n[^1]: 각주 내용\n\n- [ ] 할 일\n- [x] 완료\n';
  const d = mdToDoc(src);
  assert.equal(d.body[0].style, 'Heading1');
  assert.ok(d.body.some((b) => b.type === 'tbl'));
  const r = d.body[1].runs;
  assert.ok(r.some((x) => x.b && x.text === '굵게'));
  assert.ok(r.some((x) => x.link?.url === 'https://a.kr'));
  const md = docToMd(d);
  for (const s of ['# 제목', '**굵게**', '*기울임*', '`코드`', '[링크](https://a.kr)', 'snake_case', '- 하나', '   - 하위', '1. 첫째', '2. 둘째', '> 인용', '```js', '| a | 1 |', '[^1]: 각주 내용', '- [ ] 할 일', '- [x] 완료']) assert.ok(md.includes(s), `없음: ${s}\n${md}`);
  // 두 번 돌려도 같은 결과
  assert.equal(docToMd(mdToDoc(md)), md);
});

test('DOCX 쓰기 → 읽기: 글 · 스타일 · 표 · 각주 · 머리글 그대로', () => {
  const d = sampleDoc();
  const out = zip(writeDocx(d));
  const files = unzip(out);
  for (const p of ['[Content_Types].xml', 'word/document.xml', 'word/styles.xml', 'word/numbering.xml', 'word/settings.xml', 'word/footnotes.xml', 'docProps/core.xml']) assert.ok(files[p], `${p} 없음`);
  const xml = new TextDecoder().decode(files['word/document.xml']);
  assert.ok(xml.includes('<w:tbl>'));
  assert.ok(/<w:footnoteReference w:id="1"\/>/.test(xml));
  // WIWORD 원본 없이 (Word 에서 고친 것처럼) 읽어도 같은 글
  delete files['word/wiword.json'];
  const back = readDocx(files).doc;
  assert.equal(docText(back), docText(d));
  assert.deepEqual(headings(back).map((h) => h.text), headings(d).map((h) => h.text));
  assert.ok(back.body.some((b) => b.type === 'tbl'));
  assert.ok(Object.keys(back.notes.foot).length >= 1);
});

test('DOCX: WIWORD 가 저장한 파일은 정확히 복원', () => {
  const d = sampleDoc();
  const back = readDocx(zip(writeDocx(d))).doc;
  assert.equal(JSON.stringify(back.body), JSON.stringify(d.body));
});

test('HWPX 쓰기 → 읽기 (모든 서식 파일)', () => {
  for (const t of TEMPLATES) {
    const d = t.build();
    const entries = writeHwpx(d);
    assert.equal(Object.keys(entries)[0], 'mimetype');
    assert.equal(entries.mimetype, 'application/hwp+zip');
    const back = readHwpx(zip(entries)).doc;
    const norm = (s) => s.replace(/[\t\n]/g, ''); // 쪽 나누기는 HWPX 에서 문단 속성 (글자 수는 같음)
    assert.equal(norm(docText(back)), norm(docText(d)), t.id);
  }
});

test('HWPX: 표 병합 · 머리글 · 각주 · 쪽 크기', () => {
  const d = newDoc();
  const t = table(3, 3, { width: 600 });
  t.rows[0].cells = [{ span: 2, blocks: [para('가로 병합')] }, { vMerge: 'restart', blocks: [para('세로 병합')] }];
  t.rows[1].cells[2] = { vMerge: 'cont', blocks: [para('')] };
  d.body = [para('본문'), t, { ...para('각주 있음'), runs: [{ text: '각주 있음' }, { obj: 'note', id: 'n1', kind: 'foot' }] }];
  d.notes.foot.n1 = [para('각주 내용')];
  d.section.headers = { default: [para('머리글 글')] };
  d.section.page = { w: 1122.5, h: 793.7, orient: 'landscape' };
  const back = readHwpx(zip(writeHwpx(d))).doc;
  const bt = back.body.find((b) => b.type === 'tbl');
  assert.equal(bt.rows[0].cells[0].span, 2);
  assert.equal(bt.rows[0].cells[1].vMerge, 'restart');
  assert.equal(bt.rows[1].cells[2].vMerge, 'cont');
  assert.equal(back.section.page.orient, 'landscape');
  assert.ok(Math.abs(back.section.page.w - 1122.5) < 1);
  assert.equal(docText({ body: back.section.headers.default }), '머리글 글');
  assert.equal(Object.values(back.notes.foot).length, 1);
});

test('한글 수식 ↔ LaTeX', () => {
  assert.equal(hwpEqToLatex('{a+b} over {c}'), '\\frac{a+b}{c}');
  assert.equal(latexToHwpEq('\\frac{1}{2}'), '{1} over {2}');
  assert.ok(hwpEqToLatex('sqrt {x} + alpha').includes('\\sqrt{x}'));
});

test('형식 판별', () => {
  assert.equal(detectKind(zip(writeDocx(newDoc())), 'a.docx'), 'docx');
  assert.equal(detectKind(zip(writeHwpx(newDoc())), 'a.hwpx'), 'hwpx');
  assert.equal(detectKind(enc.encode('%PDF-1.4'), 'x'), 'pdf');
  assert.equal(detectKind(enc.encode('{\\rtf1 x}'), 'x'), 'rtf');
  assert.equal(detectKind(enc.encode('# 제목\n\n- a\n- b\n'), 'a.txt'), 'md');
  assert.equal(detectKind(enc.encode('<!doctype html><p>x</p>'), 'a'), 'html');
  assert.equal(detectKind(enc.encode('그냥 글'), 'a.txt'), 'txt');
});

test('텍스트 · EUC-KR · CSV · HTML · RTF 읽기', () => {
  const euckr = new Uint8Array([0xc7, 0xd1, 0xb1, 0xdb]); // '한글'
  assert.equal(decodeText(euckr), '한글');
  assert.equal(docText(readBytes(enc.encode('첫 줄\n둘째\t탭'), 'a.txt').doc), '첫 줄\n둘째\t탭');
  const csv = readBytes(enc.encode('이름,값\n"가, 나",1\n'), 'a.csv').doc;
  assert.equal(csv.body[0].type, 'tbl');
  assert.deepEqual(parseCsv('a,"b ""c""",d\n1,2,3'), [['a', 'b "c"', 'd'], ['1', '2', '3']]);
  assert.equal(toCsv([['a,b', 'c']]), '"a,b",c');
  const html = readBytes(enc.encode('<html><title>제목</title><h1>큰 제목</h1><p>본문 <b>굵게</b></p><ul><li>하나</li></ul></html>'), 'a.html').doc;
  assert.equal(html.props.title, '제목');
  assert.equal(html.body[0].style, 'Heading1');
  const rtf = readRtf("{\\rtf1\\ansi\\ansicpg949{\\fonttbl{\\f0 Arial;}}\\f0\\b \\'c7\\'d1\\'b1\\'db\\b0  A\\u54620?B\\par\\intbl x\\cell y\\cell\\row}").doc;
  assert.equal(rtf.body[0].runs[0].text, '한글');
  assert.ok(rtf.body[0].runs[0].b);
  assert.ok(docText(rtf).includes('A한B'));
  assert.equal(rtf.body.find((b) => b.type === 'tbl').rows[0].cells.length, 2);
});

test('docBytes: 모든 저장 형식', () => {
  const d = sampleDoc();
  for (const k of ['docx', 'hwpx', 'md', 'html', 'txt', 'wiword', 'dotx']) {
    const out = docBytes(d, k);
    assert.ok(out.length > 100, k);
    const back = readBytes(typeof out === 'string' ? enc.encode(out) : out, `x.${k === 'wiword' ? 'wiword' : k}`);
    assert.ok(docText(back.doc).includes('WIWORD'), k);
  }
});

// 실제 파일 모음 (있을 때만): WIWORD_CORPUS=폴더 (docx · hwp · hwpx · doc · pdf)
const corpus = process.env.WIWORD_CORPUS;
test('실제 문서 모음 읽기', { skip: !corpus || !existsSync(corpus) }, () => {
  let n = 0;
  for (const f of readdirSync(corpus).filter((x) => /\.(docx|hwp|hwpx|doc|pdf)$/i.test(x))) {
    try { const r = readBytes(new Uint8Array(readFileSync(`${corpus}/${f}`)), f); assert.ok(r.doc.body.length); n++; } catch (e) { if (!/암호|배포용|손상|잘려|지원하지|OLE|ZIP|없습니다|이전/.test(e.message)) throw new Error(`${f}: ${e.message}`); }
  }
  assert.ok(n > 0);
});

test('표 · 문단 걷기 (walkParas) 가 표 안 문단까지', () => {
  const d = newDoc();
  const t = table(2, 2);
  t.rows[1].cells[1].blocks = [para('안쪽')];
  d.body = [para('밖'), t];
  const seen = [];
  walkParas(d.body, (p) => seen.push(p.runs.map((r) => r.text).join('')));
  assert.ok(seen.includes('안쪽'));
});
