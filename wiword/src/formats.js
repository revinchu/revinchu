// 문서 형식: 판별 · 읽기 · 쓰기 (DOM 없음 — 브라우저 fileio.js 와 MCP 서버 mcp.js 가 같이 씀)
import { newDoc, para, deserialize, serialize, docText, table, normalizeRuns } from './model.js';
import { readDocx, writeDocx } from './docx.js';
import { readHwpx, writeHwpx } from './hwpx.js';
import { readHwp } from './hwp.js';
import { readHwp3, isHwp3 } from './hwp3.js';
import { readDoc97 } from './doc97.js';
import { readPdfDoc } from './pdfdoc.js';
import { mdToDoc, docToMd } from './md.js';
import { htmlToBlocks, blocksToHtml } from './htmlio.js';
import { readRtf } from './rtf.js';
import { parseCsv } from './csv.js';
import { zip, unzip } from './zip.js';
import { isCfb, openCfb } from './cfb.js';

export const MIME = {
  docx: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
  dotx: 'application/vnd.openxmlformats-officedocument.wordprocessingml.template',
  hwpx: 'application/hwp+zip',
  md: 'text/markdown', html: 'text/html', txt: 'text/plain', wiword: 'application/json', pdf: 'application/pdf', rtf: 'application/rtf',
};

export const OPEN_ACCEPT = '.docx,.docm,.dotx,.dotm,.doc,.hwp,.hwpx,.hwt,.pdf,.md,.markdown,.txt,.html,.htm,.rtf,.odt,.csv,.tsv,.wiword,.json';

// ───────────── 형식 판별 · 읽기 (DOM 없이 바이트에서 문서로 — MCP 서버도 같은 순서) ─────────────
const dec = (b, enc = 'utf-8') => new TextDecoder(enc).decode(b);
/** 한국어 텍스트: UTF-8 이 깨지면 EUC-KR(CP949) */
export function decodeText(bytes) {
  if (bytes[0] === 0xff && bytes[1] === 0xfe) return dec(bytes.subarray(2), 'utf-16le');
  if (bytes[0] === 0xfe && bytes[1] === 0xff) return dec(bytes.subarray(2), 'utf-16be');
  const u = dec(bytes);
  if (!u.includes('\uFFFD')) return u.replace(/^\uFEFF/, '');
  try { return dec(bytes, 'euc-kr'); } catch { return u; }
}
export function detectKind(bytes, name = '') {
  const ext = (/\.([a-z0-9]+)$/i.exec(name)?.[1] ?? '').toLowerCase();
  if (bytes[0] === 0x25 && bytes[1] === 0x50 && bytes[2] === 0x44 && bytes[3] === 0x46) return 'pdf';
  if (bytes[0] === 0x50 && bytes[1] === 0x4b) {
    const files = unzip(bytes);
    if (files['word/document.xml'] || Object.keys(files).some((k) => /^word\/document\d*\.xml$/.test(k))) return 'docx';
    if (files['Contents/section0.xml'] || Object.keys(files).some((k) => /^Contents\/section\d+\.xml$/i.test(k))) return 'hwpx';
    if (files['content.xml'] && /opendocument\.text/.test(dec(files.mimetype ?? new Uint8Array()))) return 'odt';
    return ext === 'hwpx' ? 'hwpx' : 'docx';
  }
  if (isCfb(bytes)) {
    const cfb = openCfb(bytes);
    if (cfb.names.includes('EncryptionInfo')) throw new Error('암호로 보호된 문서입니다. Word 에서 암호를 풀고 저장한 뒤 열어 주세요.');
    if (cfb.names.includes('FileHeader')) return 'hwp';
    if (cfb.names.includes('WordDocument')) return 'doc';
    return ext === 'hwp' ? 'hwp' : 'doc';
  }
  if (bytes[0] === 0x7b && bytes[1] === 0x5c && bytes[2] === 0x72 && bytes[3] === 0x74) return 'rtf';
  if (isHwp3(bytes)) return 'hwp3';
  if (ext === 'hwp') return 'hwp';
  const head = decodeText(bytes.subarray(0, 4096));
  if (/^\s*\{\s*"(version|wiword)"/.test(head)) return 'wiword';
  if (['md', 'markdown', 'mdown'].includes(ext)) return 'md';
  if (['html', 'htm', 'xhtml', 'mht'].includes(ext) || /^\s*(<!doctype html|<html|<meta|<body|<div|<p[\s>])/i.test(head)) return 'html';
  if (ext === 'csv' || ext === 'tsv') return 'csv';
  if (ext === 'rtf') return 'rtf';
  if (/^(#{1,6} |[-*] |```)/m.test(head) && head.split('\n').length > 2) return 'md';
  return 'txt';
}
export function txtToDoc(text) {
  const d = newDoc();
  d.body = text.replace(/\r\n?/g, '\n').split('\n').map((l) => { const p = para(''); p.runs = l ? normalizeRuns(l.split('\t').flatMap((t, i) => (i ? [{ obj: 'tab' }, ...(t ? [{ text: t }] : [])] : t ? [{ text: t }] : []))) : []; return p; });
  if (!d.body.length) d.body = [para('')];
  return d;
}
export function csvToDoc(text) {
  const rows = parseCsv(text);
  const d = newDoc();
  const cols = Math.max(1, ...rows.map((r) => r.length));
  const W = d.section.page.w - d.section.margin.l - d.section.margin.r;
  const t = table(rows.length || 1, cols, { width: W });
  t.rows = rows.map((r, ri) => ({ ...(ri === 0 ? { header: true } : {}), cells: Array.from({ length: cols }, (_, ci) => ({ blocks: [{ ...para(r[ci] ?? ''), pPr: { spc: { after: 0, line: 1, rule: 'auto' } }, runs: r[ci] ? [{ text: r[ci], ...(ri === 0 ? { b: true } : {}) }] : [] }] })) }));
  d.body = [t, para('')];
  return d;
}
/** 바이트 → { doc, warnings, kind, jobs? } */
export function readBytes(bytes, name = '', opts = {}) {
  if (!bytes.length) throw new Error(`'${name}' 파일이 비어 있습니다 (0바이트). 휴대폰이라면 파일을 먼저 내려받은 뒤 다시 열어 주세요.`);
  const kind = detectKind(bytes, name);
  if (kind === 'docx') return { ...readDocx(bytes), kind };
  if (kind === 'hwpx') return { ...readHwpx(bytes, opts), kind };
  if (kind === 'hwp') return { ...readHwp(bytes, opts), kind };
  if (kind === 'hwp3') return { ...readHwp3(bytes, opts), kind };
  if (kind === 'doc') return { ...readDoc97(bytes), kind };
  if (kind === 'pdf') return { ...readPdfDoc(bytes, opts), kind };
  if (kind === 'odt') return { doc: odtToDoc(bytes), warnings: ['OpenDocument 텍스트(.odt)는 글 · 제목 · 표만 가져옵니다.'], kind };
  const text = decodeText(bytes);
  if (kind === 'wiword') return { doc: deserialize(text), warnings: [], kind };
  if (kind === 'md') return { doc: mdToDoc(text), warnings: [], kind };
  if (kind === 'rtf') return { ...readRtf(text), kind };
  if (kind === 'csv') return { doc: csvToDoc(text), warnings: [], kind };
  if (kind === 'html') {
    const d = newDoc();
    const { blocks, media, lists } = htmlToBlocks(text, d);
    d.body = blocks.length ? blocks : [para('')];
    d.media = media ?? {};
    d.lists = { ...d.lists, ...(lists ?? {}) };
    const t = /<title>([^<]*)<\/title>/i.exec(text);
    if (t) d.props.title = t[1].trim();
    return { doc: d, warnings: [], kind };
  }
  return { doc: txtToDoc(text), warnings: [], kind };
}
function odtToDoc(bytes) {
  const files = unzip(bytes);
  const xml = dec(files['content.xml']);
  const d = newDoc();
  const body = [];
  const re = /<text:(h|p)\b([^>]*)>([\s\S]*?)<\/text:\1>|<text:(h|p)\b[^>]*\/>/g;
  let m;
  while ((m = re.exec(xml))) {
    const txt = (m[3] ?? '').replace(/<text:tab\/>/g, '\t').replace(/<text:line-break\/>/g, '\n').replace(/<text:s(?: text:c="(\d+)")?\/>/g, (_, n) => ' '.repeat(Number(n ?? 1))).replace(/<[^>]+>/g, '').replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/&apos;/g, "'").replace(/&amp;/g, '&');
    const lv = m[1] === 'h' ? Number(/text:outline-level="(\d)"/.exec(m[2])?.[1] ?? 1) : 0;
    body.push({ ...para(txt), ...(lv ? { style: `Heading${Math.min(9, lv)}` } : {}) });
  }
  d.body = body.length ? body : [para('')];
  return d;
}

/** 저장할 바이트: kind → Uint8Array | string */
export function docBytes(doc, kind, opts = {}) {
  if (kind === 'docx' || kind === 'dotx') return zip(writeDocx(doc, { template: kind === 'dotx', ...opts }));
  if (kind === 'hwpx') return zip(writeHwpx(doc, opts)); // mimetype 이 첫 항목 (무압축)
  if (kind === 'md') return docToMd(doc, { images: opts.images ?? 'inline' });
  if (kind === 'html') return blocksToHtml(doc, doc.body, { full: true, title: doc.props?.title || opts.name || '' });
  if (kind === 'txt') return docText(doc);
  if (kind === 'wiword') return serialize(doc);
  throw new Error(`저장할 수 없는 형식: ${kind}`);
}

