// EPUB 3 전자책 쓰기 · 읽기 (DOM 없음) — 이음오피스의 "PDF → EPUB" 를 참고해, 문서를 바로 전자책으로 저장하고 전자책도 열 수 있게
// 쓰기: 제목 1(또는 제목 스타일)마다 장 하나(XHTML), 차례(nav.xhtml + toc.ncx), 그림은 파일로 꺼냄, mimetype 은 첫 항목(무압축)
// 읽기: container.xml → OPF 의 spine 순서대로 XHTML 을 문단으로 (그림은 문서 안으로)
import { newDoc, para, uid, docText } from './model.js';
import { blocksToHtml, htmlToBlocks } from './htmlio.js';
import { unzip } from './zip.js';

const esc = (s) => String(s ?? '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
const ENT = { nbsp: 160, emsp: 8195, ensp: 8194, thinsp: 8201, middot: 183, copy: 169, reg: 174, hellip: 8230, mdash: 8212, ndash: 8211, lsquo: 8216, rsquo: 8217, ldquo: 8220, rdquo: 8221, bull: 8226 };
const VOID = /<(br|hr|img|col|input|meta|link|wbr|area|source)\b([^<>]*?)\s*\/?>/gi;
/** HTML 조각 → XHTML (빈 요소 닫기, 이름 있는 문자 참조 → 숫자) */
export function toXhtml(h) {
  return h.replace(VOID, (m, t, a) => `<${t.toLowerCase()}${a}/>`)
    .replace(/&([a-z][a-z0-9]*);/gi, (m, n) => (['amp', 'lt', 'gt', 'quot', 'apos'].includes(n) ? m : ENT[n.toLowerCase()] ? `&#${ENT[n.toLowerCase()]};` : `&amp;${n};`))
    .replace(/&(?!(?:[a-z][a-z0-9]*|#\d+|#x[0-9a-f]+);)/gi, '&amp;');
}
const isHeading1 = (b) => b.type === 'p' && (b.style === 'Heading1' || b.style === 'Title');
const plain = (b) => (b.runs ?? []).map((r) => r.text ?? '').join('').trim();

/** 문서 → { 경로: 바이트|글 } (zip 에 넣을 순서) */
export function writeEpub(doc, { name = '문서', lang = 'ko' } = {}) {
  const title = doc.props?.title || name;
  // 각주 · 미주: 본문의 표시를 번호 달린 위 첨자 링크로 바꾸고, 내용은 마지막 '주' 장으로
  const notes = [];
  const mapRuns = (p) => {
    if (p.type === 'tbl') return { ...p, rows: p.rows.map((r) => ({ ...r, cells: r.cells.map((c) => ({ ...c, blocks: c.blocks.map(mapRuns) })) })) };
    if (!p.runs?.some((r) => r.obj === 'note')) return p;
    return { ...p, runs: p.runs.map((r) => {
      if (r.obj !== 'note') return r;
      const body = doc.notes?.[r.kind ?? 'foot']?.[r.id];
      if (!body) return { text: '' };
      notes.push({ n: notes.length + 1, body });
      return { text: r.custom ?? `${notes.length}`, va: 'super', link: { url: `notes.xhtml#n${notes.length}` } };
    }) };
  };
  doc = { ...doc, body: doc.body.map(mapRuns) };
  // 장 나누기: 제목 1 / 제목 스타일 앞에서
  const chapters = [];
  let cur = null;
  for (const b of doc.body) {
    if (!cur || (isHeading1(b) && cur.blocks.some((x) => x.type === 'tbl' || plain(x) || x.runs?.some((r) => r.obj)))) { cur = { title: isHeading1(b) ? plain(b) : '', blocks: [] }; chapters.push(cur); }
    else if (isHeading1(b) && !cur.title) cur.title = plain(b);
    cur.blocks.push(b);
  }
  if (!chapters.length) chapters.push({ title, blocks: [para('')] });
  // 그림: data URL → 파일
  const images = new Map(); // dataURL → { path, type }
  const files = {};
  const pull = (html) => html.replace(/src="(data:(image\/[a-z0-9.+-]+);base64,([^"]+))"/gi, (m, url, type, b64) => {
    let it = images.get(url);
    if (!it) {
      const ext = { 'image/jpeg': 'jpg', 'image/png': 'png', 'image/gif': 'gif', 'image/svg+xml': 'svg', 'image/webp': 'webp' }[type.toLowerCase()] ?? 'png';
      it = { path: `images/img${images.size + 1}.${ext}`, type: type.toLowerCase() };
      images.set(url, it);
      files[`OEBPS/${it.path}`] = Uint8Array.from(atob(b64.replace(/&#x?[0-9a-f]+;|\s/gi, '')), (c) => c.charCodeAt(0));
    }
    return `src="${it.path}"`;
  });
  const toc = [];
  const chFiles = chapters.map((ch, ci) => {
    let hn = 0;
    let body = pull(toXhtml(blocksToHtml(doc, ch.blocks)));
    // 제목 1~3 에 id → 차례
    body = body.replace(/<h([1-3])(\s[^>]*)?>([\s\S]*?)<\/h\1>/g, (m, lv, attrs = '', inner) => {
      const id = `h${ci + 1}-${++hn}`;
      const text = inner.replace(/<[^>]+>/g, '').replace(/&#(\d+);/g, (x, d) => String.fromCharCode(d)).replace(/&amp;/g, '&').replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').trim();
      if (text) toc.push({ lv: Number(lv), text, href: `ch${ci + 1}.xhtml#${id}` });
      return `<h${lv} id="${id}"${attrs}>${inner}</h${lv}>`;
    });
    const t = ch.title || (ci === 0 ? title : `${ci + 1}`);
    return { path: `ch${ci + 1}.xhtml`, title: t, xhtml: `<?xml version="1.0" encoding="utf-8"?>\n<!DOCTYPE html>\n<html xmlns="http://www.w3.org/1999/xhtml" xmlns:epub="http://www.idpf.org/2007/ops" xml:lang="${lang}" lang="${lang}">\n<head>\n<meta charset="utf-8"/>\n<title>${esc(t)}</title>\n<link rel="stylesheet" type="text/css" href="style.css"/>\n</head>\n<body>\n${body}\n</body>\n</html>\n` };
  });
  if (notes.length) {
    const items = notes.map((nt) => `<li id="n${nt.n}">${pull(toXhtml(blocksToHtml(doc, nt.body))).replace(/^<p[^>]*>|<\/p>$/g, '')}</li>`).join('\n');
    chFiles.push({ path: 'notes.xhtml', title: '주', xhtml: `<?xml version="1.0" encoding="utf-8"?>\n<!DOCTYPE html>\n<html xmlns="http://www.w3.org/1999/xhtml" xmlns:epub="http://www.idpf.org/2007/ops" xml:lang="${lang}" lang="${lang}">\n<head>\n<meta charset="utf-8"/>\n<title>주</title>\n<link rel="stylesheet" type="text/css" href="style.css"/>\n</head>\n<body>\n<section epub:type="endnotes"><h1>주</h1>\n<ol>\n${items}\n</ol>\n</section>\n</body>\n</html>\n` });
    toc.push({ lv: 1, text: '주', href: 'notes.xhtml' });
  }
  if (!toc.length) chFiles.forEach((c) => toc.push({ lv: 1, text: c.title, href: c.path }));
  // 차례 (단계가 건너뛰어도 올바른 목록이 되게)
  const navList = () => {
    let html = '<ol>'; let depth = 1; let open = false;
    for (const it of toc) {
      const lv = Math.max(1, Math.min(it.lv, depth + 1));
      if (lv > depth) { html += '<ol>'; depth = lv; } else { if (open) html += '</li>'; while (depth > lv) { html += '</ol></li>'; depth--; } }
      html += `<li><a href="${it.href}">${esc(it.text)}</a>`; open = true;
    }
    if (open) html += '</li>';
    while (depth > 1) { html += '</ol></li>'; depth--; }
    return `${html}</ol>`;
  };
  const uidv = `urn:uuid:${doc.props?.uuid ?? 'xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx'.replace(/[xy]/g, (c) => { const r = Math.floor(Math.random() * 16); return (c === 'x' ? r : (r & 3) | 8).toString(16); })}`;
  const now = new Date().toISOString().replace(/\.\d+Z$/, 'Z');
  const css = `body{font-family:serif;line-height:1.7;margin:0 4%}h1,h2,h3{font-family:sans-serif;line-height:1.35}h1{font-size:1.6em;margin:1.2em 0 .6em}h2{font-size:1.3em}h3{font-size:1.1em}p{margin:.4em 0;text-align:justify}table{border-collapse:collapse;margin:.6em 0;max-width:100%}td,th{border:1px solid #888;padding:2px 6px;vertical-align:top}img{max-width:100%;height:auto}blockquote{margin:.6em 1.2em;color:#444}pre{white-space:pre-wrap}`;
  const manifest = [
    '<item id="nav" href="nav.xhtml" media-type="application/xhtml+xml" properties="nav"/>',
    '<item id="ncx" href="toc.ncx" media-type="application/x-dtbncx+xml"/>',
    '<item id="css" href="style.css" media-type="text/css"/>',
    ...chFiles.map((c, i) => `<item id="ch${i + 1}" href="${c.path}" media-type="application/xhtml+xml"/>`),
    ...[...images.values()].map((im, i) => `<item id="img${i + 1}" href="${im.path}" media-type="${im.type}"/>`),
  ];
  const meta = [
    `<dc:identifier id="bookid">${uidv}</dc:identifier>`, `<dc:title>${esc(title)}</dc:title>`, `<dc:language>${lang}</dc:language>`,
    doc.props?.author ? `<dc:creator>${esc(doc.props.author)}</dc:creator>` : '', doc.props?.subject ? `<dc:description>${esc(doc.props.subject)}</dc:description>` : '',
    `<meta property="dcterms:modified">${now}</meta>`, '<meta name="generator" content="WIWORD"/>',
  ].filter(Boolean);
  const opf = `<?xml version="1.0" encoding="utf-8"?>\n<package xmlns="http://www.idpf.org/2007/opf" version="3.0" unique-identifier="bookid" xml:lang="${lang}">\n<metadata xmlns:dc="http://purl.org/dc/elements/1.1/">\n${meta.join('\n')}\n</metadata>\n<manifest>\n${manifest.join('\n')}\n</manifest>\n<spine toc="ncx">\n${chFiles.map((c, i) => `<itemref idref="ch${i + 1}"/>`).join('\n')}\n</spine>\n</package>\n`;
  const nav = `<?xml version="1.0" encoding="utf-8"?>\n<!DOCTYPE html>\n<html xmlns="http://www.w3.org/1999/xhtml" xmlns:epub="http://www.idpf.org/2007/ops" xml:lang="${lang}" lang="${lang}">\n<head><meta charset="utf-8"/><title>차례</title><link rel="stylesheet" type="text/css" href="style.css"/></head>\n<body>\n<nav epub:type="toc" id="toc"><h1>차례</h1>\n${navList()}\n</nav>\n</body>\n</html>\n`;
  let po = 0;
  const ncx = `<?xml version="1.0" encoding="utf-8"?>\n<ncx xmlns="http://www.daisy.org/z3986/2005/ncx/" version="2005-1">\n<head><meta name="dtb:uid" content="${uidv}"/></head>\n<docTitle><text>${esc(title)}</text></docTitle>\n<navMap>\n${toc.map((it) => `<navPoint id="np${++po}" playOrder="${po}"><navLabel><text>${esc(it.text)}</text></navLabel><content src="${it.href}"/></navPoint>`).join('\n')}\n</navMap>\n</ncx>\n`;
  return {
    mimetype: 'application/epub+zip',
    'META-INF/container.xml': '<?xml version="1.0" encoding="utf-8"?>\n<container version="1.0" xmlns="urn:oasis:names:tc:opendocument:xmlns:container">\n<rootfiles><rootfile full-path="OEBPS/content.opf" media-type="application/oebps-package+xml"/></rootfiles>\n</container>\n',
    'OEBPS/content.opf': opf,
    'OEBPS/nav.xhtml': nav,
    'OEBPS/toc.ncx': ncx,
    'OEBPS/style.css': css,
    ...Object.fromEntries(chFiles.map((c) => [`OEBPS/${c.path}`, c.xhtml])),
    ...files,
  };
}

// ───────────── 읽기 ─────────────
const td = new TextDecoder();
const attr = (tag, n) => new RegExp(`\\b${n}\\s*=\\s*"([^"]*)"`, 'i').exec(tag)?.[1] ?? new RegExp(`\\b${n}\\s*=\\s*'([^']*)'`, 'i').exec(tag)?.[1];
const join = (base, rel) => {
  const parts = (base ? `${base}/${rel}` : rel).split('/');
  const out = [];
  for (const p of parts) { if (p === '..') out.pop(); else if (p && p !== '.') out.push(p); }
  return out.join('/');
};
export const isEpub = (files) => files.mimetype && td.decode(files.mimetype).trim() === 'application/epub+zip';
/** EPUB 바이트 → { doc, warnings } */
export function readEpub(bytes) {
  const files = unzip(bytes);
  const cont = files['META-INF/container.xml'] ? td.decode(files['META-INF/container.xml']) : '';
  const opfPath = attr(/<rootfile\b[^>]*>/i.exec(cont)?.[0] ?? '', 'full-path') ?? Object.keys(files).find((k) => k.endsWith('.opf'));
  if (!opfPath || !files[opfPath]) throw new Error('EPUB 의 목차 파일(content.opf)을 찾지 못했습니다.');
  if (files['META-INF/encryption.xml'] && /EncryptedData/.test(td.decode(files['META-INF/encryption.xml'])) && !/obfuscation/i.test(td.decode(files['META-INF/encryption.xml']))) throw new Error('DRM 으로 보호된 전자책은 열 수 없습니다.');
  const opf = td.decode(files[opfPath]);
  const base = opfPath.includes('/') ? opfPath.slice(0, opfPath.lastIndexOf('/')) : '';
  const items = {};
  for (const m of opf.matchAll(/<item\b[^>]*>/gi)) items[attr(m[0], 'id')] = { href: decodeURIComponent(attr(m[0], 'href') ?? ''), type: attr(m[0], 'media-type') ?? '' };
  const spine = [...opf.matchAll(/<itemref\b[^>]*>/gi)].map((m) => attr(m[0], 'idref')).filter((id) => items[id] && /html/.test(items[id].type));
  const d = newDoc();
  const t = /<dc:title[^>]*>([^<]*)<\/dc:title>/i.exec(opf);
  if (t) d.props.title = t[1].trim();
  const a = /<dc:creator[^>]*>([^<]*)<\/dc:creator>/i.exec(opf);
  if (a) d.props.author = a[1].trim();
  const body = [];
  const MIME = { jpg: 'image/jpeg', jpeg: 'image/jpeg', png: 'image/png', gif: 'image/gif', svg: 'image/svg+xml', webp: 'image/webp' };
  const b64 = (u8) => { let s = ''; for (let i = 0; i < u8.length; i += 0x8000) s += String.fromCharCode(...u8.subarray(i, i + 0x8000)); return btoa(s); };
  for (const [ci, id] of spine.entries()) {
    const path = join(base, items[id].href);
    if (!files[path]) continue;
    const dir = path.includes('/') ? path.slice(0, path.lastIndexOf('/')) : '';
    let html = td.decode(files[path]);
    if (/\bepub:type\s*=\s*"[^"]*\btoc\b/.test(html) && items[id].href.includes('nav')) continue; // 차례 쪽은 건너뜀
    // 그림 → data URL (img src, svg image xlink:href)
    html = html.replace(/(<img\b[^>]*?\bsrc|<image\b[^>]*?\bxlink:href|<image\b[^>]*?\bhref)\s*=\s*"([^"]+)"/gi, (m, pre, src) => {
      if (/^(data:|https?:)/i.test(src)) return m;
      const p = join(dir, decodeURIComponent(src.split('#')[0]));
      const f = files[p];
      if (!f) return m;
      const ext = (/\.([a-z0-9]+)$/i.exec(p)?.[1] ?? '').toLowerCase();
      return `${pre.replace(/<image\b/i, '<img').replace(/xlink:href|href$/i, 'src')}="data:${MIME[ext] ?? 'image/png'};base64,${b64(f)}"`;
    });
    const { blocks, media, lists } = htmlToBlocks(html, d);
    Object.assign(d.media, media ?? {});
    if (lists) d.lists = { ...d.lists, ...lists };
    if (ci > 0 && body.length && blocks.length) blocks[0] = { ...blocks[0], pPr: { ...(blocks[0].pPr ?? {}), pageBreakBefore: true } };
    body.push(...blocks);
  }
  d.body = body.length ? body : [para('')];
  for (const b of d.body) if (!b.id) b.id = uid(b.type === 'tbl' ? 't' : 'p');
  return { doc: d, warnings: [`전자책(EPUB)을 열었습니다: 장 ${spine.length}개 · 글자 ${docText(d).length.toLocaleString()}자. 저장은 .docx · .epub 등으로.`] };
}
