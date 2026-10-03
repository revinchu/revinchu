// OpenDocument 스프레드시트(.ods / .fods) 읽기 · 쓰기 — DOM 없음
// 값 · 수식(of:= ↔ A1) · 병합 · 열 너비 · 행 높이 · 기본 서식(굵게 · 기울임 · 글자색 · 채우기 · 크기 · 맞춤 · 숫자 형식)
import { unzip, zip, textOf } from './zip.js';
import { parseXml, child, kids } from './xml.js';
import { CellMap } from './cellmap.js';

const esc = (s) => String(s ?? '').replace(/[&<>"]/g, (m) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[m]))
  .replace(/[\u0000-\u0008\u000b\u000c\u000e-\u001f]/g, '');

// 길이 단위 → 픽셀 (96dpi)
function toPx(len) {
  const m = /^([\d.]+)\s*(cm|mm|in|pt|pc|px)?$/.exec(String(len ?? '').trim());
  if (!m) return null;
  const v = Number(m[1]);
  const f = { cm: 96 / 2.54, mm: 96 / 25.4, in: 96, pt: 96 / 72, pc: 16, px: 1 }[m[2] ?? 'px'];
  return Math.round(v * f);
}
const pxToCm = (px) => `${(px * 2.54 / 96).toFixed(3)}cm`;

// ─────────────── 수식 변환 ───────────────
/** ODF 수식(of:=SUM([.A1:.B2];[$Sheet2.C3])) → A1 수식 본문 (앞 '=' 없음) */
export function fromOdfFormula(f) {
  let s = String(f).replace(/^[a-z]+:/i, '').replace(/^=/, '');
  let out = '';
  for (let i = 0; i < s.length; i++) {
    const ch = s[i];
    if (ch === '"') {
      const j = s.indexOf('"', i + 1);
      let k = j;
      while (k >= 0 && s[k + 1] === '"') k = s.indexOf('"', k + 2);
      const end = k < 0 ? s.length - 1 : k;
      out += s.slice(i, end + 1);
      i = end;
    } else if (ch === '[') {
      const j = s.indexOf(']', i);
      out += odfRef(s.slice(i + 1, j < 0 ? s.length : j));
      i = j < 0 ? s.length : j;
    } else if (ch === ';') out += ',';
    else out += ch;
  }
  return out;
}

function odfRef(r) {
  const one = (p) => {
    const m = /^(\$?(?:'(?:[^']|'')*'|[^.]*))?\.(\$?[A-Z]+\$?\d+|\$?[A-Z]+|\$?\d+)$/i.exec(p.trim());
    if (!m) return { sheet: null, cell: p.trim().replace(/^\./, '') };
    let sheet = m[1] ? m[1].replace(/^\$/, '') : null;
    if (sheet === '') sheet = null;
    return { sheet, cell: m[2] };
  };
  const [a, b] = r.split(':');
  const A = one(a);
  const B = b ? one(b) : null;
  const q = (sh) => (sh.startsWith("'") ? sh : /^[\p{L}_][\p{L}\p{N}_.]*$/u.test(sh) ? sh : `'${sh.replace(/'/g, "''")}'`);
  const pre = A.sheet ? `${q(A.sheet)}!` : '';
  if (!B) return `${pre}${A.cell}`;
  if (B.sheet && B.sheet !== A.sheet) return `${pre}${A.cell}:${q(B.sheet)}!${B.cell}`;
  return `${pre}${A.cell}:${B.cell}`;
}

/** A1 수식 본문 → ODF 수식 (of:=...) */
export function toOdfFormula(f) {
  const s = String(f).replace(/^=/, '');
  const REF = /((?:'(?:[^']|'')+'|[\p{L}_][\p{L}\p{N}_.]*)!)?(\$?[A-Z]{1,3}\$?\d+)(?::(\$?[A-Z]{1,3}\$?\d+))?(?![\p{L}\p{N}_(])/uy;
  let out = '';
  for (let i = 0; i < s.length;) {
    const ch = s[i];
    if (ch === '"') {
      let j = i + 1;
      for (;;) { const k = s.indexOf('"', j); if (k < 0) { j = s.length; break; } if (s[k + 1] === '"') { j = k + 2; continue; } j = k + 1; break; }
      out += s.slice(i, j);
      i = j;
      continue;
    }
    const prev = s[i - 1] ?? '';
    if (!/[\p{L}\p{N}_.$]/u.test(prev)) {
      REF.lastIndex = i;
      const m = REF.exec(s);
      if (m) {
        const sheet = m[1] ? m[1].slice(0, -1) : '';
        const sh = sheet ? `$${sheet}` : '';
        out += m[3] ? `[${sh}.${m[2]}:${sh}.${m[3]}]` : `[${sh}.${m[2]}]`;
        i = REF.lastIndex;
        continue;
      }
    }
    out += ch === ',' ? ';' : ch;
    i++;
  }
  return `of:=${out}`;
}

// ─────────────── 읽기 ───────────────
/** .ods 바이트(또는 .fods 글자) → { data: { sheets }, warnings } */
export function readOds(input) {
  let xml;
  let stylesXml = null, settingsXml = null;
  if (typeof input === 'string') xml = input;
  else {
    const files = unzip(input instanceof Uint8Array ? input : new Uint8Array(input));
    xml = textOf(files['content.xml']);
    stylesXml = textOf(files['styles.xml']);
    settingsXml = textOf(files['settings.xml']);
    if (!xml) throw new Error('OpenDocument 스프레드시트(.ods)가 아닙니다');
  }
  const root = parseXml(xml);
  const styleRoots = [root, stylesXml ? parseXml(stylesXml) : null].filter(Boolean);
  // 스타일: 열 너비 · 행 높이 · 셀 서식 · 숫자 형식
  const colW = {}; const rowH = {}; const cellSt = {}; const numSt = {}; const tableHidden = {};
  for (const sr of styleRoots) {
    for (const holder of [child(sr, 'automatic-styles'), child(sr, 'styles')]) {
      for (const st of holder?.children ?? []) {
        const name = st.attrs['style:name'];
        if (!name) continue;
        if (st.name === 'number-style' || st.name === 'percentage-style' || st.name === 'currency-style' || st.name === 'date-style' || st.name === 'time-style') {
          const num = child(st, 'number');
          const dec = num ? Number(num.attrs['number:decimal-places'] ?? 0) : 0;
          const grp = num?.attrs['number:grouping'] === 'true';
          numSt[name] = st.name === 'percentage-style' ? { numFmt: 'percent', decimals: dec }
            : st.name === 'currency-style' ? { numFmt: 'currency', decimals: dec }
              : st.name === 'date-style' ? { numFmt: 'date' } : st.name === 'time-style' ? { numFmt: 'time' }
                : grp ? { numFmt: 'comma', decimals: dec } : dec ? { numFmt: 'number', decimals: dec } : null;
          continue;
        }
        if (st.name !== 'style') continue;
        const fam = st.attrs['style:family'];
        if (fam === 'table') { tableHidden[name] = child(st, 'table-properties')?.attrs['table:display'] === 'false'; } else if (fam === 'table-column') { const w = toPx(child(st, 'table-column-properties')?.attrs['style:column-width']); if (w) colW[name] = w; } else if (fam === 'table-row') { const h = toPx(child(st, 'table-row-properties')?.attrs['style:row-height']); if (h) rowH[name] = h; } else if (fam === 'table-cell') {
          const s = {};
          const tp = child(st, 'text-properties')?.attrs ?? {};
          const cp = child(st, 'table-cell-properties')?.attrs ?? {};
          const pp = child(st, 'paragraph-properties')?.attrs ?? {};
          if (tp['fo:font-weight'] === 'bold') s.bold = true;
          if (tp['fo:font-style'] === 'italic') s.italic = true;
          if (tp['style:text-underline-style'] && tp['style:text-underline-style'] !== 'none') s.underline = true;
          if (/^#[0-9a-f]{6}$/i.test(tp['fo:color'] ?? '')) s.color = tp['fo:color'].toLowerCase();
          if (tp['fo:font-size']) { const pt = parseFloat(tp['fo:font-size']); if (pt) s.size = pt; }
          if (tp['style:font-name']) s.font = tp['style:font-name'];
          if (/^#[0-9a-f]{6}$/i.test(cp['fo:background-color'] ?? '')) s.fill = cp['fo:background-color'].toLowerCase();
          if (cp['fo:wrap-option'] === 'wrap') s.wrap = true;
          const va = cp['style:vertical-align'];
          if (va === 'middle') s.valign = 'middle'; else if (va === 'top') s.valign = 'top';
          const ta = pp['fo:text-align'];
          if (ta === 'center') s.align = 'center'; else if (ta === 'end' || ta === 'right') s.align = 'right'; else if (ta === 'start' || ta === 'left') s.align = 'left';
          const border = cp['fo:border'];
          if (border && border !== 'none') { s.bt = s.bb = s.bl = s.br = 'thin'; }
          const ds = st.attrs['style:data-style-name'];
          cellSt[name] = { s, ds, parent: st.attrs['style:parent-style-name'] };
        }
      }
    }
  }
  const styleOf = (name) => {
    const c = cellSt[name];
    if (!c) return null;
    const base = c.parent ? styleOf(c.parent) ?? {} : {};
    const n = c.ds ? numSt[c.ds] : null;
    const st = { ...base, ...c.s, ...(n ?? {}) };
    return Object.keys(st).length ? st : null;
  };
  const body = child(child(root, 'body'), 'spreadsheet');
  const sheets = [];
  const warnings = [];
  const found = new Set();
  const scan = node => {
    if (!node) return;
    const a = node.attrs ?? {};
    if (['true', '1'].includes(a['table:protected']) || ['true', '1'].includes(a['table:structure-protected'])) throw new Error('보호된 OpenDocument의 보호 규칙은 아직 지원하지 않습니다. 보호가 해제된 사본 또는 Excel .xlsx 파일을 여세요.');
    if (node.name === 'content-validation' || a['table:content-validation-name']) found.add('데이터 유효성');
    if (node.name === 'calculation-settings') found.add('계산 옵션');
    if (node.name === 'page-layout' || a['table:print-ranges'] || node.name === 'table-header-rows' || node.name === 'table-header-columns') found.add('인쇄 및 반복 제목');
    if (node.name === 'database-range' || node.name === 'filter') found.add('필터');
    if (node.name === 'config-item' && /^(ShowGrid|GridColor|ShowZeroValues|ShowFormulas|HasColumnRowHeaders|ZoomValue|HorizontalSplit|VerticalSplit|HorizontalSplitPosition|VerticalSplitPosition|PositionLeft|PositionTop)$/.test(a['config:name'] ?? '')) found.add('보기 및 틀 고정');
    for (const sub of node.children ?? []) scan(sub);
  };
  for (const sr of [...styleRoots, settingsXml ? parseXml(settingsXml) : null]) scan(sr);
  if (found.size) warnings.push(`OpenDocument의 ${[...found].join(' · ')} 설정은 현재 가져오지 않습니다. 해당 설정을 확인하고 필요하면 .xlsx 원본을 사용하세요.`);
  for (const t of kids(body, 'table')) {
    const sheet = { name: (t.attrs['table:name'] ?? `Sheet${sheets.length + 1}`).slice(0, 31), cells: new CellMap(), colWidths: {}, rowHeights: {}, merges: [], hiddenRows: {}, hiddenCols: {} };
    if (t.attrs['table:display'] === 'false' || tableHidden[t.attrs['table:style-name']]) sheet.state = 'hidden';
    const colDefault = [];
    let c = 0;
    const cols = [];
    const collect = (el) => { for (const k of el.children) { if (k.name === 'table-column') cols.push(k); else if (k.name === 'table-column-group' || k.name === 'table-header-columns' || k.name === 'table-columns') collect(k); } };
    collect(t);
    for (const col of cols) {
      const n = Math.min(Number(col.attrs['table:number-columns-repeated'] ?? 1), 16384 - c);
      const w = colW[col.attrs['table:style-name']];
      const ds = col.attrs['table:default-cell-style-name'];
      for (let i = 0; i < n && c < 16384; i++, c++) {
        if (w && n < 1000) sheet.colWidths[c] = w;
        if (col.attrs['table:visibility'] === 'collapse') sheet.hiddenCols[c] = true;
        colDefault[c] = ds;
      }
    }
    let r = 0, hiddenRowCount = 0;
    const rows = [];
    const collectRows = (el) => { for (const k of el.children) { if (k.name === 'table-row') rows.push(k); else if (k.name === 'table-row-group' || k.name === 'table-header-rows' || k.name === 'table-rows') collectRows(k); } };
    collectRows(t);
    for (const row of rows) {
      const rep = Number(row.attrs['table:number-rows-repeated'] ?? 1);
      if (!Number.isInteger(rep) || rep < 1) throw new Error('OpenDocument 반복 행 개수가 올바르지 않습니다.');
      if (row.attrs['table:visibility'] === 'collapse') {
        hiddenRowCount += Math.min(rep, Math.max(0, 1048576 - r));
        if (hiddenRowCount > 100000) throw new Error('OpenDocument 숨김 행이 100,000개를 넘습니다. 대규모 숨김 구간은 아직 지원하지 않으므로 .xlsx 형식으로 여세요.');
      }
      const cellsIn = row.children.filter((k) => k.name === 'table-cell' || k.name === 'covered-table-cell');
      const hasContent = cellsIn.some((k) => k.children.length || k.attrs['office:value-type'] || k.attrs['table:formula'] || k.attrs['table:number-columns-spanned'] || k.attrs['table:number-rows-spanned'] || k.attrs['table:style-name']);
      if (!hasContent) {
        const h = rowH[row.attrs['table:style-name']];
        if (h && rep < 1000) for (let i = 0; i < rep; i++) sheet.rowHeights[r + i] = h;
        if (row.attrs['table:visibility'] === 'collapse') for (let i = 0; i < Math.min(rep, 1048576 - r); i++) sheet.hiddenRows[r + i] = true;
        r += rep;
        continue;
      }
      for (let k = 0; k < Math.min(rep, 100000); k++, r++) {
        const h = rowH[row.attrs['table:style-name']];
        if (h) sheet.rowHeights[r] = h;
        if (row.attrs['table:visibility'] === 'collapse') sheet.hiddenRows[r] = true;
        let cc = 0;
        for (const cell of cellsIn) {
          const n = Number(cell.attrs['table:number-columns-repeated'] ?? 1);
          const covered = cell.name === 'covered-table-cell';
          const type = cell.attrs['office:value-type'];
          const f = cell.attrs['table:formula'];
          const text = kids(cell, 'p').map((p) => pText(p)).join('\n');
          const st = styleOf(cell.attrs['table:style-name'] ?? colDefault[cc]);
          const cs = Number(cell.attrs['table:number-columns-spanned'] ?? 1);
          const rs = Number(cell.attrs['table:number-rows-spanned'] ?? 1);
          if (!covered && (cs > 1 || rs > 1)) sheet.merges.push({ r1: r, c1: cc, r2: r + rs - 1, c2: cc + cs - 1 });
          const hasVal = !covered && (type || f || text);
          if (hasVal || st) {
            for (let i = 0; i < Math.min(n, hasVal ? n : 256); i++) {
              let raw = '';
              if (!covered) {
                if (f) raw = `=${fromOdfFormula(f)}`;
                else if (type === 'float' || type === 'percentage' || type === 'currency') raw = String(Number(cell.attrs['office:value']));
                else if (type === 'date') { const d = cell.attrs['office:date-value'] ?? ''; raw = d.length > 10 ? d.replace('T', ' ').slice(0, 19) : d; } else if (type === 'time') raw = timeText(cell.attrs['office:time-value']);
                else if (type === 'boolean') raw = cell.attrs['office:boolean-value'] === 'true' ? 'TRUE' : 'FALSE';
                else if (text) raw = /^[=+\-@]/.test(text) || /^[\d.,\s%-]+$/.test(text) ? `'${text}` : text;
              }
              const style = st ? { ...st, ...(type === 'percentage' && !st.numFmt ? { numFmt: 'percent' } : {}), ...(type === 'date' && !st.numFmt ? { numFmt: 'date' } : {}) } : type === 'percentage' ? { numFmt: 'percent' } : type === 'date' ? { numFmt: 'date' } : undefined;
              if (raw || style) sheet.cells.setRC(r, cc + i, { raw, ...(style ? { style } : {}) });
            }
          }
          cc += n;
          if (cc > 16384) break;
        }
      }
      if (r > 1048576 * 20) break;
    }
    sheets.push(sheet);
  }
  if (!sheets.length) throw new Error('시트가 없습니다');
  return { data: { sheets }, warnings };
}

function pText(p) {
  let s = p.text ?? '';
  for (const k of p.children) {
    if (k.name === 's') s += ' '.repeat(Number(k.attrs['text:c'] ?? 1));
    else if (k.name === 'tab') s += '\t';
    else if (k.name === 'line-break') s += '\n';
    else s += pText(k);
  }
  return s;
}

function timeText(v) {
  const m = /^PT(\d+)H(\d+)M([\d.]+)S$/.exec(v ?? '');
  return m ? `${m[1].padStart(2, '0')}:${m[2].padStart(2, '0')}:${String(Math.floor(Number(m[3]))).padStart(2, '0')}` : '';
}

// ─────────────── 쓰기 ───────────────
const NS = 'xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0" xmlns:number="urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0" xmlns:of="urn:oasis:names:tc:opendocument:xmlns:of:1.2"';

/**
 * 통합 문서 → .ods 바이트
 * io: { raw(si,r,c), value(si,r,c), style(si,r,c), used(si) → {rows, cols}, colWidth(si,c), rowHeight(si,r) }
 */
export function writeOds(sheetsInfo, io) {
  const cellStyles = new Map();
  const numStyles = new Map();
  const colStyles = new Map();
  const rowStyles = new Map();
  const numStyleFor = (st) => {
    if (!st?.numFmt || st.numFmt === 'general' || st.numFmt === 'text') return null;
    const dec = st.decimals ?? (st.numFmt === 'number' ? 2 : 0);
    const key = `${st.numFmt}:${dec}`;
    if (numStyles.has(key)) return numStyles.get(key).name;
    const name = `N${numStyles.size + 1}`;
    let xml;
    if (st.numFmt === 'percent') xml = `<number:percentage-style style:name="${name}"><number:number number:decimal-places="${dec}" number:min-integer-digits="1"/><number:text>%</number:text></number:percentage-style>`;
    else if (/date/.test(st.numFmt)) xml = `<number:date-style style:name="${name}"><number:year number:style="long"/><number:text>-</number:text><number:month number:style="long"/><number:text>-</number:text><number:day number:style="long"/></number:date-style>`;
    else if (st.numFmt === 'time') xml = `<number:time-style style:name="${name}"><number:hours number:style="long"/><number:text>:</number:text><number:minutes number:style="long"/><number:text>:</number:text><number:seconds number:style="long"/></number:time-style>`;
    else xml = `<number:number-style style:name="${name}"><number:number number:decimal-places="${dec}" number:min-integer-digits="1"${st.numFmt === 'number' && !st.comma ? '' : ' number:grouping="true"'}/></number:number-style>`;
    numStyles.set(key, { name, xml });
    return name;
  };
  const cellStyleFor = (st) => {
    if (!st) return null;
    const ds = numStyleFor(st);
    const tp = [st.bold ? 'fo:font-weight="bold"' : '', st.italic ? 'fo:font-style="italic"' : '', st.underline ? 'style:text-underline-style="solid" style:text-underline-width="auto" style:text-underline-color="font-color"' : '',
      /^#[0-9a-f]{6}$/i.test(st.color ?? '') ? `fo:color="${st.color}"` : '', st.size ? `fo:font-size="${st.size}pt"` : '', st.font ? `style:font-name="${esc(st.font)}"` : ''].filter(Boolean).join(' ');
    const bd = (k) => (st[k] ? (st[k] === 'double' ? '1.1pt double #000000' : /medium|thick/.test(st[k]) ? '1.5pt solid #000000' : '0.5pt solid #000000') : 'none');
    const hasB = st.bt || st.bb || st.bl || st.br;
    const cp = [/^#[0-9a-f]{6}$/i.test(st.fill ?? '') ? `fo:background-color="${st.fill}"` : '', st.wrap ? 'fo:wrap-option="wrap"' : '', st.valign === 'middle' ? 'style:vertical-align="middle"' : st.valign === 'top' ? 'style:vertical-align="top"' : '',
      hasB ? `fo:border-top="${bd('bt')}" fo:border-bottom="${bd('bb')}" fo:border-left="${bd('bl')}" fo:border-right="${bd('br')}"` : ''].filter(Boolean).join(' ');
    const pp = st.align === 'center' ? 'fo:text-align="center"' : st.align === 'right' ? 'fo:text-align="end"' : st.align === 'left' ? 'fo:text-align="start"' : '';
    if (!tp && !cp && !pp && !ds) return null;
    const key = `${tp}|${cp}|${pp}|${ds}`;
    if (cellStyles.has(key)) return cellStyles.get(key).name;
    const name = `ce${cellStyles.size + 1}`;
    const xml = `<style:style style:name="${name}" style:family="table-cell" style:parent-style-name="Default"${ds ? ` style:data-style-name="${ds}"` : ''}>${cp ? `<style:table-cell-properties ${cp}/>` : ''}${pp ? `<style:paragraph-properties ${pp}/>` : ''}${tp ? `<style:text-properties ${tp}/>` : ''}</style:style>`;
    cellStyles.set(key, { name, xml });
    return name;
  };
  const colStyle = (px) => {
    const k = String(px);
    if (!colStyles.has(k)) colStyles.set(k, { name: `co${colStyles.size + 1}`, xml: `<style:style style:name="co${colStyles.size + 1}" style:family="table-column"><style:table-column-properties style:column-width="${pxToCm(px)}"/></style:style>` });
    return colStyles.get(k).name;
  };
  const rowStyle = (px) => {
    const k = String(px);
    if (!rowStyles.has(k)) rowStyles.set(k, { name: `ro${rowStyles.size + 1}`, xml: `<style:style style:name="ro${rowStyles.size + 1}" style:family="table-row"><style:table-row-properties style:row-height="${pxToCm(px)}" style:use-optimal-row-height="false"/></style:style>` });
    return rowStyles.get(k).name;
  };
  const tables = [];
  for (const sh of sheetsInfo) {
    const { si, name, merges = [], hidden, hiddenRows = {}, hiddenCols = {} } = sh;
    const used = io.used(si);
    let rowsN = used.rows, colsN = used.cols;
    for (const m of merges) { rowsN = Math.max(rowsN, m.r2 + 1); colsN = Math.max(colsN, m.c2 + 1); }
    const contentRowsN = Math.min(rowsN, 1048576);
    const hiddenRowIndices = Object.keys(hiddenRows).filter(r => hiddenRows[r] && Number.isInteger(Number(r)) && Number(r) >= 0 && Number(r) < 1048576).map(Number).sort((a,b) => a-b);
    for (const r of hiddenRowIndices) rowsN = Math.max(rowsN, r + 1);
    for (const c of Object.keys(hiddenCols)) if (hiddenCols[c]) colsN = Math.max(colsN, Number(c) + 1);
    rowsN = Math.min(rowsN, 1048576); colsN = Math.min(colsN, 16384);
    const mergeAt = new Map();
    const covered = new Set();
    for (const m of merges) {
      mergeAt.set(`${m.r1},${m.c1}`, m);
      for (let r = m.r1; r <= m.r2; r++) for (let c = m.c1; c <= m.c2; c++) if (r !== m.r1 || c !== m.c1) covered.add(`${r},${c}`);
    }
    const colsXml = [];
    for (let c = 0; c < Math.max(1, colsN); c++) colsXml.push(`<table:table-column table:style-name="${colStyle(io.colWidth(si, c))}"${hiddenCols[c] ? ' table:visibility="collapse"' : ''}/>`);
    const rowsXml = [];
    for (let r = 0; r < contentRowsN; r++) {
      const cells = [];
      let blank = 0;
      const flush = () => { if (blank) { cells.push(`<table:table-cell${blank > 1 ? ` table:number-columns-repeated="${blank}"` : ''}/>`); blank = 0; } };
      for (let c = 0; c < colsN; c++) {
        const key = `${r},${c}`;
        if (covered.has(key)) { flush(); cells.push('<table:covered-table-cell/>'); continue; }
        const raw = io.raw(si, r, c);
        const st = io.style(si, r, c);
        const stName = cellStyleFor(st);
        const m = mergeAt.get(key);
        if (!raw && !stName && !m) { blank++; continue; }
        flush();
        const a = [stName ? `table:style-name="${stName}"` : '', m ? `table:number-columns-spanned="${m.c2 - m.c1 + 1}" table:number-rows-spanned="${m.r2 - m.r1 + 1}"` : ''];
        const v = io.value(si, r, c);
        if (raw.startsWith('=')) a.push(`table:formula="${esc(toOdfFormula(raw))}"`);
        let body = '';
        if (typeof v === 'number') {
          if (/date/.test(st?.numFmt ?? '')) {
            const d = new Date(Math.round((v - 25569) * 86400000));
            a.push(`office:value-type="date" office:date-value="${d.toISOString().slice(0, v % 1 ? 19 : 10)}"`);
          } else a.push(`office:value-type="${st?.numFmt === 'percent' ? 'percentage' : 'float'}" office:value="${v}"`);
          body = `<text:p>${esc(io.display ? io.display(si, r, c) : String(v))}</text:p>`;
        } else if (typeof v === 'boolean') {
          a.push(`office:value-type="boolean" office:boolean-value="${v}"`);
          body = `<text:p>${v ? 'TRUE' : 'FALSE'}</text:p>`;
        } else if (v !== null && v !== undefined && v !== '') {
          const t = typeof v === 'object' ? String(v.error ?? v.code ?? '') : String(v);
          a.push('office:value-type="string"');
          body = t.split('\n').map((line) => `<text:p>${esc(line)}</text:p>`).join('');
        }
        cells.push(`<table:table-cell ${a.filter(Boolean).join(' ')}>${body}</table:table-cell>`);
      }
      const h = io.rowHeight(si, r);
      rowsXml.push(`<table:table-row${hiddenRows[r] ? ' table:visibility="collapse"' : ''}${h ? ` table:style-name="${rowStyle(h)}"` : ''}>${cells.join('') || '<table:table-cell/>'}</table:table-row>`);
    }
    // 사용 영역 뒤의 희소 숨김 행은 빈 간격을 반복행으로 기록한다. 100만 행을 셀 단위로 훑지 않는다.
    let tail = contentRowsN;
    const emptyRows = (start, count, hidden) => {
      if (!count) return;
      const h = io.rowHeight(si, start);
      rowsXml.push(`<table:table-row${count > 1 ? ` table:number-rows-repeated="${count}"` : ''}${hidden ? ' table:visibility="collapse"' : ''}${h ? ` table:style-name="${rowStyle(h)}"` : ''}><table:table-cell/></table:table-row>`);
    };
    for (const r of hiddenRowIndices) if (r >= contentRowsN) { emptyRows(tail, r - tail, false); emptyRows(r, 1, true); tail = r + 1; }
    if (!rowsXml.length) rowsXml.push('<table:table-row><table:table-cell/></table:table-row>');
    tables.push(`<table:table table:name="${esc(name)}"${hidden ? ' table:style-name="taHidden"' : ''}>${colsXml.join('')}${rowsXml.join('')}</table:table>`);
  }
  const auto = [...numStyles.values(), ...colStyles.values(), ...rowStyles.values(), ...cellStyles.values()].map((x) => x.xml).join('')
    + '<style:style style:name="taHidden" style:family="table"><style:table-properties table:display="false"/></style:style>';
  const content = `<?xml version="1.0" encoding="UTF-8"?>\n<office:document-content ${NS} office:version="1.3"><office:automatic-styles>${auto}</office:automatic-styles><office:body><office:spreadsheet>${tables.join('')}</office:spreadsheet></office:body></office:document-content>`;
  const styles = `<?xml version="1.0" encoding="UTF-8"?>\n<office:document-styles ${NS} office:version="1.3"><office:styles><style:style style:name="Default" style:family="table-cell"><style:text-properties style:font-name="맑은 고딕" fo:font-size="11pt"/></style:style></office:styles></office:document-styles>`;
  const manifest = '<?xml version="1.0" encoding="UTF-8"?>\n<manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0" manifest:version="1.3"><manifest:file-entry manifest:full-path="/" manifest:media-type="application/vnd.oasis.opendocument.spreadsheet"/><manifest:file-entry manifest:full-path="content.xml" manifest:media-type="text/xml"/><manifest:file-entry manifest:full-path="styles.xml" manifest:media-type="text/xml"/></manifest:manifest>';
  return zip({ mimetype: 'application/vnd.oasis.opendocument.spreadsheet', 'META-INF/manifest.xml': manifest, 'content.xml': content, 'styles.xml': styles });
}

