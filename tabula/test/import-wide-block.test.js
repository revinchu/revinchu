import { test } from 'node:test';
import assert from 'node:assert/strict';
import { zip } from '../src/zip.js';
import { Workbook } from '../src/workbook.js';
import { readXlsx, readXlsxAsync, writeXlsx } from '../src/xlsx.js';
import { formatValue } from '../src/format.js';
import { toDelimited } from '../src/csv.js';

const main = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const rel = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const col = n => n < 26 ? String.fromCharCode(65 + n) : `A${String.fromCharCode(65 + n - 26)}`;
function fixture(body, dimension = 'A1:AF8192', annotations = '', columns = '') {
  return zip({
    '[Content_Types].xml': '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>',
    '_rels/.rels': `<Relationships><Relationship Id="r1" Type="${rel}/officeDocument" Target="xl/workbook.xml"/></Relationships>`,
    'xl/workbook.xml': `<workbook xmlns="${main}" xmlns:r="${rel}"><sheets><sheet name="Data" sheetId="1" r:id="r1"/></sheets></workbook>`,
    'xl/_rels/workbook.xml.rels': `<Relationships><Relationship Id="r1" Type="${rel}/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="r2" Type="${rel}/styles" Target="styles.xml"/></Relationships>`,
    'xl/styles.xml': `<styleSheet xmlns="${main}"><fonts count="2"><font><sz val="11"/><name val="Calibri"/></font><font><b/><sz val="11"/><name val="Calibri"/></font></fonts><fills count="3"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FFFFFF00"/></patternFill></fill><fill><patternFill patternType="solid"><fgColor rgb="FFFF0000"/></patternFill></fill></fills><borders count="1"><border/></borders><cellXfs count="5"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/><xf numFmtId="0" fontId="0" fillId="1" borderId="0"/><xf numFmtId="0" fontId="1" fillId="2" borderId="0"/><xf numFmtId="49" fontId="0" fillId="1" borderId="0"/><xf numFmtId="14" fontId="0" fillId="1" borderId="0"/></cellXfs></styleSheet>`,
    'xl/worksheets/sheet1.xml': `<worksheet xmlns="${main}" xmlns:r="${rel}">${dimension ? `<dimension ref="${dimension}"/>` : ''}<sheetFormatPr defaultRowHeight="15"/><cols><col min="2" max="2" width="22" customWidth="1"/>${columns}</cols><sheetData>${body}</sheetData>${annotations}</worksheet>`,
    'xl/worksheets/_rels/sheet1.xml.rels': `<Relationships><Relationship Id="link" Type="${rel}/hyperlink" Target="https://example.test/report" TargetMode="External"/><Relationship Id="note" Type="${rel}/comments" Target="../comments1.xml"/></Relationships>`,
    'xl/comments1.xml': `<comments xmlns="${main}"><authors><author>Audit</author></authors><commentList><comment ref="B3" authorId="0"><text><t>독립 메모😀</t></text></comment></commentList></comments>`,
  });
}

for (const [rows, cols, expected] of [[8191, 33, false], [8192, 31, false], [8192, 32, true], [50000, 1, false], [50001, 1, true]]) {
  test(`block eligibility uses both height and area: ${rows} x ${cols}`, () => {
    const last = `${col(cols - 1)}${rows}`;
    const body = `<row r="1"><c r="A1" t="inlineStr"><is><t>header</t></is></c></row><row r="2"><c r="A2"><v>1</v></c></row><row r="${rows}"><c r="${last}"><v>42</v></c></row>`;
    const wb = new Workbook(readXlsx(fixture(body, `A1:${last}`)).data);
    assert.equal(!!wb.sheets[0].blocks?.length, expected);
    assert.equal(wb.getValue(0, rows - 1, cols - 1), 42);
  });
}

test('wide-sheet blocks preserve types, local styles, formulas, links, notes, edits, undo and XLSX reopen', async () => {
  const rows = [];
  rows.push('<row r="1">' + Array.from({ length: 32 }, (_, c) => `<c r="${col(c)}1" t="inlineStr"><is><t>field ${c}</t></is></c>`).join('') + '</row>');
  for (let r = 1; r < 8192; r++) {
    const cells = [];
    for (let c = 0; c < 32; c++) {
      const ref = `${col(c)}${r + 1}`, style = c === 5 && r % 127 === 0 ? 2 : c === 6 || c === 7 ? 3 : c === 9 ? 4 : 1;
      let type = '', value = `<v>${c === 9 ? 46000 + r : r * 32 + c}</v>`;
      if (c === 1) { type = ' t="inlineStr"'; value = `<is><t>분류${r % 4}😀</t></is>`; }
      if (c === 2) { type = ' t="b"'; value = `<v>${r % 2}</v>`; }
      if (c === 3) { type = ' t="e"'; value = `<v>${r % 2 ? '#N/A' : '#DIV/0!'}</v>`; }
      if (c === 4) value = `<f>A${r + 1}+1</f><v>${r * 32 + 1}</v>`;
      if (c === 6) { type = ' t="inlineStr"'; value = `<is><t>000${r}</t></is>`; }
      if (c === 8 && r % 251 === 0) { type = ' t="inlineStr"'; value = '<is><t></t></is>'; }
      cells.push(`<c r="${ref}" s="${style}"${type}>${value}</c>`);
    }
    rows.push(`<row r="${r + 1}"${r === 1 ? ' ht="25" customHeight="1"' : ''}>${cells.join('')}</row>`);
  }
  const body = rows.join(''), annotations = '<hyperlinks><hyperlink ref="A3" r:id="link"/></hyperlinks>';
  // A dimension-less valid file exercises the old ordinary-cell path as the
  // reference; the exact same content with its real extent exercises blocks.
  const reference = new Workbook(readXlsx(fixture(body, null, annotations)).data);
  const diagnostics = [];
  const wb = new Workbook((await readXlsxAsync(fixture(body, undefined, annotations), null, { onDiagnostics: d => diagnostics.push(d) })).data);
  assert.equal(diagnostics.at(-1).blockMode, true);
  assert.equal(diagnostics.at(-1).plainFormulaCells, 8191);
  assert.ok(wb.sheets[0].blocks.length > 0);
  assert.ok(wb.sheets[0].cells.size < 26000, `remaining cell objects: ${wb.sheets[0].cells.size}`);
  const points = new Set([0, 1, 2, 127, 251, 8190, 8191]);
  for (let r = 1; r < 8192; r += 137) points.add(r);
  const compare = (left, right) => {
    for (const r of points) for (let c = 0; c < 32; c++) {
      assert.deepEqual(left.getValue(0, r, c), right.getValue(0, r, c), `value ${r},${c}`);
      // The editor and formula bar both use getRaw directly: its visible
      // input representation must stay identical, including @ numeric text.
      assert.equal(left.getRaw(0, r, c), right.getRaw(0, r, c), `input ${r},${c}`);
      const clean = style => Object.fromEntries(Object.entries(style).filter(([, v]) => v !== undefined));
      assert.deepEqual(clean(left.styleAt(0, r, c)), clean(right.styleAt(0, r, c)), `style ${r},${c}`);
      assert.deepEqual(formatValue(left.getValue(0, r, c), left.styleAt(0, r, c)), formatValue(right.getValue(0, r, c), right.styleAt(0, r, c)), `display ${r},${c}`);
    }
    const csv = book => toDelimited([...points].map(r => Array.from({ length: 32 }, (_, c) => formatValue(book.getValue(0, r, c), book.styleAt(0, r, c)).text)));
    assert.equal(csv(left), csv(right));
    assert.equal(left.getCell(0, 2, 0).link, 'https://example.test/report');
    assert.equal(left.getCell(0, 2, 1).comment, '독립 메모😀');
    assert.deepEqual(left.sheets[0].colWidths, right.sheets[0].colWidths);
    assert.deepEqual(left.sheets[0].rowHeights, right.sheets[0].rowHeights);
  };
  compare(wb, reference);
  const edit = book => book.transact(() => {
    book.setInput(0, 2, 0, '900');
    book.setInput(0, 127, 1, '편집😀');
    book.setStyle(0, 251, 10, { italic: true, fill: '#123456' });
  });
  edit(wb); edit(reference);
  assert.equal(wb.getValue(0, 2, 4), 901);
  compare(wb, reference);
  wb.undo(); reference.undo(); compare(wb, reference);
  wb.redo(); reference.redo(); compare(wb, reference);
  const restored = new Workbook(readXlsx(writeXlsx(wb)).data);
  compare(restored, reference);
  assert.equal(restored.getRaw(0, 2, 4), '=A3+1');
});


test('inflated sparse dimensions do not spread staged column styles into absent cells', () => {
  const body = '<row r="1"><c r="A1" t="inlineStr"><is><t>Header</t></is></c></row>'
    + [2, 3, 10, 30].map(r => `<row r="${r}"><c r="A${r}" s="1"><v>${r}</v></c><c r="B${r}" s="2" t="inlineStr"><is><t>item</t></is></c><c r="C${r}" s="1"><v>${r + 2}</v></c>${r === 2 ? '<c r="D2" s="2"><f>A2+1</f><v>3</v></c>' : ''}</row>`).join('')
    + '<row r="4"><c r="A4" s="1"/><c r="B4" s="2"/><c r="C4" s="1"/></row>'
    + '<row r="11" s="2" customFormat="1"><c r="A11" s="1"><v>11</v></c></row>'
    + '<row r="13"><c r="A13"><v>13</v></c></row>'
    + '<row r="8192"><c r="A8192" s="1"/><c r="B8192" s="2"/></row>';
  const annotations = '<hyperlinks><hyperlink ref="A3" r:id="link"/></hyperlinks>';
  const columns = '<col min="3" max="3" style="2"/>';
  const reference = new Workbook(readXlsx(fixture(body, null, annotations, columns)).data);
  const loaded = new Workbook(readXlsx(fixture(body, undefined, annotations, columns)).data);
  assert.equal(loaded.sheets[0].blocks.length, 0, 'a real gap requires the original sparse representation');
  assert.equal(loaded.sheets[0].cells.size, reference.sheets[0].cells.size);
  const compare = wb => {
    for (const r of [...Array(35).keys(), 8191]) for (let c = 0; c < 5; c++) {
      assert.deepEqual(wb.getValue(0, r, c), reference.getValue(0, r, c), `value ${r},${c}`);
      assert.equal(wb.getRaw(0, r, c), reference.getRaw(0, r, c), `input ${r},${c}`);
      assert.deepEqual(wb.styleAt(0, r, c), reference.styleAt(0, r, c), `style ${r},${c}`);
    }
    assert.equal(wb.getCell(0, 2, 0).link, 'https://example.test/report');
    assert.equal(wb.getCell(0, 2, 1).comment, '독립 메모😀');
  };
  compare(loaded);
  loaded.transact(() => loaded.setInput(0, 6, 1, 'new'));
  loaded.undo(); compare(loaded);
  compare(new Workbook(readXlsx(writeXlsx(loaded)).data));
});


for (const middleStyle of [0, 1]) test(`a later ordinary cell resets earlier block formatting (style ${middleStyle})`, () => {
  const body = '<row r="1"><c r="A1" t="inlineStr"><is><t>Header</t></is></c></row>'
    + `<row r="2"><c r="A2" s="2"><v>2</v></c></row><row r="3"><c r="A3" s="${middleStyle}"><v>3</v></c></row><row r="4"><c r="A4" s="2"><v>4</v></c></row>`;
  const reference = new Workbook(readXlsx(fixture(body, null)).data);
  const loaded = new Workbook(readXlsx(fixture(body)).data);
  // No own style needs fallback; a nonempty partial style replaces the block
  // style entirely in styleAt and can safely stay an ordinary-cell override.
  assert.equal(loaded.sheets[0].blocks.length, middleStyle ? 1 : 0);
  assert.equal(!!loaded.styleAt(0, 2, 0).bold, false);
  for (const wb of [loaded, new Workbook(readXlsx(writeXlsx(loaded)).data)]) {
    for (let r = 1; r < 4; r++) {
      assert.equal(wb.getValue(0, r, 0), reference.getValue(0, r, 0));
      assert.deepEqual(wb.styleAt(0, r, 0), reference.styleAt(0, r, 0));
    }
  }
});
