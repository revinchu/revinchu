import { Workbook } from './workbook.js';
import { writeXlsx, readChartTemplatePackage } from './xlsx.js';
import { zip, unzip } from './zip.js';
import { parseXml, descendants } from './xml.js';
import { chartTemplateFormat } from './chart-context.js';

const dec = new TextDecoder();
const esc = s => String(s).replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' })[c]);
// Excel과 같은 OPC 차트 서식 파일. 실제 셀 데이터 대신 예시 계열만 넣는다.
export function writeChartTemplate(chart, { seriesCount = 1, theme, themeFonts } = {}) {
  const count = Math.max(1, Math.min(1000, Number(seriesCount) || 1));
  const format = chartTemplateFormat(chart);
  const sample = { ...format, id: 'template', x: 0, y: 0, w: 480, h: 288,
    series: Array.from({ length: count }, (_, i) => ({ name: `계열 ${i + 1}`, cache: [i + 1, i + 3, i + 2], catCache: ['항목 1', '항목 2', '항목 3'], xCache: [1, 2, 3], sizeCache: [1, 2, 3] })) };
  const book = new Workbook({ theme, themeFonts, sheets: [{ name: '서식 예제', cells: {}, charts: [sample] }] });
  const original = unzip(writeXlsx(book)), files = {};
  for (const [path, bytes] of Object.entries(original)) if (/^xl\/(charts|media|theme)\//.test(path)) files[path] = bytes;
  const overrides = descendants(parseXml(dec.decode(original['[Content_Types].xml'])), 'Override').filter(x => files[x.attrs.PartName.replace(/^\//, '')]);
  const defaults = descendants(parseXml(dec.decode(original['[Content_Types].xml'])), 'Default');
  files['[Content_Types].xml'] = `<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">${defaults.map(x => `<Default Extension="${esc(x.attrs.Extension)}" ContentType="${esc(x.attrs.ContentType)}"/>`).join('')}${overrides.map(x => `<Override PartName="${esc(x.attrs.PartName)}" ContentType="${esc(x.attrs.ContentType)}"/>`).join('')}</Types>`;
  files['_rels/.rels'] = '<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/charts/chart1.xml"/></Relationships>';
  const relPath = 'xl/charts/_rels/chart1.xml.rels';
  const themeRel = '<Relationship Id="rIdTemplateTheme" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="../theme/theme1.xml"/>';
  files[relPath] = original[relPath] ? dec.decode(original[relPath]).replace('</Relationships>', themeRel + '</Relationships>') : `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${themeRel}</Relationships>`;
  return zip(files);
}
export function readChartTemplate(bytes) {
  if (bytes.byteLength > 30 * 1024 * 1024) throw new Error('차트 서식 파일은 30MB 이하만 지원합니다.');
  return chartTemplateFormat(readChartTemplatePackage(bytes));
}
