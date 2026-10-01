import { test } from 'node:test';
import assert from 'node:assert/strict';
import { pictureEffects, pictureShadowStyle } from '../src/picture.js';
import { Workbook } from '../src/workbook.js';
import { writeXlsx, readXlsx } from '../src/xlsx.js';
import { unzip, zip, textOf } from '../src/zip.js';
const src = 'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+/a9sAAAAASUVORK5CYII=';
const fixture = patch => new Workbook({ sheets: [{ name: '그림 서식', cells: {}, images: [{ id: 'im', name: '서식 검증', src, x: 40, y: 60, w: 240, h: 120, ...patch }] }] });
const close = (a,b) => assert.ok(Math.abs(a-b) < .001, `${a} != ${b}`);
test('그림 효과는 비정상 수치를 유한 범위로 제한하고 그림자 색을 검증한다', () => {
  assert.deepEqual(pictureEffects({ w: 240, h: 120, opacity: -2, radius: 200, shadow: { dx: Infinity, dy: -3000, blur: -1, color: 'red;display:none', opacity: 2 } }), { opacity: 0, radius: 60, shadow: { dx: 3, dy: -1000, blur: 0, color: '#000000', opacity: 1 } });
  assert.equal(pictureShadowStyle({ shadow: { dx: -4, dy: 5, blur: 6, opacity: .25, color: '#336699' } }, .5), '-2px 2.5px 3px rgba(51,102,153,0.25)');
});
test('그림 투명도·둥근 모서리·그림자와 소수 테두리를 표준 DrawingML로 왕복한다', () => {
  const p = { opacity: .35, radius: 8, shadow: { dx: -4, dy: 5, blur: 6, opacity: .25, color: '#336699' }, border: '#112233', borderW: .75 };
  const bytes = writeXlsx(fixture(p)), files = unzip(bytes), xml = textOf(files['xl/drawings/drawing1.xml']);
  assert.match(xml, /alphaModFix amt="35000"/); assert.match(xml, /prst="roundRect"/); assert.match(xml, /outerShdw/);
  // No WIXEL extension is needed to retain these effects in another OOXML editor.
  const back = readXlsx(bytes).data.sheets[0].images[0];
  close(back.opacity, p.opacity); close(back.radius, p.radius); close(back.borderW, p.borderW);
  for (const k of ['dx','dy','blur','opacity']) close(back.shadow[k], p.shadow[k]);
  assert.equal(back.shadow.color, p.shadow.color);
  for (let n=0;n<3;n++) { const again=readXlsx(writeXlsx(fixture(back))).data.sheets[0].images[0]; close(again.radius,p.radius); close(again.opacity,p.opacity); }
});
test('외부 roundRect의 기본 반경과 생략된 alphaModFix 값, 완전 투명도 0을 보존한다', () => {
  const files=unzip(writeXlsx(fixture({opacity:0,radius:20,shadow:true}))), key='xl/drawings/drawing1.xml';
  assert.equal(readXlsx(zip(files)).data.sheets[0].images[0].opacity,0);
  files[key]=textOf(files[key]).replace('amt="0"','').replace(/<a:gd name="adj"[^>]*\/>/,'');
  const back=readXlsx(zip(files)).data.sheets[0].images[0];
  assert.equal(back.opacity,1); close(back.radius,20.0004); assert.ok(back.shadow);
});

test('그림·차트 앵커 기준 행 높이는 Excel 글꼴 기본값으로 재계산되지 않게 저장한다', () => {
 const files=unzip(writeXlsx(fixture({}))), sheet=textOf(files['xl/worksheets/sheet1.xml']);
 assert.match(sheet, /<sheetFormatPr[^>]*defaultRowHeight="15"[^>]*customHeight="1"/);
});
