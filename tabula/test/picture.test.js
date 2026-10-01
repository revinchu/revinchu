import { test } from 'node:test';
import assert from 'node:assert/strict';
import { resizePicture, setPictureCrop, pictureCropStyle, pictureTransform, resetPictureFormatting } from '../src/picture.js';
import { Workbook } from '../src/workbook.js';
import { writeXlsx, readXlsx } from '../src/xlsx.js';
import { unzip, zip, textOf } from '../src/zip.js';
const src = 'data:image/svg+xml;base64,' + Buffer.from('<svg xmlns="http://www.w3.org/2000/svg" width="240" height="120"><rect width="120" height="120" fill="red"/><rect x="120" width="120" height="120" fill="blue"/></svg>').toString('base64');
const fixture = patch => new Workbook({ sheets: [{ name: '합성 그림', cells: {}, images: [{ id: 'im', src, name: '그림', x: 40, y: 60, w: 240, h: 120, ...patch }] }] });
test('그림 비율 잠금은 양쪽 크기 한도를 함께 지키고 원본을 변경하지 않는다', () => {
  const p = { w: 240, h: 120 };
  assert.deepEqual(resizePicture(p, 'w', 360), { w: 360, h: 180 });
  assert.deepEqual(resizePicture(p, 'h', 60), { w: 120, h: 60 });
  assert.deepEqual(resizePicture(p, 'w', 360, false), { w: 360, h: 120 });
  assert.deepEqual(resizePicture({ w: 4, h: 20000 }, 'w', 20000), { w: 4, h: 20000 });
  assert.deepEqual(resizePicture(p, 'h', 1), { w: 8, h: 4 });
  assert.deepEqual(p, { w: 240, h: 120 });
});
test('반대편 자르기까지 고려하여 100% 이상이 되지 않고 미리보기에 무한수가 없다', () => {
  const crop = { l: .4, r: .3 };
  assert.deepEqual(setPictureCrop(crop, 'l', 90), { l: .69, r: .3 });
  assert.deepEqual(crop, { l: .4, r: .3 });
  assert.deepEqual(pictureCropStyle({ l: .25, r: .25 }), { left: '-50%', top: '0%', width: '200%', height: '100%' });
  assert.equal(JSON.stringify(pictureCropStyle({ l: 1, r: 1, t: Infinity })).includes('Infinity'), false);
  assert.deepEqual(pictureCropStyle({ l: -.25, r: -.25 }), { left: '16.667%', top: '0%', width: '66.667%', height: '100%' });
});
test('회전·대칭 렌더와 서식 초기화는 내용·크기·대체텍스트를 변경하지 않는다', () => {
  assert.equal(pictureTransform({ rot: 30, flip: true, flipV: true }), 'rotate(30deg) scale(-1,-1)');
  assert.equal(pictureTransform({ rot: NaN }), '');
  const p = { src, name: '원본', alt: '빨강과 파랑', x: 20, y: 10, w: 240, h: 120, rot: 45, crop: { l: .2 }, border: '#000000' };
  const next = { ...p, ...resetPictureFormatting() };
  for (const k of ['src', 'name', 'alt', 'x', 'y', 'w', 'h']) assert.equal(next[k], p[k]);
  assert.equal(next.rot, undefined); assert.equal(next.crop, undefined); assert.equal(next.border, undefined);
});
test('그림 회전·대칭·자르기·설명·비율 잠금·배치가 표준 XLSX로 왕복한다', () => {
  const p = { rot: 22.5, flip: true, flipV: true, crop: { l: .1, r: .2, t: .05, b: .15 }, alt: '색 "빨강" & 파랑 <미리보기>', lockAspect: false, placement: 'absolute' };
  const bytes = writeXlsx(fixture(p)), back = readXlsx(bytes).data.sheets[0].images[0], xml = textOf(unzip(bytes)['xl/drawings/drawing1.xml']);
  for (const k of Object.keys(p)) assert.deepEqual(back[k], p[k], k);
  assert.match(xml, /noChangeAspect="0"/); assert.match(xml, /rot="1350000"/); assert.match(xml, /descr="색 &quot;빨강&quot; &amp; 파랑 &lt;미리보기&gt;"/);
});
test('새 그림 비율 잠금 기본값과 외부 XLSX의 잠금 생략·true 값을 구분한다', () => {
  const bytes = writeXlsx(fixture({})); assert.equal(readXlsx(bytes).data.sheets[0].images[0].lockAspect, true);
  const files = unzip(bytes), key = 'xl/drawings/drawing1.xml', source = textOf(files[key]);
  files[key] = source.replace('noChangeAspect="1"', ''); assert.equal(readXlsx(zip(files)).data.sheets[0].images[0].lockAspect, false);
  files[key] = source.replace('noChangeAspect="1"', 'noChangeAspect="true"').replace('<a:xfrm', '<a:xfrm flipH="true" flipV="true"');
  const pic = readXlsx(zip(files)).data.sheets[0].images[0]; assert.equal(pic.lockAspect, true); assert.equal(pic.flip, true); assert.equal(pic.flipV, true);
});
test('한 번의 그림 속성 변경은 실행 취소·다시 실행으로 온전히 복원된다', () => {
  const wb = fixture({ crop: { l: .2 }, alt: '원래 설명' }), before = structuredClone(wb.sheets[0].images);
  wb.transact(() => wb.setSheetProp(0, 'images', wb.sheets[0].images.map(p => ({ ...p, ...resetPictureFormatting(), ...resizePicture(p, 'w', 480), alt: '변경 설명', lockAspect: false }))));
  const after = structuredClone(wb.sheets[0].images); wb.undo(); assert.deepEqual(wb.sheets[0].images, before); wb.redo(); assert.deepEqual(wb.sheets[0].images, after);
});
