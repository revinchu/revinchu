import test from 'node:test';
import assert from 'node:assert/strict';
import { PICTURE_STYLES, pictureStylePatch, pictureCropPreset, dragPictureCrop, nudgePictureCrop, pictureCm, picturePx, pictureSourcePatch, resetPictureSource, resetPictureFormatting } from '../src/picture.js';
import { normalizePictureVisual, pictureVisual, needsPictureBake } from '../src/picture-filter.js';
import { removePictureBackground, pictureCompressionSize, pictureDataBytes } from '../src/picture-pixels.js';

test('24 actual picture styles are unique, cloned, and preserve content/geometry', () => {
  assert.equal(PICTURE_STYLES.length, 24); assert.equal(new Set(PICTURE_STYLES.map(s => s.id)).size, 24);
  const pic = { src: 'original', x: 3, y: 4, w: 30, h: 20, correction: { brightness: .2 } }, saved = structuredClone(pic);
  for (const style of PICTURE_STYLES) { const patch = pictureStylePatch(style.id, pic); assert.equal(patch.src, undefined); assert.equal(patch.w, undefined); assert.equal(patch.correction, undefined); assert.ok(!patch.radius || patch.radius <= 10); if (patch.shadow) patch.shadow.dx = 999; }
  assert.deepEqual(pic, saved); assert.notEqual(pictureStylePatch('soft-shadow').shadow.dx, 999);
  assert.throws(() => pictureStylePatch('missing'), /찾을/);
});
test('crop ratios and handles use original dimensions and never mutate input', () => {
  const pic = { w: 240, h: 120 }, square = pictureCropPreset(pic, 1, { w: 1000, h: 500 });
  assert.deepEqual(square, { crop: { l: .25, r: .25, t: 0, b: 0 }, w: 240, h: 240 });
  const wide = pictureCropPreset(pic, 16 / 9, { w: 1000, h: 1000 }); assert.equal(wide.w / wide.h, 16 / 9);
  const before = { l: .1, r: .2, t: .1, b: .1 }, after = dragPictureCrop(before, 'lt', .1, .2);
  assert.ok(Math.abs(after.l - .17) < 1e-10); assert.ok(Math.abs(after.t - .26) < 1e-10); assert.deepEqual(before, { l: .1, r: .2, t: .1, b: .1 });
  assert.ok(dragPictureCrop(before, 'r', -2, 0).r + before.l <= .99 + 1e-9);
  assert.equal(pictureCm(96), 2.54); assert.equal(picturePx(2.54), 96);
});
test('crop keyboard nudges use original percentage points after cropping and preserve limits', () => {
  const original = { l: .2, r: .1, t: .15, b: .05 }, saved = { ...original };
  const one = nudgePictureCrop(original, 'r', -.01, 0), five = nudgePictureCrop(one, 'r', -.05, 0);
  assert.ok(Math.abs(one.r - .11) < 1e-10); assert.ok(Math.abs(five.r - .16) < 1e-10);
  assert.equal(five.l, .2); assert.equal(five.t, .15); assert.equal(five.b, .05);
  const corner = nudgePictureCrop(original, 'lt', .01, -.05);
  assert.ok(Math.abs(corner.l - .21) < 1e-10); assert.ok(Math.abs(corner.t - .1) < 1e-10);
  assert.ok(Math.abs(nudgePictureCrop(original, 'b', 0, -.01).b - .06) < 1e-10);
  assert.ok(Math.abs(nudgePictureCrop(original, 'r', .01, 0).r - .09) < 1e-10);
  const limited = nudgePictureCrop({ l: .95, r: .03 }, 'r', -.05, 0);
  assert.ok(limited.l + limited.r <= .99 + 1e-10);
  assert.deepEqual(original, saved);
  assert.ok(Math.abs(dragPictureCrop(original, 'r', -.01, 0).r - .107) < 1e-10);
});
test('raster edit keeps first source and vector data, and explicit restore is lossless', () => {
  const p = { src: 'svg', emf: 'vector', png: 'fallback', w: 240, h: 120 };
  const edited = { ...p, ...pictureSourcePatch(p, 'raster1', { w: 1000, h: 500 }) };
  assert.equal(edited.emf, undefined); assert.equal(edited.originalEmf, 'vector'); assert.equal(edited.originalWidth, 1000);
  const twice = { ...edited, ...pictureSourcePatch(edited, 'raster2', { w: 400, h: 200 }) };
  assert.equal(twice.originalSrc, 'svg'); assert.equal(twice.originalWidth, 1000);
  const restored = { ...twice, ...resetPictureSource(twice) }; assert.equal(restored.src, p.src); assert.equal(restored.emf, p.emf); assert.equal(restored.png, p.png); assert.equal(restored.w, 240);
  assert.deepEqual(resetPictureSource(twice, true).w, 1000);
  assert.equal(resetPictureFormatting().correction, undefined); assert.equal(resetPictureFormatting().src, undefined);
});
test('visual normalization guards injection and distinguishes native vs baked corrections', () => {
  const v = normalizePictureVisual({ correction: { brightness: 99, contrast: -99, sharpness: 99 }, color: { saturation: -3, temperature: Infinity }, glow: { color: 'red"/><script>', size: 900 }, reflection: { size: 20 } });
  assert.deepEqual(v.correction, { brightness: 1, contrast: -1, sharpness: 1 }); assert.equal(v.color.saturation, 0); assert.equal(v.color.temperature, 0); assert.equal(v.glow.color, '#5b9bd5'); assert.equal(v.glow.size, 100); assert.equal(v.reflection.size, 1);
  assert.equal(needsPictureBake({ correction: { brightness: .2 }, color: { recolor: 'grayscale' } }), false);
  assert.equal(needsPictureBake({ color: { saturation: .8 } }), true); assert.equal(needsPictureBake({ artistic: { type: 'pencil' } }), true);
  assert.equal(pictureVisual({}, 'x').defs, '');
  const svg = pictureVisual({ shadow: true, artistic: { type: 'posterize' } }, '"><script>'); assert.doesNotMatch(svg.defs, /<script/); assert.match(svg.defs, /feDropShadow/); assert.match(svg.defs, /type="discrete"/);
  const duo = pictureVisual({ color: { duotone: ['#123456', '#abcdef'] } }, 'duo'); assert.match(duo.defs, /feColorMatrix/); assert.deepEqual(duo.effects.color.duotone, ['#123456', '#abcdef']);
});
const image = () => { const width = 7, height = 7, data = new Uint8ClampedArray(width * height * 4); for (let y = 0; y < height; y++) for (let x = 0; x < width; x++) { const i = (y * width + x) * 4, white = x === 0 || x === 6 || y === 0 || y === 6 || x === 3 && y === 3; data[i] = data[i + 1] = data[i + 2] = white ? 255 : 0; data[i + 3] = 255; } return { width, height, data }; };
test('contiguous background removal does not delete enclosed same-color foreground', () => {
  const original = image(), before = original.data.slice(), next = removePictureBackground(original, { tolerance: 0 });
  assert.equal(next.removed, 24); assert.equal(next.data[3], 0); assert.equal(next.data[(3 * 7 + 3) * 4 + 3], 255); assert.deepEqual(original.data, before);
  const seeded = removePictureBackground(original, { seed: { x: 3, y: 3 }, tolerance: 0 }); assert.equal(seeded.removed, 1); assert.equal(seeded.data[3], 255);
});
test('keep/remove marks restore original alpha or remove explicitly; pixel limit fails closed', () => {
  const original = image(), next = removePictureBackground(original, { marks: [{ x: 0, y: 0, radius: 1, mode: 'keep' }, { x: 3, y: 3, radius: 1, mode: 'remove' }] });
  assert.equal(next.data[3], 255); assert.equal(next.data[(3 * 7 + 3) * 4 + 3], 0);
  assert.throws(() => removePictureBackground({ width: 4000, height: 4000, data: [] }), /400만/);
  assert.throws(() => removePictureBackground(original, { marks: Array(2001).fill({}) }), /2,000/);
});
test('compression dimensions never upscale; estimates include base64 padding and unicode', () => {
  assert.deepEqual(pictureCompressionSize(4000, 2000, 1000, 1000), { w: 1000, h: 500 });
  assert.deepEqual(pictureCompressionSize(100, 50, 1000, 1000), { w: 100, h: 50 });
  assert.equal(pictureDataBytes('data:image/png;base64,YQ=='), 1); assert.equal(pictureDataBytes('data:image/svg+xml,' + encodeURIComponent('한글')), 6); assert.equal(pictureDataBytes('https://example.com/image.png'), null);
});
