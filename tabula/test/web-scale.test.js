import { test } from 'node:test';
import assert from 'node:assert/strict';
import { WEB, NET, sparklineSvg } from '../src/fx-web.js';
import { Range } from '../src/fxcore.js';

test('IMPORTDATA: 15만 행의 CSV도 오류 대신 범위로 가져옴', () => {
  const url = 'https://example.com/synthetic-large.csv';
  NET.cache.set(url, { state: 'ok', data: '1,2\n'.repeat(150000), sheets: new Set() });
  try {
    const result = WEB.IMPORTDATA([url], { here: { si: 0 } });
    assert.ok(result instanceof Range, result?.code);
    assert.equal(result.height, 150000);
    assert.equal(result.width, 2);
    assert.deepEqual(result.rows[149999], [1, 2]);
  } finally { NET.cache.delete(url); }
});

test('SPARKLINE: 큰 원본의 축 범위와 최고·최저 색상', () => {
  const svg = decodeURIComponent(sparklineSvg(Array.from({ length: 150000 }, (_, i) => i % 101)).split(',').slice(1).join(','));
  assert.match(svg, /<path/);
  assert.doesNotMatch(svg, /NaN|Infinity/);
  const columns = decodeURIComponent(sparklineSvg([5, -2, 9], { charttype: 'column', highcolor: '#00ff00', lowcolor: '#ff0000', ymin: -10, ymax: 20 }).split(',').slice(1).join(','));
  assert.match(columns, /fill="#00ff00"/);
  assert.match(columns, /fill="#ff0000"/);
});
