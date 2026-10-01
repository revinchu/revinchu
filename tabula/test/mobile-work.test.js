import { test } from 'node:test';
import assert from 'node:assert/strict';
import { mobileLayout, mobileSheetZoom } from '../src/mobile-work.js';
test('mobile auto detects phones, touch tablets and landscape without changing desktop', () => {
  for (const [width,height,coarse,active] of [[320,568,false,true],[390,844,true,true],[844,390,true,true],[768,1024,true,true],[1024,768,false,false],[1440,900,true,false]]) assert.equal(mobileLayout({width,height,coarse}).active, active);
});
test('explicit device preference wins over the viewport and keyboard', () => {
  assert.equal(mobileLayout({width:320,height:568,preference:'off'}).active,false);
  assert.equal(mobileLayout({width:1440,height:900,preference:'on'}).active,true);
  assert.equal(mobileLayout({width:390,height:360,layoutHeight:844,typing:true}).keyboard,true);
  assert.equal(mobileLayout({width:390,height:360,layoutHeight:844,typing:false}).keyboard,false);
  assert.equal(mobileLayout({width:390,height:844,layoutHeight:844,typing:true}).keyboard,false);
});
test('mobile fit keeps text readable and respects a smaller saved workbook zoom', () => {
  assert.equal(mobileSheetZoom(320),75);assert.equal(mobileSheetZoom(390),85);
  assert.equal(mobileSheetZoom(768),100);assert.equal(mobileSheetZoom(390,50),50);
  assert.equal(mobileSheetZoom(390,200),85);
});
