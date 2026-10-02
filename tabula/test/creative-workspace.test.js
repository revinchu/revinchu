import test from 'node:test';
import assert from 'node:assert/strict';
import { Workbook } from '../src/workbook.js';
import { alternatingRules, isBandingRule } from '../src/alternating-colors.js';
import { prepareCond, condFormatAt } from '../src/condfmt.js';
import { writeXlsx, readXlsx } from '../src/xlsx.js';
import { normalizeVideo, VIDEO_POSTER } from '../src/media-object.js';
test('교차색상은 값·기본서식 보존, 첫/끝행·홀짝행과 Excel 왕복 유지',()=>{
 const w=new Workbook();const s=w.sheets[0];w.setCellData(0,0,0,{raw:'보존',style:{bold:true}});
 s.cond=alternatingRules({r1:0,c1:0,r2:5,c2:2},{footer:true});
 const fills=book=>{const preps=prepareCond(book,0);return Array.from({length:6},(_,r)=>condFormatAt(preps,book,0,r,0,book.getValue(0,r,0)).style.fill);};
 assert.deepEqual(fills(w),['#185c45','#ffffff','#e8f3ed','#ffffff','#e8f3ed','#c9e2d3']);
 const next=new Workbook();next.restore(readXlsx(writeXlsx(w)).data);
 assert.deepEqual(fills(next),fills(w));assert.ok(next.sheets[0].cond.every(isBandingRule));assert.equal(next.getRaw(0,0,0),'보존');assert.equal(next.styleAt(0,0,0).bold,true);
 const one=alternatingRules({r1:2,c1:1,r2:2,c2:1},{footer:true});assert.equal(one.length,1);
});
test('영상 메타데이터는 실행 주소 거부, 사진·링크와 Excel 왕복 유지',()=>{
 for(const src of ['javascript:alert(1)','data:text/html,x','file:///a','https://user:pass@example.com/a'])assert.equal(normalizeVideo({kind:'video',src}),null);
 const w=new Workbook(),media=normalizeVideo({kind:'video',src:'https://example.com/movie.mp4',title:'합성 영상',source:'test'});
 w.sheets[0].images=[{id:'video1',src:VIDEO_POSTER,media,hyperlink:{target:media.src},x:10,y:10,w:480,h:270,name:'합성 영상'}];
 const next=readXlsx(writeXlsx(w)).data.sheets[0].images[0];assert.deepEqual(next.media,media);assert.equal(next.hyperlink.target,media.src);assert.ok(next.src.startsWith('data:image/png'));
});
