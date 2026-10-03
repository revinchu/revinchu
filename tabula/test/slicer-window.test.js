import test from 'node:test';
import assert from 'node:assert/strict';
import {slicerWindow} from '../src/slicer-window.js';
test('슬라이서 창은 첫/중간/마지막 항목과 다중열을 제한된 항목 수로 표시한다',()=>{
 for(const count of [0,5,501,50000,1000000])for(const columns of [1,3,12])for(const top of [0,1000,Infinity]){
  const w=slicerWindow(count,columns,24,3,240,top);assert.ok(w.first>=0);assert.ok(w.last<=count);assert.ok(w.last-w.first<=columns*18);assert.ok(w.physical<=8000000);
  if(top===Infinity)assert.equal(w.last,count);
 }
});
test('초대형 항목 목록도 끝부분 스크롤과 실제 단추 간격을 유지한다',()=>{
 const w=slicerWindow(1000000,1,24,3,240,Infinity);assert.equal(w.last,1000000);assert.equal(w.physical,8000000);assert.ok(w.top<8000000);assert.ok(w.top>7999000);assert.equal(w.pitch,27);
});
