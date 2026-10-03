import {test} from 'node:test';
import assert from 'node:assert/strict';
import {THEME,DEFAULT_THEME,setThemeColors,withThemeColors} from '../src/stylepresets.js';
import {computePivot,resolvePivot} from '../src/pivot.js';
const newer=['FFFFFF','000000','E8E8E8','0E2841','156082','E97132','196B24','0F9ED5','A02B93','4EA72E','467886','96607D'];
function pagePaint(){
 const def={rows:[],cols:[],values:[{field:'Amount',agg:'sum'}],pages:['Region'],style:'PivotStyleLight16'};
 const resolved=resolvePivot([['Region','Amount'],['A',10]],def);
 return computePivot(resolved,resolved.def).grid[0][1].style.fill;
}
test('새 문서의 첫 피벗도 이전 문서가 아닌 저장된 테마로 칠한다',()=>{
 const previous=THEME.colors;
 try{
  setThemeColors(DEFAULT_THEME);assert.equal(pagePaint(),'#dae2f3');
  assert.equal(withThemeColors(newer,pagePaint),'#c0e6f5');
  assert.strictEqual(THEME.colors,DEFAULT_THEME);assert.equal(pagePaint(),'#dae2f3');
  setThemeColors(newer);assert.equal(pagePaint(),'#c0e6f5');
 }finally{setThemeColors(previous);}
});
test('준비 실패와 중첩/늦은 준비도 현재 문서의 테마를 덮지 않는다',()=>{
 const previous=THEME.colors;
 try{
  setThemeColors(newer);
  assert.throws(()=>withThemeColors(DEFAULT_THEME,()=>{assert.equal(pagePaint(),'#dae2f3');throw Error('cancel');}),/cancel/);
  assert.strictEqual(THEME.colors,newer);
  const key=THEME.key;
  withThemeColors(DEFAULT_THEME,()=>{
   assert.equal(withThemeColors(newer,pagePaint),'#c0e6f5');
   assert.strictEqual(THEME.colors,DEFAULT_THEME);
  });
  assert.strictEqual(THEME.colors,newer);assert.equal(THEME.key,key);assert.equal(pagePaint(),'#c0e6f5');
 }finally{setThemeColors(previous);}
});

test('잘못된 테마 자료형이 현재 문서의 전역 색/키를 오염시키지 않는다',()=>{
 const previous=THEME.colors;
 try{
  setThemeColors(newer);const key=THEME.key;
  for(const bad of ['invalid-theme',{length:2}]){
   let called=false;
   assert.throws(()=>withThemeColors(bad,()=>{called=true;}),TypeError);
   assert.equal(called,false);assert.strictEqual(THEME.colors,newer);assert.equal(THEME.key,key);
  }
  assert.throws(()=>withThemeColors(DEFAULT_THEME,()=>withThemeColors('invalid-theme',pagePaint)),TypeError);
  assert.strictEqual(THEME.colors,newer);assert.equal(THEME.key,key);assert.equal(pagePaint(),'#c0e6f5');
 }finally{setThemeColors(previous);}
});
