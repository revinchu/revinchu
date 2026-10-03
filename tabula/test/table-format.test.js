import test from 'node:test';
import assert from 'node:assert/strict';
import { Workbook } from '../src/workbook.js';
import { tableCellStyle } from '../src/tables.js';
import { tableCellDisplayStyle, clearedTableCellStyle, explicitTableStylePatch, TABLE_VISUAL_KEYS } from '../src/table-format.js';
const table={id:'t',name:'Table1',r1:0,c1:0,r2:3,c2:1,header:true,banded:true,style:'TableStyleMedium3'};
function fixture(baseStyle={},cells={}) {return new Workbook({baseStyle,sheets:[{name:'Sheet1',tables:[table],cells}]});}
test('base white/black defaults do not cover table header or stripe colors',()=>{
  const wb=fixture({fill:'#ffffff',color:'#000000',font:'Arial',size:13});
  for(const r of [0,1,2]) {
    const actual=tableCellDisplayStyle(wb,0,r,0), expected=tableCellStyle(table,r,0);
    assert.equal(actual.fill,expected.fill??'#ffffff'); assert.equal(actual.color,expected.color??'#000000');
    assert.equal(actual.font,'Arial');assert.equal(actual.size,13);
  }
});
test('importing a table preserves deliberate direct formats before a quick style is chosen',()=>{
  const wb=fixture({fill:'#ffffff'},{'0,0':{raw:'header',style:{fill:'#ffff00',color:'#ff0000',bold:false}}});
  const st=tableCellDisplayStyle(wb,0,0,0);
  assert.equal(st.fill,'#ffff00');assert.equal(st.color,'#ff0000');assert.equal(st.bold,false);
});
test('clear removes each visual channel and named style link without losing number/layout/protection',()=>{
  const own={...Object.fromEntries(TABLE_VISUAL_KEYS.map(k=>[k,true])),cellStyleName:'Old',numFmt:'custom',code:'0.00',font:'Arial',size:15,align:'right',wrap:true,locked:false};
  const cleared=clearedTableCellStyle(own,{});
  for(const key of TABLE_VISUAL_KEYS)assert.equal(cleared[key],undefined,key);
  assert.equal(cleared.cellStyleName,undefined);
  assert.deepEqual(cleared,{numFmt:'custom',code:'0.00',font:'Arial',size:15,align:'right',wrap:true,locked:false});
});
test('clear masks all inherited fill channels and displayed table style has no private metadata',()=>{
  const wb=fixture(),sh=wb.sheets[0];sh.allStyle={fill:'#ff0000',pattern:'darkGrid',patternColor:'#00ff00',color:'#ff0000',bold:true};
  sh.cells.setRC(0,0,{raw:'header',style:clearedTableCellStyle({numFmt:'text'},sh.allStyle)});
  const st=tableCellDisplayStyle(wb,0,0,0),ts=tableCellStyle(table,0,0);
  assert.equal(st.fill,ts.fill);assert.equal(st.color,ts.color);assert.equal(st.pattern,undefined);assert.equal(st.tableStyleInherit,undefined);assert.equal(st.numFmt,'text');
});
test('a later solid fill removes inherited fill reset as a group and wins over table fill',()=>{
  const reset=clearedTableCellStyle({}, {fill:'#ff0000',gradient:{stops:[]},pattern:'darkGrid',color:'#ff0000'});
  const style=explicitTableStylePatch(reset,{fill:'#00ffff'});
  assert.deepEqual(style.tableStyleInherit,{color:''});
  const wb=fixture({}, {'0,0':{raw:'header',style}});
  assert.equal(tableCellDisplayStyle(wb,0,0,0).fill,'#00ffff');
});
test('number-only edits preserve resets while direct border edits release just that border',()=>{
  const reset=clearedTableCellStyle({}, {bt:true,bts:'double',btc:'#ff0000',fill:'#ffff00'});
  assert.deepEqual(explicitTableStylePatch(reset,{numFmt:'percent'}).tableStyleInherit,reset.tableStyleInherit);
  assert.deepEqual(explicitTableStylePatch(reset,{bt:true,bts:'thin',btc:'#0000ff'}).tableStyleInherit,{fill:''});
});
test('custom style elements apply to blank cells and can replace base gradients',()=>{
  const wb=fixture({gradient:{stops:[{color:'#000000',pos:0},{color:'#ffffff',pos:1}]}});
  wb.sheets[0].tables[0]={...table,style:'Custom',styleElements:[{type:'wholeTable',style:{fill:'#abcdef'}}]};
  const st=tableCellDisplayStyle(wb,0,3,1);
  assert.equal(st.fill,'#abcdef');assert.equal(st.gradient,undefined);
});
test('effective computed number hints remain above base defaults',()=>{
  const wb=fixture({numFmt:'number',code:'0'}),styleAt=wb.styleAt.bind(wb);
  wb.styleAt=(...args)=>({...styleAt(...args),numFmt:'custom',code:'0.000',queryFormat:'0.000'});
  assert.equal(tableCellDisplayStyle(wb,0,1,0).code,'0.000');
});


test('QUERY spill number formats retain precedence over an existing direct number format', () => {
  const wb = fixture({}, { '1,0': { raw: '2', style: { numFmt: 'number', code: '0' } } });
  wb.spills.set('0:1,0', { r: 1, c: 0, formatStart: 0, formats: ['0.000'] });
  const effective = wb.styleAt(0, 1, 0), displayed = tableCellDisplayStyle(wb, 0, 1, 0);
  for (const key of ['numFmt', 'code', 'queryFormat']) assert.equal(displayed[key], effective[key]);
  assert.equal(displayed.code, '0.000');
});
