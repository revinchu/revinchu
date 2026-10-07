import test from 'node:test';
import assert from 'node:assert/strict';
import {CellMap,getSharedBlankCell} from '../src/cellmap.js';
const astA=Object.freeze({kind:'a'}),astB=Object.freeze({kind:'b'});
const formula=(ast=astA)=>({raw:'=A1',formula:true,ast,link:'#A1'});

test('100k formula candidates retain only shared AST values without per-coordinate wrapper objects',()=>{
 const cells=new CellMap(),a=formula(),b=formula(astB);
 for(let r=0;r<100000;r++)cells.setRC(r,2,r%2?a:b);
 const candidates=cells.candidates.get(2);assert.equal(candidates.size,100000);assert.equal(cells.size,100000);
 let count=0;for(const value of candidates.values()){assert.ok(value===astA||value===astB);count++;}
 assert.equal(count,100000);assert.equal(cells.formulaAsts.get(astA),50000);assert.equal(cells.formulaAsts.get(astB),50000);
 assert.equal(new Set(candidates.values()).size,2);let sum=0;cells.forEachFormulaRC((cell,r,c)=>{assert.equal(c,2);assert.equal(cell,r%2?a:b);sum+=r;});assert.equal(sum,4999950000);
});

test('undefined raw/link candidate slots remain present and null separately tracks malformed formulas',()=>{
 const cells=new CellMap();cells.setRC(0,0,{raw:'=A1'});cells.setRC(1,0,{raw:'link',link:'#A1'});cells.setRC(2,0,{raw:'=bad',formula:true});
 const col=cells.candidates.get(0);assert.equal(col.has(0),true);assert.equal(col.get(0),undefined);assert.equal(col.has(1),true);assert.equal(col.get(1),undefined);assert.equal(col.get(2),null);assert.equal(cells.formulaAsts.get(null),1);
 cells.setRC(0,0,formula());assert.equal(col.get(0),astA);cells.setRC(0,0,{raw:'text'});assert.equal(col.has(0),false);assert.equal(cells.formulaAsts.has(astA),false);
 cells.setRC(2,0,{raw:'link',link:'#A1'});assert.equal(col.get(2),undefined);assert.equal(cells.formulaAsts.has(null),false);
});

test('same mutable cell shared at two coordinates subtracts each previous AST during in-place normalization',()=>{
 const cells=new CellMap(),cell=formula();cells.setRC(1,0,cell);cells.setRC(3,0,cell);
 for(const count of cells.mapValues(value=>{value.ast=astB;return value;},value=>value))assert.ok(count);
 assert.equal(cells.formulaAsts.has(astA),false);assert.equal(cells.formulaAsts.get(astB),2);assert.deepEqual([...cells.candidates.get(0).values()],[astB,astB]);
 const saved=new CellMap(cells);cells.setRunRC(0,0,5,getSharedBlankCell({bold:true}));assert.equal(cells.candidates.size,0);assert.equal(cells.formulaAsts.size,0);
 assert.equal(saved.formulaAsts.get(astB),2);assert.equal([...saved.formulaEntries()].length,2);
});

for(const method of ['forEachFormulaRC','forEachLinkRC','formulaEntries'])test(method+' preserves existing inner-column order across clear and same-column recreation',()=>{
 const cells=new CellMap();cells.setRC(1,0,formula());cells.setRC(3,0,formula(astB));const seen=[];
 const mutate=(cell,r,c)=>{seen.push([r,c]);if(r===1){cells.clear();cells.setRC(5,0,formula());}};
 if(method==='formulaEntries')for(const[r,c,cell]of cells.formulaEntries())mutate(cell,r,c);else cells[method](mutate);
 assert.deepEqual(seen,[[1,0],[3,0],[5,0]]);assert.equal(cells.size,1);assert.deepEqual([...cells.candidates.get(0).keys()],[5]);assert.equal(cells.formulaAsts.get(astA),1);assert.equal(cells.formulaAsts.has(astB),false);
});

test('candidate callbacks observe current replacements, skip removed rows and visit newly appended candidates',()=>{
 const cells=new CellMap(),a=formula(),b=formula(astB);cells.setRC(0,0,a);cells.setRC(1,0,a);cells.setRC(2,0,a);const seen=[];
 cells.forEachFormulaRC((cell,r,c)=>{seen.push([r,cell.ast]);if(r===0){cells.deleteRC(1,c);cells.setRC(2,c,b);cells.setRC(3,c,b);}});
 assert.deepEqual(seen,[[0,astA],[2,astB],[3,astB]]);assert.equal(cells.formulaAsts.get(astA),1);assert.equal(cells.formulaAsts.get(astB),2);assert.deepEqual([...cells.candidates.get(0).values()],[astA,astB,astB]);
});
