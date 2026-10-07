import test from 'node:test';
import assert from 'node:assert/strict';
import { CellMap } from '../src/cellmap.js';
import { DepGraph } from '../src/depgraph.js';
import { parse } from '../src/formula.js';

const ast=parse('A1*2');
function fixture(n=7001,order='reverse',sheetCount=1) {
  const sheets=Array.from({length:sheetCount},(_,si)=>{
    const cells=new CellMap();cells.setRC(0,10,{raw:'literal',v:'literal'});
    for(const c of [4,1,7])for(let i=0;i<n;i++) {
      const r=order==='reverse'?n-1-i:i;
      cells.setRC(r,c,{raw:'=A'+(r+1)+'*2',formula:true,ast,dr:r});
    }
    cells.setRC(0,12,{raw:'=text',v:'=text',inputType:'text'});
    cells.setRC(1,12,{raw:'link',v:'link',link:'#Synthetic!A1'});
    return {name:'Synthetic'+si,cells};
  });
  return {sheets,sheetIndexByName:name=>sheets.findIndex(s=>s.name===name)};
}
function* legacySteps() {
  let n=0;
  for(let si=0;si<this.wb.sheets.length;si++) {
    const sheet=this.wb.sheets[si],columns=new Map();
    for(const [r,c]of sheet.cells.formulaEntries()) {let rows=columns.get(c);if(!rows){rows=[];columns.set(c,rows);}rows.push(r);}
    for(const [c]of sheet.cells.cols) {const rows=columns.get(c);if(!rows)continue;rows.sort((a,b)=>a-b);
      for(const r of rows){this.add(si,r,c,sheet.cells.getRC(r,c));if(++n%20000===0)yield n;}
    }
  }
  this.open.clear();
}
function runs(g) {
  const found=new Set();for(const cols of g.cols)if(cols)for(const b of cols.values())for(const r of b.runs)found.add(r);
  for(const list of g.wide)if(list)for(const r of list)found.add(r);
  return [...found].map(({ts,s,fc,fr0,n,r1,r2,r1s,r2s,c1,c2,lo,hi,dr0,drs,dc})=>({ts,s,fc,fr0,n,r1,r2,r1s,r2s,c1,c2,lo,hi,dr0,drs,dc}));
}
function build(wb,legacy=false) {
  const g=new DepGraph(wb,true),trace=[],add=g.add.bind(g);
  g.add=(si,r,c,cell)=>{trace.push([si,r,c]);add(si,r,c,cell);};
  const values=[...(legacy?legacySteps.call(g):g.steps())];
  return {g,trace,values};
}

test('candidate collection yields at 20k before any graph registration or sort/add work',()=>{
  const wb=fixture(),cells=wb.sheets[0].cells,original=cells.formulaEntries.bind(cells);let collected=0,added=0;
  cells.formulaEntries=function*(){for(const row of original()){collected++;yield row;}};
  const g=new DepGraph(wb,true),add=g.add.bind(g);g.add=(...args)=>{added++;add(...args);};const it=g.steps();
  assert.deepEqual(it.next(),{value:undefined,done:false});assert.equal(collected,20000);assert.equal(added,0);assert.equal(g.open.size,0);
  // The tail and first column sort have their own checkpoint before registration.
  assert.deepEqual(it.next(),{value:undefined,done:false});assert.equal(collected,21003);assert.equal(added,0);
  assert.deepEqual(it.next(),{value:undefined,done:false});assert.equal(added,7001);assert.equal(g.open.size,1);
  const rest=[...it];assert.deepEqual(rest.filter(x=>typeof x==='number'),[20000]);assert.equal(added,21003);assert.equal(g.open.size,0);
});

for(const order of ['ascending','reverse'])test('yield checkpoints preserve sheet/column/sorted-row add order and positive 20k counts: '+order,()=>{
  const wb=fixture(7001,order,2),before=wb.sheets.map(s=>[...s.cells.formulaEntries()].map(([r,c])=>[r,c]));
  const old=build(wb,true),now=build(wb);
  assert.deepEqual(now.trace,old.trace);assert.deepEqual(now.values.filter(x=>typeof x==='number'),old.values);
  assert.deepEqual(old.values,[20000,40000]);assert.deepEqual(runs(now.g),runs(old.g));assert.equal(now.g.open.size,0);
  assert.deepEqual(wb.sheets.map(s=>[...s.cells.formulaEntries()].map(([r,c])=>[r,c])),before,'iterator checkpoints never reorder cell storage');
  for(let si=0;si<2;si++)for(const r of [0,3500,7000])assert.deepEqual(now.g.dependentsOf(si,r,0),old.g.dependentsOf(si,r,0));
});

test('synchronous constructor drains new checkpoints and produces the unchanged fill runs',()=>{
  const wb=fixture(57,'reverse',2),old=build(wb,true).g,now=new DepGraph(wb);
  assert.deepEqual(runs(now),runs(old));assert.equal(now.open.size,0);
  assert.equal(runs(now).length,6);assert.ok(runs(now).every(r=>r.n===57));
  for(let si=0;si<2;si++)assert.deepEqual(now.dependentsOf(si,25,0),old.dependentsOf(si,25,0));
});

test('empty and non-formula candidate sheets remain an empty synchronously drained graph',()=>{
  const cells=new CellMap();cells.setRC(1,0,{raw:'=text',v:'=text',inputType:'text'});cells.setRC(2,0,{raw:'link',v:'link',link:'#Synthetic!A1'});
  const g=new DepGraph({sheets:[{name:'Synthetic',cells}],sheetIndexByName:()=>0});
  assert.deepEqual(runs(g),[]);assert.equal(g.dyn.size,0);assert.equal(g.open.size,0);
});
