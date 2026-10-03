import test from 'node:test';
import assert from 'node:assert/strict';
import { Workbook } from '../src/workbook.js';
import { CellMap } from '../src/cellmap.js';
import { ColBuilder } from '../src/block.js';
import { xlsxRowPlan } from '../src/xlsx-row-stream.js';
import { writeXlsx, writeXlsxAsync, readXlsx, xlsxOverflow } from '../src/xlsx.js';
import { EXCEL_MAX_ROWS, MAX_COLS } from '../src/formula.js';
import { unzip,textOf } from '../src/zip.js';
const opts={rowLimit:EXCEL_MAX_ROWS,colLimit:MAX_COLS};

test('compressed blank runs remain lazy when row order and Excel overflow are planned',()=>{
 const cells=new CellMap();cells.setRunRC(10,3,1000000,{raw:'',style:{fill:'#abcdef'}});
 cells.setRC(12,3,{raw:'7'});cells.forEachRC=()=>{throw Error('logical expansion')};cells[Symbol.iterator]=()=>{throw Error('logical expansion')};
 const plan=xlsxRowPlan(cells,opts);assert.equal(plan.maxR,1000009);assert.equal(plan.maxC,3);
 for(let r=10;r<14;r++){const next=plan.rows.next().value;assert.equal(next[0],r);assert.equal(next[1][0][1].raw,r===12?'7':'');}plan.rows.return();
 cells.setRunRC(EXCEL_MAX_ROWS-2,4,9,{raw:'',style:{bold:true}});
 const wb=new Workbook();wb.sheets[0].cells=cells;assert.equal(xlsxOverflow(wb),7);
});

test('row intervals merge sparse points, blanks, metadata, blocks and spill positions in order',()=>{
 const cells=new CellMap();cells.setRC(45,2,{raw:'45'});cells.setRunRC(10,3,200,{raw:'',style:{bold:true}});cells.setRC(11,3,{raw:'11'});cells.setRC(0,0,{raw:'start'});
 const before=[...cells.storageEntries()];const plan=xlsxRowPlan(cells,{...opts,extraRows:[400,11,402],blocks:[{r0:200,c0:8,n:5,cols:[{}]}],spills:[{r:1,c:1,h:2,w:2}]});
 const actual=[...plan.rows];const expected=[0,1,2,...Array.from({length:200},(_,i)=>i+10),400,402];assert.deepEqual(actual.map(([r])=>r),expected);assert.equal(actual.find(([r])=>r===11)[1].filter(([c])=>c===3).length,1);
 assert.equal(actual.find(([r])=>r===1)[1].length,1);assert.equal(actual.find(([r])=>r===2)[1].length,2);assert.deepEqual([...cells.storageEntries()],before);
});

for(const async of [false,true])test(`XLSX streamed rows preserve styles, types, formulas, metadata and ordered cells: async=${async}`,async()=>{
 const wb=new Workbook({sheets:[{name:'Sheet1',cells:{},cellRuns:[[3,0,1000,{raw:'',style:{fill:'#abcdef',bold:true}}]]}]});
 const col=new ColBuilder(4);[0,false,{error:'#N/A'},'text'].forEach((v,i)=>col.set(i,v));wb.sheets[0].blocks=[{r0:1,c0:3,n:4,cols:[col.finish()]}];
 wb.setInput(0,5,0,'12');wb.setInput(0,5,1,'=A6*2');wb.setCellData(0,8,1,{raw:'link',link:'https://example.invalid',comment:'메모'});
 wb.sheets[0].merges=[{r1:0,c1:0,r2:0,c2:2}];wb.sheets[0].rowHeights[1400]=36;wb.sheets[0].rowManual[1400]=true;wb.sheets[0].hiddenRows[1401]=true;
 wb.setInput(0,1,6,'=SEQUENCE(2,2,10)');wb.getValue(0,1,6);const before=JSON.stringify(wb.serialize());
 wb.sheets[0].cells.forEachRC=()=>{throw Error('logical cell expansion')};wb.sheets[0].cells[Symbol.iterator]=()=>{throw Error('logical cell expansion')};
 const bytes=async?await writeXlsxAsync(wb):writeXlsx(wb),files=unzip(bytes),xml=textOf(files['xl/worksheets/sheet1.xml']);
 const back=new Workbook(readXlsx(bytes).data);assert.equal(back.getValue(0,5,0),12);assert.equal(back.getValue(0,5,1),24);assert.equal(back.getRaw(0,5,1),'=A6*2');
 assert.equal(back.styleAt(0,4,0).fill,'#abcdef');assert.equal(back.styleAt(0,1002,0).bold,true);assert.equal(back.getValue(0,1,3),0);assert.equal(back.getValue(0,2,3),false);assert.equal(back.getValue(0,3,3).code,'#N/A');assert.equal(back.getValue(0,4,3),'text');
 assert.equal(back.getCell(0,8,1).comment,'메모');assert.equal(back.getCell(0,8,1).link,'https://example.invalid');assert.deepEqual(back.sheets[0].merges,wb.sheets[0].merges);assert.equal(back.sheets[0].rowHeights[1400],36);assert.equal(back.sheets[0].hiddenRows[1401],true);
 assert.match(xml,/<c r="H3"[^>]*><v>13<\/v><\/c>/);const rows=[...xml.matchAll(/<row r="(\d+)"/g)].map(m=>+m[1]);assert.ok(rows.every((r,i)=>!i||r>rows[i-1]));assert.equal(JSON.stringify(wb.serialize()),before);
});
