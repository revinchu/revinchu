import test from 'node:test';
import assert from 'node:assert/strict';
import {Workbook} from '../src/workbook.js';
import {chartModelData} from '../src/chart.js';
import {writeXlsx,readXlsx} from '../src/xlsx.js';
import {unzip,zip,textOf} from '../src/zip.js';
import {parseXml,descendants,child} from '../src/xml.js';

function fixture(n=3000,type='column'){
 const cells={'0,0':{raw:'항목'},'0,1':{raw:'금액'}};
 for(let r=1;r<=n;r++){cells[r+',0']={raw:'행 '+r};cells[r+',1']={raw:String(r===n?99999:r)};}
 return new Workbook({sheets:[{name:'전체 원본',cells,charts:[{id:'full-cache',type,range:{r1:0,c1:0,r2:n,c2:1},x:0,y:0,w:520,h:340}]}]});
}
const chart=wb=>wb.sheets[0].charts[0];
const chartFile=wb=>{const files=unzip(writeXlsx(wb));const name=Object.keys(files).find(n=>/^xl\/charts\/chart\d+\.xml$/.test(n));return{files,name,xml:parseXml(textOf(files[name]))};};
const cache=(xml,tag)=>descendants(descendants(xml,tag)[0],'numCache')[0];
const points=node=>new Map(descendants(node,'pt').map(p=>[Number(p.attrs.idx),child(p,'v')?.text??p.text]));
const count=node=>Number(child(node,'ptCount')?.attrs.val);

// The UI budget is not a data-export budget. Real source indexes must never be renumbered from a sample.
test('화면 샘플은 유지하고 저장 모드는 3000행의 원래 값·순서를 읽는다',()=>{
 const wb=fixture(),ch=chart(wb),ui=chartModelData(wb,0,ch);assert.ok(ui.categories.length<=2000);assert.equal(ui.series[0].values.at(-1),99999);
 let calls=0,rows;const get=wb.getValue.bind(wb);wb.getValue=(...args)=>{calls++;return get(...args);};
 const full=chartModelData(wb,0,ch,{sample:false,onRange:r=>{rows=r;}});
 assert.equal(calls,6002);assert.equal(rows.length,3001);assert.equal(full.categories.length,3000);assert.equal(full.series[0].values[1000],1001);assert.equal(full.series[0].values[2999],99999);assert.equal(ch.range.r2,3000);
});

test('표준 XLSX 전체 범위 캐시는 count=3000·idx=2999 및 중간 빈칸 위치를 보존한다',()=>{
 const wb=fixture();wb.transact(()=>wb.setInput(0,1235,1,''));chart(wb).hiddenCats=[9];
 const {files,name,xml}=chartFile(wb),values=cache(xml,'val'),vp=points(values),categories=descendants(descendants(xml,'cat')[0],'strCache')[0];
 assert.equal(count(values),3000);assert.equal(vp.get(0),'1');assert.equal(vp.get(1000),'1001');assert.equal(vp.has(1234),false);assert.equal(vp.get(1235),'1236');assert.equal(vp.get(2999),'99999');
 assert.equal(count(categories),3000);assert.equal(points(categories).get(2999),'행 3000');
 assert.ok(descendants(xml,'f').some(n=>n.text.endsWith('$B$2:$B$3001')));
 files[name]=textOf(files[name]).replace(/<c:extLst>[\s\S]*?<\/c:extLst>/g,'');
 const restored=new Workbook(readXlsx(zip(files)).data),data=chartModelData(restored,0,chart(restored),{sample:false});
 assert.equal(data.series[0].values.length,3000);assert.equal(data.series[0].values[1000],1001);assert.equal(data.series[0].values[2999],99999);
});

test('분산·거품 명시 계열의 X/Y/크기 캐시도 샘플링 없이 같은 행 인덱스를 쓴다',()=>{
 const wb=fixture(2400,'bubble');wb.transact(()=>{for(let r=1;r<=2400;r++){wb.setInput(0,r,0,String(r*2));wb.setInput(0,r,2,String(r*3));}});
 const ref=c=>({sheet:'전체 원본',r1:1,c1:c,r2:2400,c2:c});const ch=chart(wb);delete ch.range;ch.series=[{name:{text:'계열'},x:ref(0),val:ref(1),size:ref(2)}];
 const {xml}=chartFile(wb);
 for(const [tag,last]of [['xVal','4800'],['yVal','99999'],['bubbleSize','7200']]){const v=cache(xml,tag);assert.equal(count(v),2400,tag);assert.equal(points(v).get(2399),last,tag);}
});

test('저장 모드의 가로 계열은 101열 화면 제한을 적용하지 않는다',()=>{
 const cells={'0,0':{raw:'항목'},'1,0':{raw:'계열'}};for(let c=1;c<=130;c++){cells['0,'+c]={raw:'열 '+c};cells['1,'+c]={raw:String(c)};}
 const wb=new Workbook({sheets:[{name:'가로 원본',cells,charts:[{type:'line',range:{r1:0,c1:0,r2:1,c2:130},w:500,h:300}]}]});
 assert.equal(chartModelData(wb,0,chart(wb)).series[0].values.length,100);
 const {xml}=chartFile(wb),v=cache(xml,'val');assert.equal(count(v),130);assert.equal(points(v).get(129),'130');
});

test('전체 열 참조는 값과 동적 배열까지 읽고 빈 100만행 꼬리는 읽지 않는다',()=>{
 const wb=new Workbook({sheets:[{name:'원본',cells:{'0,0':{raw:'항목'},'0,1':{raw:'값'},'1,0':{raw:'=SEQUENCE(3)'},'1,1':{raw:'=SEQUENCE(3,1,10)'}},charts:[{type:'column',range:{r1:0,c1:0,r2:1048575,c2:1},w:500,h:300}]}]});
 wb.transact(()=>wb.setCellData(0,1048575,0,{raw:'',style:{fill:'#FF0000'}}));
 let calls=0;const get=wb.getValue.bind(wb);wb.getValue=(...args)=>{assert.ok(++calls<200,'값이 없는 전체 열 순회 금지');return get(...args);};
 const data=chartModelData(wb,0,chart(wb),{sample:false});assert.equal(data.series[1]?.values.at(-1)??data.series[0].values.at(-1),12);assert.equal(chart(wb).range.r2,1048575);
 const {xml}=chartFile(wb);assert.ok(descendants(xml,'f').some(n=>n.text.endsWith('$1048576')));assert.ok(descendants(xml,'numCache').every(c=>count(c)<=3));
});

test('전체 행 참조의 빈 16384열 꼬리를 저장 시 순회하지 않는다',()=>{
 const wb=new Workbook({sheets:[{name:'원본',cells:{'0,0':{raw:'항목'},'0,1':{raw:'A'},'0,2':{raw:'B'},'1,0':{raw:'값'},'1,1':{raw:'10'},'1,2':{raw:'20'}},charts:[{type:'column',range:{r1:0,c1:0,r2:1,c2:16383},w:500,h:300}]}]});
 let calls=0;const get=wb.getValue.bind(wb);wb.getValue=(...args)=>{assert.ok(++calls<200,'값 없는 전체 행 순회 금지');return get(...args);};
 const {xml}=chartFile(wb),v=cache(xml,'val');assert.equal(count(v),2);assert.equal(points(v).get(1),'20');assert.ok(descendants(xml,'f').some(n=>n.text.endsWith('$XFD$2')));
});

test('ChartEx 폭포의 표준 데이터도 전체 2400개 값과 마지막 outlier를 저장한다',()=>{
 const wb=fixture(2400,'waterfall'),files=unzip(writeXlsx(wb)),path=Object.keys(files).find(n=>/^xl\/charts\/chart\d+\.xml$/.test(n));assert.ok(path);
 const xml=parseXml(textOf(files[path])),dims=descendants(xml,'numDim'),lv=descendants(dims[0],'lvl')[0];assert.equal(Number(lv.attrs.ptCount),2400);assert.equal(points(lv).get(2399),'99999');
});


for(const header of [false,true])test(`날짜 첫 열은 머리글 ${header?'있음':'없음'}에서도 범주/값 참조와 캐시가 일치한다`,()=>{
 const cells={};if(header){cells['0,0']={raw:'일자'};cells['0,1']={raw:'금액'};}
 for(let i=0;i<3;i++){const r=i+(header?1:0);cells[r+',0']={raw:String(45000+i),style:{numFmt:'date',code:'yyyy-mm-dd'}};cells[r+',1']={raw:String(10+i)};}
 const wb=new Workbook({sheets:[{name:'날짜',cells,charts:[{type:'column',range:{r1:0,c1:0,r2:header?3:2,c2:1},w:500,h:300}]}]});
 const original=JSON.stringify(wb.serialize().sheets[0].cells),{xml}=chartFile(wb),v=descendants(xml,'val')[0],c=descendants(xml,'cat')[0],off=header?2:1;
 assert.equal(descendants(xml,'ser').length,1);assert.equal(descendants(v,'f')[0].text,`날짜!$B$${off}:$B$${off+2}`);assert.equal(descendants(c,'f')[0].text,`날짜!$A$${off}:$A$${off+2}`);assert.equal(count(cache(xml,'val')),3);assert.equal(points(cache(xml,'val')).get(0),'10');assert.equal(JSON.stringify(wb.serialize().sheets[0].cells),original);
});

test('전체 열 표준 재열기의 머리글 제외 A2 참조도 빈 꼬리 없이 재저장한다',()=>{
 const wb=fixture(3);chart(wb).range.r2=1048575;const {files,name}=chartFile(wb);files[name]=textOf(files[name]).replace(/<c:extLst>[\s\S]*?<\/c:extLst>/g,'');
 const restored=new Workbook(readXlsx(zip(files)).data);let calls=0;const get=restored.getValue.bind(restored);restored.getValue=(...a)=>{assert.ok(++calls<200,'A2 전체 열 참조에서 빈 꼬리 순회 금지');return get(...a);};
 const {xml}=chartFile(restored),v=cache(xml,'val');assert.equal(count(v),3);assert.equal(points(v).get(2),'99999');assert.ok(descendants(xml,'f').some(n=>n.text.endsWith('$B$2:$B$1048576')));
});

test('범위를 가진 스냅샷 차트는 기존 표준 원본 연결과 저장값을 함께 보존한다',()=>{
 const wb=fixture(3);chart(wb).snapshotData={categories:['A','B','C'],series:[{name:'저장 값',values:[7,8,9]}]};const {xml}=chartFile(wb);
 assert.ok(descendants(xml,'f').some(n=>n.text.endsWith('$B$2:$B$4')));assert.equal(points(cache(xml,'val')).get(2),'9');
});


for(const explicit of [false,true])test('3000항목 '+(explicit?'명시 계열':'범위')+'의 화면 색·필터는 원본 항목 번호로 저장된다',()=>{
 const wb=fixture(),ch=chart(wb);if(explicit){delete ch.range;ch.series=[{name:{text:'값'},cat:{r1:1,c1:0,r2:3000,c2:0},val:{r1:1,c1:1,r2:3000,c2:1}}];}
 const ui=chartModelData(wb,0,ch),sr=ui.series[0];assert.equal(sr._pi.at(-1),2999);assert.equal(ui._ci.at(-1),2999);
 ch.seriesFmt=[{pointColors:{[sr._pi.at(-1)]:'#123456'}}];const {xml}=chartFile(wb),point=descendants(xml,'dPt').find(n=>child(n,'idx')?.attrs.val==='2999');assert.ok(point);assert.equal(descendants(point,'srgbClr')[0].attrs.val,'123456');assert.equal(points(cache(xml,'val')).get(2999),'99999');
 ch.hiddenCats=[2999];const hidden=chartModelData(wb,0,ch);assert.equal(hidden.categories.includes('행 3000'),false);assert.equal(hidden.series[0].values.includes(99999),false);assert.equal(hidden.series[0]._pi.includes(2999),false);
});
test('행 기준 대규모 차트는 표본 계열도 원래 계열 서식과 필터 번호를 유지한다',()=>{
 const wb=fixture(),ch=chart(wb);ch.byRows=true;ch.seriesFmt=[];ch.seriesFmt[2999]={color:'#123456'};
 const data=chartModelData(wb,0,ch),last=data.series.at(-1);assert.equal(last._fi,2999);assert.equal(last.color,'#123456');assert.deepEqual(last.values,[99999]);ch.hiddenSeries=[2999];assert.equal(chartModelData(wb,0,ch).series.some(s=>s._fi===2999),false);
});
