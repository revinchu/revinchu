import test from 'node:test';
import assert from 'node:assert/strict';
import { Workbook } from '../src/workbook.js';
import { oleColor, sampleNativeTargets, compareNativeSamples, POINT_TO_CSS_PX } from '../tools/native-layout-compare.mjs';
const nativeStyle=()=>({font:{name:'Arial',size:11,bold:true,italic:false,underline:-4142,strike:false,color:0x332211},fill:{pattern:-4142,color:0xffffff},borders:Object.fromEntries(['left','top','right','bottom','diagonalDown','diagonalUp'].map(k=>[k,{style:-4142,weight:2,color:0}])),numberFormat:'G/표준',alignment:{horizontal:1,vertical:-4107,wrap:false,shrink:false,indent:0,orientation:-4128}});
const targets={sheets:[{index:0,name:'표본',cells:['A1'],rows:[1],columns:['A']}]};
function workbook(extra={}){const wb=new Workbook({defaultFont:{name:'Arial',size:11},sheets:[{name:'표본',cells:{'0,0':{raw:'',style:{bold:true,color:'#112233'}}},...extra}]});wb.getValue=()=>{throw Error('원본 재계산 금지');};return wb;}
function native(){const direct=nativeStyle();return {opened:true,readOnly:true,ownershipVerified:true,originalUnchanged:true,status:'complete',sheets:[{index:0,name:'표본',cells:[{address:'A1',status:'complete',direct,display:structuredClone(direct),geometry:{left:0,top:0,width:48,height:15},merged:false}],rows:[{row:1,top:0,height:15,hidden:false}],columns:[{column:'A',left:0,width:48,columnWidth:999,hidden:false}],drawings:[],shapesLimited:false}]};}

test('OLE RGB 순서·CSS96/pt72는 화면 DPR·ColumnWidth 문자 값과 독립적이다',async()=>{
 assert.equal(oleColor(0x332211),'#112233');assert.equal(oleColor(0),'#000000');assert.equal(oleColor(-1),null);assert.equal(POINT_TO_CSS_PX,4/3);
 const model=await sampleNativeTargets(workbook(),targets);const comparison=compareNativeSamples(native(),model);assert.equal(comparison.equal,true,JSON.stringify(comparison));assert.equal(comparison.units.zoomDprApplied,false);
});
test('Native 숫자 null은 0으로 만들지 않으며 보이지 않는 선의 sentinel은 비교하지 않는다',async()=>{
 const wb=workbook({shapes:[{name:'상자',z:1,x:96,y:192,w:48,h:24,kind:'rect',stroke:null,fill:null}]}),model=await sampleNativeTargets(wb,targets),n=native();
 n.sheets[0].drawings=[{Name:'상자',ZOrderPosition:1,Left:72,Top:144,Width:36,Height:18,Visible:-1,Placement:2,Rotation:0,line:{visible:0,weight:-2147483600,color:{rgb:null}},fill:{visible:0,type:-2}}];
 assert.equal(compareNativeSamples(n,model).equal,true);
 n.sheets[0].cells[0].geometry.left=null;const incomplete=compareNativeSamples(n,model);assert.equal(incomplete.equal,false);assert.ok(incomplete.counts.unavailable>0);assert.equal(incomplete.counts.mismatches,0);
});
test('테두리 종류/굵기/색·그림 위치와 선 point두께를 구별한다',async()=>{
 const wb=workbook({shapes:[{name:'선',z:1,x:96,y:192,w:48,h:24,kind:'rect',stroke:'#112233',strokeWidth:2,dash:'dashDot',compound:'dbl',fill:null}]});
 wb.sheets[0].cells.getRC(0,0).style.bb=true;wb.sheets[0].cells.getRC(0,0).style.bbs='medium';wb.sheets[0].cells.getRC(0,0).style.bbc='#112233';
 const model=await sampleNativeTargets(wb,targets),n=native();n.sheets[0].cells[0].direct.borders.bottom={style:1,weight:-4138,color:0x332211};n.sheets[0].cells[0].display=structuredClone(n.sheets[0].cells[0].direct);
 n.sheets[0].drawings=[{Name:'선',ZOrderPosition:1,Left:72,Top:144,Width:36,Height:18,Visible:-1,Placement:2,Rotation:0,line:{visible:-1,weight:1.5,dash:5,style:2,color:{rgb:0x332211}},fill:{visible:0,type:1}}];
 assert.equal(compareNativeSamples(n,model).equal,true,JSON.stringify(compareNativeSamples(n,model)));
 n.sheets[0].drawings[0].Width=39;n.sheets[0].cells[0].direct.borders.bottom.weight=2;n.sheets[0].drawings[0].line.color.rgb=0;const changed=compareNativeSamples(n,model);
 assert.ok(changed.categories.geometry>0);assert.ok(changed.categories.border>0);assert.ok(changed.categories['drawing-line']>0);
});
test('DisplayFormat 변화는 기록하고 직접 서식 비교를 실제 CF 렌더 합격으로 바꾸지 않는다',async()=>{
 const n=native(),model=await sampleNativeTargets(workbook(),targets);n.sheets[0].cells[0].display.font.color=0xff;
 const result=compareNativeSamples(n,model);assert.equal(result.equal,true);assert.equal(result.counts.displayStyleChanged,1);assert.match(result.scope,/no actual screen/);
 n.originalUnchanged=false;assert.equal(compareNativeSamples(n,model).equal,false);
});
test('마지막 sparse 셀 표본은 사각형 전체 또는 getValue를 스캔하지 않는다',async()=>{
 const wb=workbook({rowHeights:{1:27},colWidths:{1:35},hiddenRows:{3:true},hiddenCols:{2:true}}),t={sheets:[{index:0,name:'표본',cells:['B2','XFD1048576'],rows:[4],columns:['C']}]};let calls=0;const at=wb.styleAt.bind(wb);wb.styleAt=(...args)=>{calls++;return at(...args);};
 const sampled=await sampleNativeTargets(wb,t);assert.equal(calls,2);assert.equal(sampled.sheets[0].cells[0].geometry.width,35);assert.equal(sampled.sheets[0].cells[0].geometry.height,27);assert.equal(sampled.sheets[0].rows[0].height,0);assert.equal(sampled.sheets[0].columns[0].width,0);
});
test('개체 추가/누락은 좌표가 같은 다른 이름도 별도로 검출한다',async()=>{
 const model=await sampleNativeTargets(workbook({shapes:[{name:'추가',z:1,x:0,y:0,w:20,h:10,kind:'rect'}]}),targets),result=compareNativeSamples(native(),model);
 assert.equal(result.equal,false);assert.ok(result.categories['drawing-presence']>0);
});


test('그룹 표본은 총 node 예산과 깊이8 안에서만 복사한다',async()=>{
 const children=Array.from({length:20},(_,i)=>({name:'자식'+i,x:0,y:0,w:10,h:10,kind:'rect'})),wb=workbook({shapes:[{name:'그룹',z:1,x:0,y:0,w:10,h:10,kind:'group',children}]});
 const model=await sampleNativeTargets(wb,targets,{maxObjects:3});assert.equal(model.sheets[0].drawings[0].children.length,2);assert.equal(model.sheets[0].drawings[0].childrenLimited,true);
 assert.equal(wb.sheets[0].shapes[0].children.length,20);
});
