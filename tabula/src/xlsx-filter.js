import { child, kids, esc } from './xml.js';
import { cellName, parseRangeName } from './formula.js';

// Sheet/table filters share this codec. Model column keys are absolute.
const OP = { eq:'equal', ne:'notEqual', gt:'greaterThan', ge:'greaterThanOrEqual', lt:'lessThan', le:'lessThanOrEqual' };
const FROM_OP = Object.fromEntries(Object.entries(OP).map(([k,v])=>[v,k]));
const yes = v => v==='1'||v==='true';
const no = v => v==='0'||v==='false';
const ref = r => `${cellName(r.r1,r.c1)}:${cellName(r.r2,r.c2)}`;
const TAGS = new Set(['filters','filter','dateGroupItem','customFilters','customFilter','top10','dynamicFilter','colorFilter','iconFilter','sortState','sortCondition']);
const ATTRS = new Set(['blank','calendarType','val','year','month','day','hour','minute','second','dateTimeGrouping','operator','and','top','percent','filterVal','type','maxVal','cellColor','dxfId','iconSet','iconId','ref','caseSensitive','columnSort','sortMethod','descending','sortBy','customList']);
function cleanNode(node,dxfs) {
  if(!node||!TAGS.has(node.name))return null;
  const out={name:node.name,attrs:{},children:[]};
  for(const [k,v] of Object.entries(node.attrs??{}))if(ATTRS.has(k))out.attrs[k]=String(v);
  if(out.attrs.dxfId!==undefined&&dxfs?.[Number(out.attrs.dxfId)])out.dxfStyle=structuredClone(dxfs[Number(out.attrs.dxfId)]);
  for(const n of node.children??[]){const c=cleanNode(n,dxfs);if(c)out.children.push(c);}
  return out;
}
function nodeXml(node,dxf){
  const n=cleanNode(node);if(!n)return '';
  if(n.attrs.dxfId!==undefined&&node.dxfStyle&&dxf)n.attrs.dxfId=String(dxf(node.dxfStyle));
  const attrs=Object.entries(n.attrs).map(([k,v])=>` ${k}="${esc(v)}"`).join('');
  const inner=(node.children??[]).map(c=>nodeXml(c,dxf)).join('');
  return inner?`<${n.name}${attrs}>${inner}</${n.name}>`:`<${n.name}${attrs}/>`;
}
function readCriterion(fc,dxfs){
  const filters=child(fc,'filters');
  if(filters&&!kids(filters,'dateGroupItem').length){
    const vals=kids(filters,'filter').map(f=>f.attrs.val??'');
    if(yes(filters.attrs.blank)&&!vals.includes(''))vals.push('');return vals;
  }
  const custom=child(fc,'customFilters');
  if(custom){const cs=kids(custom,'customFilter');
    if(cs.length&&cs.length<=2&&cs.every(c=>FROM_OP[c.attrs.operator??'equal']&&c.attrs.val!==undefined&&c.attrs.val!=='')){
      const out={type:'custom',op1:FROM_OP[cs[0].attrs.operator??'equal'],v1:cs[0].attrs.val??'',join:yes(custom.attrs.and)?'and':'or'};
      if(cs[1]){out.op2=FROM_OP[cs[1].attrs.operator??'equal'];out.v2=cs[1].attrs.val??'';}return out;
    }
  }
  const top=child(fc,'top10');
  if(top&&Number.isFinite(Number(top.attrs.val)))return {type:'top',n:Number(top.attrs.val),bottom:no(top.attrs.top),percent:yes(top.attrs.percent)};
  const dyn=child(fc,'dynamicFilter');
  if(dyn&&['aboveAverage','belowAverage'].includes(dyn.attrs.type))return {type:'avg',above:dyn.attrs.type==='aboveAverage'};
  const color=child(fc,'colorFilter');
  if(color){const type=no(color.attrs.cellColor)?'font':'fill',st=dxfs?.[Number(color.attrs.dxfId)],value=type==='font'?(st?.color??st?.fill):st?.fill;
    if(value!==undefined)return {type,value:value??''};
  }
  const node=(fc.children??[]).find(n=>TAGS.has(n.name));
  return node?{type:'xlsx',node:cleanNode(node,dxfs)}:undefined;
}
export function readFilterSort(node,dxfs=[]){
  if(!node)return undefined;
  const cs=kids(node,'sortCondition');if(!cs.length)return undefined;
  const rg=parseRangeName(cs[0].attrs.ref??'');
  const basic=rg&&rg.c1===rg.c2&&!yes(node.attrs.columnSort);
  const sort=basic?{col:rg.c1,asc:!yes(cs[0].attrs.descending)}:{};
  if(!basic||cs.length>1||yes(node.attrs.caseSensitive)||node.attrs.sortMethod||cs.some(c=>(c.attrs.sortBy&&c.attrs.sortBy!=='value')||c.attrs.customList))sort.xlsx=cleanNode(node,dxfs);
  return sort;
}
export function readAutoFilter(node,range,{dxfs=[],sortNode}={}){
  if(!node)return null;
  const criteria={},hiddenButtons={};
  for(const fc of kids(node,'filterColumn')){
    const id=Number(fc.attrs.colId),c=range.c1+id;if(!Number.isInteger(id)||id<0||c>range.c2)continue;
    if(yes(fc.attrs.hiddenButton)||no(fc.attrs.showButton))hiddenButtons[c]=true;
    const cr=readCriterion(fc,dxfs);if(cr!==undefined)criteria[c]=cr;
  }
  const sort=readFilterSort(sortNode??child(node,'sortState'),dxfs);
  return {criteria,hidden:{},...(Object.keys(hiddenButtons).length?{hiddenButtons}:{}),...(sort?{sort}:{})};
}
function customXml(op,value){
  if(!op||value===undefined||value===null||value==='')return '';
  let operator=OP[op],val=String(value);
  if(['begins','notBegins','ends','notEnds','contains','notContains'].includes(op)){
    operator=op.startsWith('not')?'notEqual':'equal';
    // Application prefix/suffix tests are literal; contains accepts wildcards.
    if(!/contains/i.test(op))val=val.replace(/~/g,'~~').replace(/\*/g,'~*').replace(/\?/g,'~?');
    if(/begins/i.test(op))val+='*';else if(/ends/i.test(op))val=`*${val}`;else val=`*${val}*`;
  }
  if(!operator)throw new Error('이 필터 조건은 Excel 표준으로 저장할 수 없습니다. 정규식 조건을 일반 조건으로 바꾼 후 저장하세요.');
  return `<customFilter operator="${operator}" val="${esc(val)}"/>`;
}
function criterionXml(cr,dxf){
  if(Array.isArray(cr))return `<filters${cr.includes('')?' blank="1"':''}>${cr.filter(v=>v!=='').map(v=>`<filter val="${esc(v)}"/>`).join('')}</filters>`;
  if(!cr||typeof cr!=='object')return '';
  if(cr.type==='custom'){const a=customXml(cr.op1,cr.v1),b=customXml(cr.op2,cr.v2);return a||b?`<customFilters${a&&b&&cr.join!=='or'?' and="1"':''}>${a}${b}</customFilters>`:'';}
  if(cr.type==='top')return `<top10 top="${cr.bottom?0:1}" percent="${cr.percent?1:0}" val="${Math.max(1,Number(cr.n)||10)}"/>`;
  if(cr.type==='avg')return `<dynamicFilter type="${cr.above?'aboveAverage':'belowAverage'}"/>`;
  if(cr.type==='fill'||cr.type==='font'){
    if(!dxf)throw new Error('색 필터의 Excel 서식을 저장할 수 없습니다.');
    return `<colorFilter dxfId="${dxf({fill:cr.value||(cr.type==='font'?'#000000':null)})}" cellColor="${cr.type==='fill'?1:0}"/>`;
  }
  if(cr.type==='xlsx'&&cr.node)return nodeXml(cr.node,dxf);
  throw new Error('지원하지 않는 필터 조건을 삭제하지 않고 저장하려면 조건을 먼저 확인하세요.');
}
export function filterSortXml(sort,range,{dxf,header=true}={}){
  if(!sort)return '';if(sort.xlsx)return nodeXml(sort.xlsx,dxf);
  if(!Number.isInteger(sort.col)||sort.col<range.c1||sort.col>range.c2||range.r2<range.r1+(header?1:0))return '';
  const data={...range,r1:range.r1+(header?1:0)};
  return `<sortState ref="${ref(data)}"><sortCondition${sort.asc===false?' descending="1"':''} ref="${ref({...data,c1:sort.col,c2:sort.col})}"/></sortState>`;
}
export function autoFilterXml(filter,range,{dxf,includeSort=true}={}){
  if(!filter)return '';
  const keys=new Set([...Object.keys(filter.criteria??{}),...Object.keys(filter.hiddenButtons??{})]);
  const cols=[...keys].map(Number).filter(c=>Number.isInteger(c)&&c>=range.c1&&c<=range.c2).sort((a,b)=>a-b).map(c=>{
    const inner=criterionXml(filter.criteria?.[c],dxf),hidden=filter.hiddenButtons?.[c]===true;
    return inner||hidden?`<filterColumn colId="${c-range.c1}"${hidden?' hiddenButton="1"':''}>${inner}</filterColumn>`:'';
  }).join('');
  return `<autoFilter ref="${ref(range)}">${cols}${includeSort?filterSortXml(filter.sort,range,{dxf}):''}</autoFilter>`;
}
