// 로컬 native 표본과 이미 열린 모델 비교. 셀 값/수식 본문은 결과에 기록하지 않는다.
import { fmtCode, fileCode } from '../src/format.js';
import { toFileFormula } from '../src/xlfn.js';
const equal=(a,b)=>JSON.stringify(a)===JSON.stringify(b);
const color=n=>typeof n==='number'&&n>=0?'#'+[n&255,(n>>8)&255,(n>>16)&255].map(x=>x.toString(16).padStart(2,'0')).join(''):null;
const align={1:'general',[-4131]:'left',[-4108]:'center',[-4152]:'right',5:'fill',[-4130]:'justify',7:'centerContinuous',[-4117]:'distributed'};
const nativeErrors = new Set([2000,2007,2015,2023,2029,2036,2042,2043,2045,2046,2047,2048,2049,2050,2051,2052]);
export function nativeSampleValue(c) {
 const code=typeof c.value==='number'&&Number.isInteger(c.value)?c.value>>>0:0;
 const cvError=(code>>>16)===0x800a&&nativeErrors.has(code&0xffff);
 return (c.isError===true||cvError)&&typeof c.text==='string'&&/^#(?:NULL!|DIV\/0!|VALUE!|REF!|NAME\?|NUM!|N\/A|SPILL!|CALC!|GETTING_DATA|UNKNOWN!|FIELD!|BLOCKED!|CONNECT!|BUSY!|DATA!|PYTHON!)$/.test(c.text)?{code:c.text}:c.value;
}
export function sameNativeFormula(a,b,hereTable) {
 if(a===b)return true;if(!a||!b)return false;
 try{return toFileFormula(a.replace(/^=/,''),{hereTable})===toFileFormula(b.replace(/^=/,''),{hereTable});}catch{return false;}
}
export function nativeRichImageReview(cell, actual) {
 return nativeSampleValue(cell)?.code==='#UNKNOWN!' && actual.imagePresent===true;
}
export function sameNativeValue(a,b){
 if((a==null||a==='')&&(b==null||b===''))return true;
 if(typeof a==='number'&&typeof b==='number')return Number.isFinite(a)&&Number.isFinite(b)&&Math.abs(a-b)<=Math.max(1,Math.abs(a),Math.abs(b))*1e-10;
 return equal(a,b);
}
const inRange=(address,range)=>{const parse=s=>{const m=/^([A-Z]+)(\d+)$/.exec(s.replace(/\$/g,''));if(!m)return null;let c=0;for(const l of m[1])c=c*26+l.charCodeAt(0)-64;return[+m[2]-1,c-1];};const p=parse(address),[a,b]=String(range||'').split(':').map(parse);return p&&a&&p[0]>=a[0]&&p[1]>=a[1]&&p[0]<=(b||a)[0]&&p[1]<=(b||a)[1];};
export async function auditNativeSamples(page,native){
 const actual=await page.evaluate(sheets=>{
  const w=tabula.wb(),own=w.sheets.slice(0,w.ownSheetCount?.()??w.sheets.length);
  const buttons=(f,a,z)=>!!f&&Array.from({length:Math.max(0,z-a+1)},(_,i)=>a+i).some(c=>!f.hiddenButtons?.[c]);
  return sheets.map(ns=>{const si=own.findIndex(s=>s.name===ns.name),s=own[si];if(!s)return{missing:true};
   return{sample:ns.sample.map(cell=>{const [,col,row]=cell.address.match(/^([A-Z]+)(\d+)$/);let c=0;for(const ch of col)c=c*26+ch.charCodeAt(0)-64;const r=+row-1;--c;
    const st=w.styleAt(si,r,c),v=w.getValue(si,r,c),raw=w.getRaw(si,r,c);return{address:cell.address,hereTable:s.tables?.find(t=>r>=t.r1&&r<=t.r2&&c>=t.c1&&c<=t.c2)?.name,rawType:typeof raw,imagePresent:!!s.cells.getRC(r,c)?.image,errorCode:v&&typeof v==='object'?(v.code??v.error??null):null,value:v&&typeof v==='object'?{code:v.code??v.error??'object'}:v,formula:typeof raw==='string'&&raw.startsWith('=')?raw:null,style:st,font:{name:st.font||w.defaultFont?.name||'맑은 고딕',size:st.size||w.defaultFont?.size||11,bold:!!st.bold},wrap:!!st.wrap};}),
    filter:{range:s.filter?[s.filter.r1,s.filter.c1,s.filter.r2,s.filter.c2]:null,buttons:buttons(s.filter,s.filter?.c1??0,s.filter?.c2??-1),active:Object.keys(s.filter?.criteria||{}).length>0},
    tables:(s.tables||[]).map(t=>({name:t.name,range:[t.r1,t.c1,t.r2,t.c2],headers:!!t.header,totals:!!t.totals,style:t.style,filterButton:buttons(t.filter,t.c1,t.c2),firstCol:!!t.firstCol,lastCol:!!t.lastCol,rowStripes:!!t.banded,colStripes:!!t.bandedCols})),
    pivots:[s.pivot,...(s.pivotsExtra||[])].filter(Boolean).map(d=>({name:d.name,rows:[...(d.rows||[]),...(d.values?.length>1&&d.valuesOnRows?['Σ']:[])],columns:[...(d.cols||[]),...(d.values?.length>1&&!d.valuesOnRows?['Σ']:[])],pages:d.pages||[],values:d.values||[],rowGrand:d.grandCols!==false,colGrand:d.grandRows!==false,fieldHeaders:d.showHeaders!==false,preserveFormat:d.preserveFormat!==false,pageWrap:d.pageWrap||0,pageOrder:d.pageOrder==='over'?2:1,style:d.style}))};
  });
 },native.sheets.map(s=>({name:s.name,sample:(s.sample||[]).map(c=>({address:c.address}))})));
 const styles={count:0,checks:0,differences:[],reviews:[]},content={count:0,nonEmpty:0,formulas:0,pivotCells:0,differences:[],formulaDifferences:[],equivalentFormulaSyntax:[],richValueReviews:[]},options={checks:0,differences:[],skippedChecks:[]};
 for(let si=0;si<native.sheets.length;si++){
  const ns=native.sheets[si],bs=actual[si];if(bs.missing){options.differences.push({sheet:si,field:'missing'});continue;}
  const option=(field,a,b)=>{if(a==null)return;options.checks++;if(!equal(a,b))options.differences.push({sheet:si,field,native:a,wixel:b});};
  option('autoFilter.buttons',ns.autoFilterMode,bs.filter.buttons);option('autoFilter.active',ns.filterMode,bs.filter.active);
  for(let j=0;j<(ns.sample||[]).length;j++){
   const e=ns.sample[j],a=bs.sample[j],at={sheet:si,address:e.address};styles.count++;content.count++;
   const val=nativeSampleValue(e),formula=typeof e.formula==='string'&&e.formula.startsWith('=')?e.formula:null;
   if(val!=null&&val!=='')content.nonEmpty++;if(formula)content.formulas++;if((ns.pivots||[]).some(p=>inRange(e.address,p.range)))content.pivotCells++;
   if(!sameNativeValue(val,a.value)) {
    if(nativeRichImageReview(e,a))content.richValueReviews.push({...at,field:'value',nativeCode:'#UNKNOWN!',wixelImagePresent:true,reason:'Rich cell represented as an image; image identity and adjacent cells require the separate image probe'});
    else content.differences.push({...at,field:'value',nativeType:typeof val,wixelType:typeof a.value,rawType:a.rawType,errorCode:a.errorCode,imagePresent:a.imagePresent,formula:!!formula,pivot:(ns.pivots||[]).some(p=>inRange(e.address,p.range))});
   }
   if(formula!==a.formula)(sameNativeFormula(formula,a.formula,a.hereTable)?content.equivalentFormulaSyntax:content.formulaDifferences).push({...at,field:'formula',nativeFormula:!!formula,wixelFormula:!!a.formula});
   const style=(field,v,w,review=false)=>{if(v==null)return;styles.checks++;if(!equal(v,w))(review?styles.reviews:styles.differences).push({...at,field,native:v,wixel:w});};
   for(const f of ['name','size','bold'])style('font.'+f,e.font?.[f],a.font[f]);style('wrap',e.wrap,a.wrap);
   style('font.color',color(e.font?.color),(a.style.color||'#000000').toLowerCase());style('fill',color(e.fill),(a.style.fill||'#ffffff').toLowerCase());
   style('rotation',e.rotation===-4166?255:e.rotation===-4128?0:e.rotation,a.style.rotate||0);style('align',align[e.align],a.style.align||'general');
   style('numberFormat',(/^(G\/표준|General)$/i.test(e.format||'')?'General':fileCode(e.format||'General')),fmtCode(a.style)||'General',true);
  }
  for(let i=0;i<(ns.tables||[]).length;i++){const e=ns.tables[i],a=bs.tables.find(x=>x.name===e.name);if(!a){option(`table[${i}].present`,true,false);continue;}
   for(const field of ['headers','totals','filterButton','firstCol','lastCol','rowStripes','colStripes'])option(`table[${i}].${field}`,e[field],a[field]);
   if(e.style&&!e.style.includes('__ComObject'))option(`table[${i}].style`,e.style,a.style);else options.skippedChecks.push(`sheet[${si}].table[${i}].style: native style name unavailable`);
  }
  for(let i=0;i<(ns.pivots||[]).length;i++){const e=ns.pivots[i],a=bs.pivots.find(x=>x.name===e.name);if(!a){option(`pivot[${i}].present`,true,false);continue;}
   for(const field of ['rowGrand','colGrand','fieldHeaders','preserveFormat','pageWrap','pageOrder'])option(`pivot[${i}].${field}`,e[field],a[field]);
   for(const field of ['rows','columns','pages','values'])option(`pivot[${i}].${field}.count`,e[field]?.length??0,a[field].length);
   if(e.style&&!e.style.includes('__ComObject'))option(`pivot[${i}].style`,e.style,a.style);else options.skippedChecks.push(`sheet[${si}].pivot[${i}].style: native style name unavailable`);
  }
 }
 return{nativeStyleSamples:styles,nativeContentSamples:content,nativeOptionSamples:options};
}
