import test from 'node:test';
import assert from 'node:assert/strict';
import { nativeSampleValue, nativeRichImageReview, sameNativeFormula, auditNativeSamples } from '../tools/real-file-native-samples.mjs';
test('native audit distinguishes literal error text from COM CVErr',()=>{
 assert.equal(nativeSampleValue({text:'#N/A',value:'#N/A'}),'#N/A');
 assert.equal(nativeSampleValue({text:'#UNKNOWN!',value:'#UNKNOWN!'}),'#UNKNOWN!');
 assert.equal(nativeSampleValue({text:'#N/A',value:2042}),2042);
 assert.deepEqual(nativeSampleValue({text:'#N/A',value:-2146826246}),{code:'#N/A'});
 assert.deepEqual(nativeSampleValue({text:'#UNKNOWN!',value:-2146826240}),{code:'#UNKNOWN!'});
});

test('native formula audit normalizes current-table references without altering text literals',()=>{
 assert.equal(sameNativeFormula('=[@Date]-1','=Sales[[#This Row],[Date]]-1','Sales'),true);
 assert.equal(sameNativeFormula('=[@Date]-1','=Other[[#This Row],[Date]]-1','Sales'),false);
 assert.equal(sameNativeFormula('="[@Date]"','="Sales[[#This Row],[Date]]"','Sales'),false);
});

test('unsupported native rich value with an actual cell image is a separate representation review',()=>{
 assert.equal(nativeRichImageReview({value:-2146826240,text:'#UNKNOWN!'},{imagePresent:true}),true);
 assert.equal(nativeRichImageReview({value:-2146826240,text:'#UNKNOWN!'},{imagePresent:false}),false);
 assert.equal(nativeRichImageReview({value:'#UNKNOWN!',text:'#UNKNOWN!'},{imagePresent:true}),false);
 assert.equal(nativeRichImageReview({value:-2146826273,text:'#VALUE!'},{imagePresent:true}),false);
});

test('native worksheet FilterMode is not compared with a separate table AutoFilter', async()=>{
 const previous=globalThis.tabula;
 const sheet={name:'Sheet1',filter:null,tables:[{name:'Table1',r1:0,c1:0,r2:2,c2:0,header:true,filter:{criteria:{0:{op:'eq',value:1}}}}]};
 const workbook={sheets:[sheet],ownSheetCount:()=>1};
 const native={sheets:[{name:'Sheet1',sample:[],autoFilterMode:false,filterMode:false,tables:[],pivots:[]}]};
 try{
  globalThis.tabula={wb:()=>workbook};
  const page={evaluate:(fn,arg)=>fn(arg)};
  let result=await auditNativeSamples(page,native);
  assert.deepEqual(result.nativeOptionSamples.differences,[],'active table must not mark worksheet AutoFilter active');
  sheet.filter={r1:0,c1:2,r2:2,c2:2,criteria:{2:{op:'eq',value:1}}};
  result=await auditNativeSamples(page,native);
  assert.ok(result.nativeOptionSamples.differences.some(x=>x.field==='autoFilter.active'&&x.native===false&&x.wixel===true),'real worksheet filter difference is still detected');
 }finally{if(previous===undefined)delete globalThis.tabula;else globalThis.tabula=previous;}
});
