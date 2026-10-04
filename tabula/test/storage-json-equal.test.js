import test from 'node:test';
import assert from 'node:assert/strict';
import {storedJsonEqual,currentSheetMetadata} from '../src/storage-json-equal.js';
const compare=(a,b,expected)=>{assert.equal(storedJsonEqual(a,b),JSON.stringify(a)===JSON.stringify(b));if(expected!==undefined)assert.equal(storedJsonEqual(a,b),expected);};

test('metadata comparison preserves JSON primitive, omission, array and typed-array semantics',()=>{
  for(const[a,b]of[[NaN,null],[Infinity,null],[-Infinity,null],[-0,0],[undefined,undefined],[{}, {x:undefined}],[[undefined,NaN,Infinity],[null,null,null]],[[,2],[null,2]],[new Uint8Array([1,2]),{'0':1,'1':2}],[new Float64Array([NaN,2]),{'0':null,'1':2}],[new Date('2026-01-02T03:04:05Z'),'2026-01-02T03:04:05.000Z'],[new Date('invalid'),null],[new Number(2),2],[new String('x'),'x'],[new Boolean(false),false],[{a:[{x:true}]},{a:[{x:true}]}],[{a:1},{a:2}],[[1],[1,2]]])compare(a,b);
});
test('metadata key order follows JSON and non-enumerable or symbol properties are ignored',()=>{
  compare({a:1,b:2},{b:2,a:1},false);
  const a={a:1},b={a:1};Object.defineProperty(a,'hidden',{value:2});a[Symbol('ignored')]=3;compare(a,b,true);
  compare(Object.assign(Object.create(null),{a:1}),{a:1},true);
  compare({x:undefined,a:1},{a:1,x:undefined},true);
});
test('only root state is ignored for state-only sheet reuse',()=>{
  const a={name:'x',state:'hidden',page:{state:'inside'},images:[]},b={name:'x',state:undefined,page:{state:'inside'},images:[]};
  assert.equal(storedJsonEqual(a,b,{omitRoot:['state']}),true);
  b.page.state='changed';assert.equal(storedJsonEqual(a,b,{omitRoot:['state']}),false);
});
test('unserializable cycles and bigint cannot silently become comparable metadata',()=>{
  const cycle={};cycle.self=cycle;assert.throws(()=>storedJsonEqual(cycle,cycle),TypeError);
  assert.throws(()=>storedJsonEqual({x:1n},{x:1n}),TypeError);
});
test('current sheet metadata reads captured fields and detects newly enabled fileValues',()=>{
  const source={name:'x',images:[{src:'data:unchanged'}],page:{margins:{top:1}},fileValues:false,cells:new Map(),_private:1};
  const captured={name:source.name,images:structuredClone(source.images),page:structuredClone(source.page)};
  const view=currentSheetMetadata(source,captured);assert.equal(view.images,source.images);assert.equal(view.page,source.page);assert.equal(view.cells,undefined);assert.equal(storedJsonEqual(view,captured),true);
  source.fileValues=true;assert.equal(storedJsonEqual(currentSheetMetadata(source,captured),captured),false);
  source.fileValues=false;source.images[0].src+='changed';assert.equal(storedJsonEqual(currentSheetMetadata(source,captured),captured),false);
});
