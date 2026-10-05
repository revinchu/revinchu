import test from 'node:test';
import assert from 'node:assert/strict';
import {shapeTextParagraphs,shapeTextString,replaceShapeText,formatShapeText,shapeSelectionStyle} from '../src/shape-text-editor.js';

test('plain shape text keeps consecutive and trailing blank paragraphs',()=>{
 const ps=shapeTextParagraphs({text:'가\r\n\r\n나\n'});
 assert.equal(ps.length,4);assert.equal(shapeTextString(ps),'가\n\n나\n');assert.deepEqual(ps[1].runs,[]);
});
test('rich paragraph copy preserves false overrides, blank size and unknown run attributes',()=>{
 const shape={paras:[{align:'right',sz:23,runs:[]},{runs:[{t:'가😀',b:false,i:true,hlink:'https://example.invalid/'}]}]};
 const ps=shapeTextParagraphs(shape);assert.deepEqual(ps,shape.paras);ps[1].runs[0].b=true;assert.equal(shape.paras[1].runs[0].b,false);
});
test('partial formatting splits only selected run and honors explicit false',()=>{
 const ps=[{align:'center',runs:[{t:'가나다라',b:true,i:true,color:'#112233'}]}];
 const out=formatShapeText(ps,{start:1,end:3},{bold:false,color:'#aabbcc'});
 assert.deepEqual(out[0].runs,[{t:'가',b:true,i:true,color:'#112233'},{t:'나다',b:false,i:true,color:'#aabbcc'},{t:'라',b:true,i:true,color:'#112233'}]);
 assert.equal(ps[0].runs[0].t,'가나다라');
});
test('range formatting preserves unselected paragraphs and excludes end paragraph boundary',()=>{
 const ps=[{align:'left',runs:[{t:'앞'}]},{align:'right',runs:[{t:'뒤'}]}];
 const out=formatShapeText(ps,{start:0,end:2},{align:'center',underline:true});
 assert.equal(out[0].align,'center');assert.equal(out[0].runs[0].u,true);assert.deepEqual(out[1],ps[1]);
});
test('empty paragraph supports pending font and false overrides without placeholder characters',()=>{
 const out=formatShapeText([{sz:24,runs:[]}],{start:0,end:0},{size:18,bold:false});
 assert.deepEqual(out,[{sz:24,runs:[{t:'',sz:18,b:false}]}]);assert.equal(shapeTextString(out),'');
});
test('replacing a cross-paragraph selection keeps prefix/suffix formatting and paragraph metadata',()=>{
 const ps=[{align:'left',runs:[{t:'앞문장',b:true}]},{align:'right',runs:[{t:'뒷문장',i:true}]}];
 const out=replaceShapeText(ps,{start:1,end:5},'새\n중\n끝',{u:true,color:'#ff0000'});
 assert.equal(shapeTextString(out),'앞새\n중\n끝문장');assert.equal(out[0].align,'left');assert.equal(out[2].align,'right');
 assert.deepEqual(out[0].runs[0],{t:'앞',b:true});assert.deepEqual(out[2].runs[1],{t:'문장',i:true});
});
test('delete across line boundaries joins the surviving runs and restores one empty paragraph',()=>{
 const ps=[{runs:[{t:'가',b:true}]},{runs:[]},{runs:[{t:'나',i:true}]}];
 assert.deepEqual(replaceShapeText(ps,{start:1,end:3},''),[{runs:[{t:'가',b:true},{t:'나',i:true}]}]);
 assert.deepEqual(replaceShapeText(ps,{start:0,end:4},''),[{runs:[]}]);
});
test('text insertion preserves unicode surrogate pairs, literal markup and line endings',()=>{
 const ps=[{runs:[{t:'가😀나',b:true}]}];
 const out=replaceShapeText(ps,{start:3,end:3},'<img>\r\n한',{b:false});
 assert.equal(shapeTextString(out),'가😀<img>\n한나');assert.equal(out[0].runs[0].t,'가😀');
});
test('selection style merges inherited values and reports mixed properties without losing false',()=>{
 const shape={bold:true,size:20,color:'#123456'},ps=[{align:'right',sz:30,runs:[{t:'가',b:false},{t:'나',b:true,i:true}]}];
 const one=shapeSelectionStyle(shape,ps,{start:0,end:1});assert.equal(one.bold,false);assert.equal(one.size,30);assert.equal(one.align,'right');
 const mixed=shapeSelectionStyle(shape,ps,{start:0,end:2});assert.equal(mixed.bold,undefined);assert.equal(mixed.italic,undefined);assert.equal(mixed.color,'#123456');
});
test('collapsed caret chooses preceding run except at start of next paragraph',()=>{
 const ps=[{runs:[{t:'가',b:true},{t:'나',b:false}]},{runs:[{t:'다',i:true}]}];
 assert.equal(shapeSelectionStyle({},ps,{start:1,end:1}).bold,true);
 assert.equal(shapeSelectionStyle({},ps,{start:3,end:3}).italic,true);
});

test('explicit empty rich paragraphs do not revive an outdated plain text cache',()=>{assert.deepEqual(shapeTextParagraphs({text:'이전 내용',paras:[]}),[{runs:[]}]);});
