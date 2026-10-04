import test from 'node:test';
import assert from 'node:assert/strict';
import {createStoredFormulaMemo} from '../src/stored-formula-memo.js';

test('stored formula memo reuses only exact text and leaves non-formulas unchanged',()=>{
 const memo=createStoredFormulaMemo();
 for(const raw of [null,undefined,17,'','=','text',"'=SUM(A1)"])assert.equal(memo.share(raw),raw);
 assert.deepEqual(memo.stats,{entries:0,chars:0,hits:0});
 const formula='=IF("한글😀"="한글😀",SUM(A1),0)';
 assert.equal(memo.share(formula),formula);
 for(let i=0;i<200;i++)assert.equal(memo.share(JSON.parse(JSON.stringify(formula))),formula);
 for(const raw of ['=SUM(A1)','=sum(A1)','= SUM(A1)','=SUM($A$1)'])assert.equal(memo.share(raw),raw);
 assert.equal(memo.stats.entries,5);assert.equal(memo.stats.hits,200);
});

test('stored formula memo enforces both entry and UTF16 character budgets without dropping content',()=>{
 const entries=createStoredFormulaMemo({limit:2,charLimit:100});
 for(const raw of ['=A1','=A2','=A3','=A3','=A1'])assert.equal(entries.share(raw),raw);
 assert.deepEqual(entries.stats,{entries:2,chars:6,hits:1});
 const chars=createStoredFormulaMemo({limit:20,charLimit:8});
 for(const raw of ['="😀"','=A1','=A2','="😀"','=A2'])assert.equal(chars.share(raw),raw);
 assert.deepEqual(chars.stats,{entries:2,chars:8,hits:1});
 const large='="'+ '한😀'.repeat(800000)+'"';assert.equal(chars.share(large),large);assert.deepEqual(chars.stats,{entries:2,chars:8,hits:1});
});

test('stored formula memo default bounds and explicit sheet cleanup release entries',()=>{
 const memo=createStoredFormulaMemo();
 for(let i=0;i<17000;i++)assert.equal(memo.share('='+i),'='+i);
 assert.equal(memo.stats.entries,16384);assert.ok(memo.stats.chars<=2<<20);
 memo.clear();assert.deepEqual(memo.stats,{entries:0,chars:0,hits:0});
 memo.share('=1');assert.equal(memo.stats.entries,1);
 const other=createStoredFormulaMemo();other.share('=1');assert.equal(other.stats.hits,0);assert.equal(memo.stats.hits,0);
});
