import test from 'node:test';
import assert from 'node:assert/strict';
import { CsvStream, CsvBlockReader, parseDelimited, guessDelimiter } from '../src/csv.js';
import { blockValue } from '../src/block.js';
function stream(text, chunks) { const rows=[]; const p=new CsvStream(',',row=>rows.push(row)); let i=0; for(const size of chunks){p.push(text.slice(i,i+size));i+=size;}if(i<text.length)p.push(text.slice(i));p.end();return rows; }
function splitEvery(text,size) {return Array(Math.ceil(text.length/size)).fill(size);}
const cases=['','\n','\r','\r\n','""','h,v','h,v\n','h,v\n\n','h,v\n\nB,2\n','h,v\rA,1\rB,2\r','h,v\r\n\r\nB,2\r\n','h,v\rA,1\nB,2\r\nC,3','h,v\n"line1\r\nline2",1\n','h,v\n"line1\rline2",1\n','"a""b",c\r\n"",2\r\n','"a\n\n끝",b\n','a"b,c\n','"closed"tail,2','"unfinished,2','a,b,','"",','\n\n'];
test('streamed CSV preserves empty rows, CR/CRLF and terminal record semantics of small files',()=>{for(const text of cases)for(const n of [1,2,3,7,65536])assert.deepEqual(stream(text,splitEvery(text,n)),parseDelimited(text),JSON.stringify({text,n}));});
test('every two-chunk split preserves quoted escapes, CRLF and Korean values',()=>{const text='이름,메모\r\n\r\n"서울""지점","첫째\r\n둘째\r셋째\n끝"\r\nB,2\r\n\r\n';const expected=parseDelimited(text);for(let at=0;at<=text.length;at++)assert.deepEqual(stream(text,[at,text.length-at]),expected,'split '+at);});
test('empty pushes before and between data do not lose BOM or split CRLF/quotes',()=>{const rows=[],p=new CsvStream(',',r=>rows.push(r));for(const s of ['', '\ufeffh,v\r','','\n"a"','','"b",1\r','','\n\r','','\n'])p.push(s);p.end();assert.deepEqual(rows,[['h','v'],['a"b','1'],['']]);});
test('completed rows are emitted before EOF while an unfinished quoted record spans chunks',()=>{const rows=[],p=new CsvStream(',',r=>rows.push(r));p.push('h,v\r');assert.deepEqual(rows,[['h','v']]);p.push('\n\r');assert.deepEqual(rows,[['h','v'],['']]);p.push('\n"a');p.push('\r');p.push('\nb",2');assert.equal(rows.length,2);p.end();assert.deepEqual(rows,[['h','v'],[''],['a\nb','2']]);});
test('block CSV preserves blank row offsets and multiline strings',()=>{const p=new CsvBlockReader(',',4);for(const s of ['h,v\r','\n\r','\n"a\r','\nb",2\r','\n\r','\nZ,9\r','\n'])p.push(s);const {header,block}=p.finish();assert.deepEqual(header,['h','v']);assert.equal(block.n,4);assert.equal(blockValue(block,1,0),null);assert.equal(blockValue(block,2,0),'a\nb');assert.equal(blockValue(block,2,1),2);assert.equal(blockValue(block,3,0),null);assert.equal(blockValue(block,4,0),'Z');assert.equal(blockValue(block,4,1),9);});
test('delimiter inference uses the first record for CR, LF and CRLF headers',()=>{for(const eol of ['\r','\n','\r\n'])assert.equal(guessDelimiter('name\tvalue'+eol+'a,b,c,d,e,f'), '\t');});
test('seeded mixed-newline quoted matrices match the small parser across streaming chunks',()=>{let seed=417;const random=()=>((seed=(seed*1664525+1013904223)>>>0)/4294967296);for(let n=0;n<100;n++){let text='h,v\r\n';for(let r=0;r<20;r++){const eol=['\r','\n','\r\n'][Math.floor(random()*3)];if(random()<.25){text+=eol;continue;}const value=['한글','a,b','say "hello"','a\rb','a\r\nb','a\nb',''][Math.floor(random()*7)];text+='"'+value.replace(/"/g,'""')+'",'+r+eol;}const sizes=[];for(let i=0;i<text.length;){const s=1+Math.floor(random()*13);sizes.push(s);i+=s;}assert.deepEqual(stream(text,sizes),parseDelimited(text),'matrix '+n);}});

test('seeded arbitrary malformed CSV preserves the permissive small-parser semantics',()=>{
 let seed=991733;const random=()=>((seed=(seed*1664525+1013904223)>>>0)/4294967296), alphabet=['a',',','"','\r','\n'];
 for(let run=0;run<5000;run++){
  const n=Math.floor(random()*120);let text='';for(let i=0;i<n;i++)text+=alphabet[Math.floor(random()*alphabet.length)];
  const chunks=[];for(let i=0;i<text.length;){const size=1+Math.floor(random()*17);chunks.push(size);i+=size;}
  assert.deepEqual(stream(text,chunks),parseDelimited(text),'case '+run+' '+JSON.stringify(text));
 }
});
