import test from 'node:test';
import assert from 'node:assert/strict';
import { Workbook } from '../src/workbook.js';

test('문서 교체 시 이전 문서의 비동기 의존 그래프를 중단하고 값과 재시작을 보존한다', async () => {
  const cells = new Map();
  for (let r=0;r<30000;r++) { cells.set(`${r},0`,{raw:String(r+1)});cells.set(`${r},1`,{raw:`=A${r+1}*2`,cached:(r+1)*2}); }
  const book=new Workbook({sheets:[{name:'합성',cells}]});
  const running=book.prepareGraph(-1);
  book.cancelGraphPreparation();
  assert.equal(await running,false);
  assert.equal(book.graph,null);
  assert.equal(book.graphRun,null);
  assert.equal(book.getValue(0,29999,1),60000);
  assert.equal(await book.prepareGraph(),true);
  book.transact(()=>book.setInput(0,29999,0,'9'));
  assert.equal(book.getValue(0,29999,1),18);
});

test('완성된 의존 그래프는 문서 전환 준비 취소로 버리지 않는다', async () => {
  const book=new Workbook({sheets:[{name:'합성',cells:{'0,0':{raw:'3'},'0,1':{raw:'=A1*2'}}}]});
  assert.equal(await book.prepareGraph(),true);const graph=book.graph;
  book.cancelGraphPreparation();assert.equal(book.graph,graph);
  assert.equal(await book.prepareGraph(),true);assert.equal(book.graph,graph);
});
