import { test } from 'node:test';
import assert from 'node:assert/strict';
import { Workbook } from '../src/workbook.js';
import { NET, LOADING, parseMarkup, evalPath, textOf, parseQuery, runQuery, parseFeed, parseTicker, parseHistory, parseQuote, bahtText } from '../src/fx-web.js';
import { fromFileFormula } from '../src/xlfn.js';
import { CellImage } from '../src/fxcore.js';

const book = (cells) => {
  const wb = new Workbook();
  wb.transact(() => { for (const [a, raw] of Object.entries(cells)) { const c = a.charCodeAt(0) - 65; const r = Number(a.slice(1)) - 1; wb.setCellData(0, r, c, { raw: String(raw) }); } });
  return wb;
};
const val = (wb, a) => { wb.ensureSpills?.(); return wb.getValue(0, Number(a.slice(1)) - 1, a.charCodeAt(0) - 65); };

const PAGE = `<!doctype html><html><head><title>Test &amp; Page</title></head><body>
<h2 class="t">첫째</h2><h2 class="t big">둘째</h2>
<ul><li>사과<li>배</ul><ol><li>하나</li><li>둘</li></ol>
<table><tr><th>품목</th><th>수량</th></tr><tr><td>A</td><td>1,200</td></tr><tr><td colspan="2">합계</td></tr></table>
<a href="/x">링크1</a><a href="https://y.com">링크2</a><br><img src="a.png"></body></html>`;

test('HTML 파서 · XPath', () => {
  const root = parseMarkup(PAGE);
  assert.deepEqual(evalPath('//h2', [root]).map((n) => textOf(n)), ['첫째', '둘째']);
  assert.deepEqual(evalPath("//h2[contains(@class,'big')]", [root]).map((n) => textOf(n)), ['둘째']);
  assert.deepEqual(evalPath('//a/@href', [root]), ['/x', 'https://y.com']);
  assert.deepEqual(evalPath('//ul/li', [root]).map((n) => textOf(n)), ['사과', '배']);
  assert.deepEqual(evalPath('//ol/li[last()]', [root]).map((n) => textOf(n)), ['둘']);
  assert.equal(textOf(evalPath('//title', [root])[0]), 'Test & Page');
  assert.equal(evalPath('//h2[2]', [root]).length, 1);
});

test('IMPORTHTML · IMPORTXML · IMPORTDATA: 로딩 중 → 받은 뒤 다시 계산', async () => {
  const pages = { 'https://ex.com/p': PAGE, 'https://ex.com/d.csv': 'a,b\n1,"x,y"\n2,z\n' };
  const done = [];
  NET.cache.clear();
  NET.fetcher = async (url) => { if (!(url in pages)) throw new Error('404'); return pages[url]; };
  const wb = book({ A1: '=IMPORTHTML("https://ex.com/p","table",1)', E1: '=IMPORTXML("https://ex.com/p","//h2")', G1: '=IMPORTDATA("https://ex.com/d.csv")', J1: '=IMPORTHTML("https://ex.com/p","list",2)', K1: '=IMPORTXML("https://ex.com/none","//a")' });
  NET.onDone = (sheets) => { done.push(...sheets); for (const si of sheets) wb.invalidate(si); };
  assert.equal(val(wb, 'A1'), LOADING);
  await new Promise((r) => setTimeout(r, 10));
  assert.ok(done.includes(0));
  assert.equal(val(wb, 'A1'), '품목');
  assert.equal(val(wb, 'B2'), 1200);
  assert.equal(val(wb, 'A3'), '합계');
  assert.equal(val(wb, 'E2'), '둘째');
  assert.equal(val(wb, 'H2'), 'x,y');
  assert.equal(val(wb, 'G3'), 2);
  assert.equal(val(wb, 'J2'), '둘');
  assert.equal(val(wb, 'K1')?.code, '#N/A');
  NET.fetcher = null;
  NET.onDone = null;
});

test('QUERY: select · where · group by · order by · pivot · label', () => {
  const wb = book({
    A1: '채널', B1: '월', C1: '비용', A2: '검색', B2: '1월', C2: 100, A3: '디스플레이', B3: '1월', C3: 50, A4: '검색', B4: '2월', C4: 120, A5: '영상', B5: '2월', C5: 30,
    E1: '=QUERY(A1:C5,"select A, sum(C) group by A order by sum(C) desc",1)',
    H1: '=QUERY(A1:C5,"select A, C where C > 40 and A != \'디스플레이\' order by C",1)',
    K1: '=QUERY(A1:C5,"select A, sum(C) group by A pivot B label sum(C) \'합계\'",1)',
  });
  assert.equal(val(wb, 'E1'), '채널');
  assert.equal(val(wb, 'F1'), 'sum 비용');
  assert.equal(val(wb, 'E2'), '검색');
  assert.equal(val(wb, 'F2'), 220);
  assert.equal(val(wb, 'E4'), '영상');
  assert.equal(val(wb, 'H2'), '검색');
  assert.equal(val(wb, 'I2'), 100);
  assert.equal(val(wb, 'I3'), 120);
  assert.equal(val(wb, 'L1'), '1월 합계'); // 명시한 label은 단일 집계 피벗에도 붙음 (Google query language)
  assert.equal(val(wb, 'K2'), '검색');
  const q = parseQuery("select Col1, count(Col2) where Col2 contains 'a' or Col2 is null limit 5 offset 1");
  assert.equal(q.limit, 5);
  assert.equal(q.offset, 1);
  const res = runQuery(parseQuery('select Col1, avg(Col2) group by Col1'), [['x', 1], ['x', 3], ['y', 5]], ['Col1', 'Col2'], ['', '']);
  assert.deepEqual(res.rows, [['x', 2], ['y', 5]]);
});

test('구글 스프레드시트 함수: SPLIT · JOIN · FLATTEN · SORTN · COUNTUNIQUE · 연산자 함수 등', () => {
  const wb = book({
    A1: 3, A2: 1, A3: 3, A4: 2,
    C1: '=SPLIT("a,b;;c",",;")', G1: '=JOIN("-",A1:A4)', H1: '=FLATTEN(A1:A2,A3:A4)', I1: '=SORTN(A1:A4,2)', J1: '=COUNTUNIQUE(A1:A4)',
    K1: '=ADD(2,3)', K2: '=DIVIDE(1,0)', K3: '=GTE(3,3)', K4: '=UNARY_PERCENT(50)', K5: '=ISBETWEEN(5,1,5,TRUE,FALSE)', K6: '=ISEMAIL("a@b.co")', K7: '=ISURL("www.naver.com")',
    L1: '=AVERAGE.WEIGHTED(A1:A2,A3:A4)', L2: '=EPOCHTODATE(86400)', L3: '=COUNTUNIQUEIFS(A1:A4,A1:A4,">1")', L4: '=TO_PURE_NUMBER("50%")', L5: '=ARRAY_CONSTRAIN(A1:A4,2,1)',
    M1: '=ENCODEURL("a b&c")', M2: '=FILTERXML("<r><i>1</i><i>2</i></r>","//i")', M4: '=BINOM.DIST.RANGE(10,0.5,5)', M5: '=ECMA.CEILING(4.1,2)', M6: '=SEARCHB("나","가나다")',
    N1: '=TRIMRANGE(A1:B6)', O1: '=SPARKLINE(A1:A4)',
  });
  assert.deepEqual([val(wb, 'C1'), val(wb, 'D1'), val(wb, 'E1')], ['a', 'b', 'c']);
  assert.equal(val(wb, 'G1'), '3-1-3-2');
  assert.equal(val(wb, 'H4'), 2);
  assert.deepEqual([val(wb, 'I1'), val(wb, 'I2')], [1, 2]);
  assert.equal(val(wb, 'J1'), 3);
  assert.equal(val(wb, 'K1'), 5);
  assert.equal(val(wb, 'K2')?.code, '#DIV/0!');
  assert.equal(val(wb, 'K3'), true);
  assert.equal(val(wb, 'K4'), 0.5);
  assert.equal(val(wb, 'K5'), false);
  assert.equal(val(wb, 'K6'), true);
  assert.equal(val(wb, 'K7'), true);
  assert.equal(val(wb, 'L1'), (3 * 3 + 1 * 2) / 5);
  assert.equal(val(wb, 'L2'), 25570);
  assert.equal(val(wb, 'L3'), 2);
  assert.equal(val(wb, 'L4'), 0.5);
  assert.equal(val(wb, 'L6'), 1);
  assert.equal(val(wb, 'L7'), null);
  assert.equal(val(wb, 'M1'), 'a%20b%26c');
  assert.equal(val(wb, 'M3'), 2);
  assert.ok(Math.abs(val(wb, 'M4') - 0.24609375) < 1e-12);
  assert.equal(val(wb, 'M5'), 6);
  assert.equal(val(wb, 'M6'), 3);
  assert.equal(val(wb, 'N4'), 2);
  assert.equal(val(wb, 'N5'), null);
  assert.ok(val(wb, 'O1') instanceof CellImage);
});

test('피드 · 시세 파서 · BAHTTEXT · 구글 내보내기 수식', () => {
  const f = parseFeed('<?xml version="1.0"?><rss><channel><title>블로그</title><link>https://b.com</link><item><title>글1</title><link>https://b.com/1</link><pubDate>Mon, 01 Jan 2024</pubDate><description>&lt;p&gt;요약&lt;/p&gt;</description></item></channel></rss>');
  assert.equal(f.feed.title, '블로그');
  assert.equal(f.items[0].url, 'https://b.com/1');
  assert.equal(f.items[0].summary, '요약');
  assert.deepEqual(parseTicker('KRX:005930'), { market: 'krx', code: '005930' });
  assert.deepEqual(parseTicker('NASDAQ:GOOG'), { market: 'us', code: 'GOOG' });
  assert.deepEqual(parseTicker('CURRENCY:USDKRW'), { market: 'fx', code: 'USDKRW' });
  const h = parseHistory({ market: 'krx' }, "[['날짜','시가','고가','저가','종가','거래량','외국인소진율'],\n[\"20240102\", 78200, 79800, 78200, 79600, 17142847, 53.3]]");
  assert.equal(h[0][4], 79600);
  const q = parseQuote({ market: 'us', code: 'AAPL' }, 'Symbol,Date,Time,Open,High,Low,Close,Volume,Name\nAAPL.US,2024-01-02,22:00:10,187.15,188.44,183.885,185.64,82488674,APPLE\n');
  assert.equal(q.price, 185.64);
  assert.equal(bahtText(1234.5), 'หนึ่งพันสองร้อยสามสิบสี่บาทห้าสิบสตางค์');
  assert.equal(fromFileFormula('IFERROR(__xludf.DUMMYFUNCTION("IMPORTXML(""https://a.com"",""//h1"")"),"제목")'), 'IMPORTXML("https://a.com","//h1")');
});
