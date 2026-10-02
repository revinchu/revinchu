import { test } from 'node:test';
import assert from 'node:assert/strict';
import { resolveImportRangeUrl } from '../src/fx-web.js';

const ID = 'WIXEL_SYNTHETIC_PUBLIC_SHEET_123456';
const OTHER = 'WIXEL_SYNTHETIC_OTHER_SHEET_654321';
const base = `https://docs.google.com/spreadsheets/d/${ID}`;
const source = sheet => ({ url: `${base}/export?format=csv&range=A1%3AC8&gid=0`, tabsUrl: `${base}/htmlview`, sheet });
const item = (name, gid, pageUrl = `${base}/htmlview/sheet?headers=true&gid=${gid}`) => `items.push({name:${JSON.stringify(name)},pageUrl:${JSON.stringify(pageUrl)},gid:${JSON.stringify(gid)},initialSheet:(${JSON.stringify(gid)} == gid)});`;
const html = (...items) => `<html><script>function init(){var items = [];${items.join('')}var sheetsViewport = null;}</script></html>`;

test('Google 탭 이름을 정확히 찾아 gid만 교체하며 원본 source는 바꾸지 않는다', () => {
  const s = source('두 번째 & 탭'), before = structuredClone(s);
  const u = new URL(resolveImportRangeUrl(s, html(item('첫 탭', '0'), item('두 번째 & 탭', '8726'))));
  assert.equal(u.searchParams.get('gid'), '8726'); assert.equal(u.searchParams.get('range'), 'A1:C8');
  assert.equal(u.searchParams.get('format'), 'csv'); assert.deepEqual(s, before);
});

test('HTML 메타데이터의 hex·unicode·따옴표·역슬래시·slash escape를 실행 없이 읽는다', () => {
  const name = "매출 & '월별' \\ 분기/한글";
  const raw = item(name, '17').replace('매출 &', '\\uB9E4\\uCD9C \\x26').replaceAll('https://', 'https:\\/\\/').replaceAll("'월별'", "\\'월별\\'");
  assert.equal(new URL(resolveImportRangeUrl(source(name), html(raw))).searchParams.get('gid'), '17');
  const literal = String.raw`items.push({name:"literal \\x41",pageUrl:"${base}/htmlview/sheet?gid=5",gid:"5"});`;
  assert.equal(new URL(resolveImportRangeUrl(source('literal \\x41'), html(literal))).searchParams.get('gid'), '5');
});

test('이름은 공백·대소문자까지 정확히 일치하며 누락·HTML 오류에 첫 탭을 대신 쓰지 않는다', () => {
  const meta = html(item(' Sheet ', '0'), item('Sheet', '1'));
  assert.equal(new URL(resolveImportRangeUrl(source(' Sheet '), meta)).searchParams.get('gid'), '0');
  for (const name of ['sheet', '없는 시트']) assert.throws(() => resolveImportRangeUrl(source(name), meta), /시트 이름/);
  assert.throws(() => resolveImportRangeUrl(source('Sheet'), '<html>로그인이 필요합니다</html>'), /시트 이름/);
  assert.throws(() => resolveImportRangeUrl(source('Sheet'), `<p>var items=[];${item('Sheet', '1')}</p>`), /시트 이름/);
});

test('중복 이름·탭 번호는 모호한 목록으로 거부한다', () => {
  assert.throws(() => resolveImportRangeUrl(source('a'), html(item('a', '1'), item('a', '2'))), /중복/);
  assert.throws(() => resolveImportRangeUrl(source('a'), html(item('a', '1'), item('b', '1'))), /중복/);
});

test('다른 문서·다른 호스트·HTTP·자격 증명·다른 gid·중복 gid를 거부한다', () => {
  for (const url of [
    `${base.replace(ID, OTHER)}/htmlview/sheet?gid=1`,
    `https://docs.google.com.evil.invalid/spreadsheets/d/${ID}/htmlview/sheet?gid=1`,
    `${base.replace('https:', 'http:')}/htmlview/sheet?gid=1`,
    `${base.replace('https://', 'https://user@')}/htmlview/sheet?gid=1`,
    `${base}/htmlview/sheet?gid=2`, `${base}/htmlview/sheet?gid=1&gid=2`, `${base}/export?gid=1`,
  ]) assert.throws(() => resolveImportRangeUrl(source('a'), html(item('a', '1', url))), /주소|일치/);
  assert.throws(() => resolveImportRangeUrl({ ...source('a'), tabsUrl: `https://docs.google.com/spreadsheets/d/${OTHER}/htmlview` }, html(item('a', '1'))), /주소/);
});

test('잘못된 escape·제어 문자·표현식 이름을 거부하고 외부 코드를 실행하지 않는다', () => {
  const valid = item('a', '1');
  for (const raw of [valid.replace('name:"a"', String.raw`name:"\xZZ"`), valid.replace('name:"a"', String.raw`name:"\u00ZZ"`), valid.replace('name:"a"', String.raw`name:"\q"`), valid.replace('name:"a"', String.raw`name:"\u0000"`), valid.replace('name:"a"', 'name:(globalThis.__wixelTabsExecuted=true,"a")')]) {
    globalThis.__wixelTabsExecuted = false;
    assert.throws(() => resolveImportRangeUrl(source('a'), html(raw)), /시트|문자|형식/);
    assert.equal(globalThis.__wixelTabsExecuted, false);
  }
  delete globalThis.__wixelTabsExecuted;
});

test('명시한 시트가 없으면 기존 주소를 그대로 사용한다', () => {
  const s = { url: `${base}/export?format=csv&gid=8&range=B2:C5` };
  assert.equal(resolveImportRangeUrl(s, ''), s.url);
});
