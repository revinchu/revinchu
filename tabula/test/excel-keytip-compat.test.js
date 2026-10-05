import test from 'node:test';
import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import { EXCEL_KEYTIP_COMPAT, EXCEL_LEGACY_GROUPS, lookupExcelKeytip, auditExcelKeytipCompat } from '../src/excel-keytip-compat.js';

import { createRibbonKeytipRegistry } from '../src/ribbon-keytips.js';

const source = readFileSync(new URL('../src/app.js', import.meta.url), 'utf8');
const commandStart = source.indexOf('const COMMANDS = {');
const commands = source.slice(commandStart, source.indexOf("window.addEventListener('error'", commandStart));
const menus = source.slice(source.indexOf('const MENUS = {'), commandStart);
const entry = path => EXCEL_KEYTIP_COMPAT.find(x => x.path === path);

test('이전 Excel 단축키는 실제 명령 또는 메뉴에 연결되고 출처가 있다', () => {
  assert.equal(EXCEL_KEYTIP_COMPAT.filter(x => x.source === 'excel-legacy').length, 91);
  assert.equal(EXCEL_KEYTIP_COMPAT.filter(x => x.source === 'excel').length, 3);
  for (const e of EXCEL_KEYTIP_COMPAT) {
    assert.ok(Object.isFrozen(e), e.path);
    assert.ok(/^https:\/\/(?:learn|support)\.microsoft\.com\//.test(e.sourceUrl), e.path);
    assert.ok(new RegExp('\\b' + e.target + '\\s*[:,]').test(e.kind === 'menu' ? menus : commands), e.path + ' → ' + e.target);
    if (e.source === 'excel-legacy') {
      assert.ok(e.evidence.controlId > 0, e.path);
      assert.ok(e.evidence.caption.includes('&'), e.path);
      assert.equal(e.evidence.excelVersion, '16.0');
      assert.equal(e.evidence.lcid, 1042);
      assert.ok(EXCEL_LEGACY_GROUPS[e.path[0]], e.path);
    }
  }
});

test('필터 구 경로는 자동 필터·모두 표시·고급 필터를 구별한다', () => {
  for (const [path, target, id] of [['dff', 'toggleFilter', 899], ['dfs', 'clearFilter', 900], ['dfa', 'advancedFilter', 901]]) {
    assert.equal(entry(path).target, target);
    assert.equal(entry(path).evidence.controlId, id);
  }
  assert.equal(lookupExcelKeytip('d').entry, null);
  assert.equal(lookupExcelKeytip('df').entry, null);
  assert.deepEqual(lookupExcelKeytip('df').next, ['a', 'f', 's']);
  assert.equal(lookupExcelKeytip('DFF').entry.target, 'toggleFilter');
});

test('입력 접두어는 완성 전 명령을 실행하거나 이전 명령으로 되돌아가지 않는다', () => {
  assert.equal(lookupExcelKeytip('oc').entry, null);
  assert.deepEqual(lookupExcelKeytip('oc').next, ['a', 'h', 'u', 'w']);
  for (const path of ['dfff', 'ezz', 'df?', 'Alt+D+F+F']) assert.deepEqual(lookupExcelKeytip(path), { entry: null, matches: [], next: [] });
  assert.equal(lookupExcelKeytip('').matches.length, EXCEL_KEYTIP_COMPAT.length);
});

test('비슷한 이름의 다른 기능에 경로를 연결하지 않는다', () => {
  assert.equal(entry('dt').target, 'dataTable');
  assert.equal(entry('dt').evidence.controlId, 862);
  assert.equal(entry('dic').target, 'createTable');
  assert.equal(entry('dic').evidence.controlId, 7193);
  assert.equal(entry('ipa').kind, 'menu');
  assert.equal(entry('ipa').target, 'shapes');
  assert.equal(entry('eah').target, 'clearHyperlinks');
  assert.equal(entry('ear').target, 'removeHyperlink');
  for (const unsupported of ['dp', 'do', 'eis', 'dga', 'dge', 'tmv', 'ocs']) assert.equal(entry(unsupported), undefined, unsupported);
});

test('현재 공식 리본 경로와 구 Excel 경로를 분리한다', () => {
  for (const [path, target] of [['hsu', 'sortDialog'], ['ay3', 'reapplyFilter'], ['hor', 'renameSheet']]) {
    assert.equal(entry(path).target, target);
    assert.equal(entry(path).source, 'excel');
  }
  assert.equal(entry('ohr').target, entry('hor').target);
  assert.equal(entry('ds').target, entry('hsu').target);
});

test('호환 경로 자체에 중복·접두어 충돌·형식 오류가 없다', () => {
  assert.deepEqual(auditExcelKeytipCompat(), { duplicatePaths: [], prefixConflicts: [], targetConflicts: [], invalidEntries: [] });
});

test('기존 단축 경로 ay가 공식 ay3 입력을 가로채는 충돌을 탐지한다', () => {
  const old = [{ path: 'ay', kind: 'command', target: 'reapplyFilter' }];
  assert.deepEqual(auditExcelKeytipCompat(EXCEL_KEYTIP_COMPAT, old).prefixConflicts, [['ay', 'ay3']]);
});

test('같은 동작의 기존 별칭은 허용하지만 다른 동작으로 덮어쓰지 않는다', () => {
  assert.deepEqual(auditExcelKeytipCompat(EXCEL_KEYTIP_COMPAT, [{ ...entry('es') }]).targetConflicts, []);
  assert.deepEqual(auditExcelKeytipCompat(EXCEL_KEYTIP_COMPAT, [{ ...entry('es'), target: 'paste' }]).targetConflicts, ['es']);
  assert.deepEqual(auditExcelKeytipCompat([entry('es'), entry('es')]).duplicatePaths, ['es']);
  assert.deepEqual(auditExcelKeytipCompat([null, { path: '!bad' }]).invalidEntries, [null, '!bad']);
});


test('실제 리본·기존 경로와 94개 호환 경로를 병합해도 모든 조작의 대표 키가 유지된다', () => {
  const ribbonSource = readFileSync(new URL('../src/ribbon.js', import.meta.url), 'utf8');
  const tabs = new Function('NUMBER_FORMATS', ribbonSource.slice(ribbonSource.indexOf('export const FONTS'), ribbonSource.indexOf('export function ribbonCommands')).replaceAll('export const ', 'const ') + '; return TABS;')([]);
  const old = new Function(`return ${source.match(/const KEYTIPS = (\{[\s\S]*?\r?\n\});/)[1]}`)();
  const registry = createRibbonKeytipRegistry(tabs, old, EXCEL_KEYTIP_COMPAT);
  assert.equal(registry.audit.total, registry.audit.covered);
  assert.deepEqual(registry.audit.prefixConflicts, []);
  assert.deepEqual(registry.audit.duplicatePaths, []);
  assert.deepEqual(registry.audit.invalidEntries, []);
  for (const control of registry.controls) {
    const tabPrefix = Object.entries(registry.tabs).find(([, tabId]) => tabId === control.tabId)?.[0];
    const primary = registry.entries.find(e => e.path === control.path);
    assert.ok(tabPrefix && control.path.startsWith(tabPrefix), control.id + ' → ' + control.path);
    assert.equal(primary?.primary, true, control.id);
    assert.notEqual(primary?.source, 'excel-legacy', control.id);
  }
  const audit = auditExcelKeytipCompat(EXCEL_KEYTIP_COMPAT, registry.entries);
  assert.deepEqual(audit.prefixConflicts, []);
  assert.deepEqual(audit.targetConflicts, []);
  for (const compatible of EXCEL_KEYTIP_COMPAT) {
    const actual = registry.entries.find(e => e.path === compatible.path);
    assert.equal(actual?.target, compatible.target, compatible.path);
    assert.equal(actual?.kind, compatible.kind, compatible.path);
  }
  assert.equal(entry('er').target, 'redoOrRepeat');
  assert.equal(entry('oht').target, 'sheetTabColor');
  assert.equal(entry('ta').target, 'autoCorrectOptions');
  assert.equal(entry('vz').target, 'zoomDialog');
});
