import { test } from 'node:test';
import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import { collectRibbonControls, createRibbonKeytipRegistry, auditRibbonKeytips } from '../src/ribbon-keytips.js';

// Evaluate only the actual TABS declaration, without importing the DOM renderer.
const ribbonSource = readFileSync(new URL('../src/ribbon.js', import.meta.url), 'utf8');
const tabs = new Function('NUMBER_FORMATS', ribbonSource.slice(ribbonSource.indexOf('export const FONTS'), ribbonSource.indexOf('export function ribbonCommands')).replaceAll('export const ', 'const ') + '; return TABS;')([]);
const appSource = readFileSync(new URL('../src/app.js', import.meta.url), 'utf8');
const legacy = new Function(`return ${appSource.match(/const KEYTIPS = (\{[\s\S]*?\r?\n\});/)[1]}`)();
const fixture = (items, launcher) => [{ id: 'home', label: '홈', groups: [{ label: '합성', items, launcher }] }];

test('실제 리본의 모든 탭·실행·메뉴·입력·런처에 고유한 대표 키가 있다', () => {
  const registry = createRibbonKeytipRegistry(tabs, legacy);
  const { controls, entries, audit } = registry;
  assert.equal(controls.length, 297);
  assert.equal(controls.filter(c => c.chrome).length, 15);
  assert.equal(audit.total, audit.covered);
  assert.deepEqual(audit.missingControlIds, []);
  assert.deepEqual(audit.duplicatePaths, []);
  assert.deepEqual(audit.prefixConflicts, []);
  assert.deepEqual(audit.invalidEntries, []);
  assert.equal(new Set(controls.map(c => c.id)).size, controls.length);
  assert.equal(new Set(controls.map(c => c.path)).size, controls.length);
  for (const tab of tabs.filter(t => !t.file)) assert.ok(Object.values(registry.tabs).includes(tab.id), tab.id);
  for (const control of controls) {
    const primary = entries.filter(e => e.controlId === control.id && e.primary);
    assert.equal(primary.length, 1, control.id);
    assert.equal(primary[0].path, control.path);
  }
  for (const path of Object.keys(legacy)) assert.ok(entries.some(e => e.path === path), `기존 경로 ${path}`);
});

test('셀 스타일·색 메뉴와 글꼴 입력은 실행 명령과 구분된다', () => {
  const { entries } = createRibbonKeytipRegistry(tabs, legacy);
  for (const [path, kind, target] of [
    ['hj', 'menu', 'cellStyles'], ['hh', 'menu', 'fillColor'], ['hfc', 'menu', 'fontColor'],
    ['hff', 'input', 'fontFamily'], ['hfs', 'input', 'fontSize'], ['hfp', 'command', 'painter'], ['nsh', 'menu', 'shapes'],
  ]) {
    const entry = entries.find(e => e.path === path);
    assert.equal(entry.kind, kind); assert.equal(entry.target, target); assert.equal(entry.primary, true);
  }
  assert.equal(entries.find(e => e.path === 'hh').source, 'excel');
  assert.equal(entries.find(e => e.path === 'hj').source, 'user');
  assert.equal(entries.find(e => e.path === 'hfg').source, 'wixel');
  assert.ok(entries.some(e => e.controlId === 'home:command:fillColor' && e.path !== 'hh'), '마지막 색 적용도 별도 조작으로 보존');
});

test('분할 단추·글꼴 목록·메뉴 전용 단추·갤러리·런처를 독립 조작으로 열거한다', () => {
  const font = { type: 'font', cmd: 'fontFamily', menu: 'fontList' };
  const allControls = collectRibbonControls(fixture([
    font, { type: 'color', cmd: 'fillColor', menu: 'fillColor' },
    { type: 'large', cmd: 'paste', menu: 'paste' },
    { type: 'large', cmd: 'cellStyle', menu: 'cellStyles' },
    { type: 'gallery', menu: 'chartStyles' },
  ], 'fontDialog'));
  const controls = allControls.filter(c => !c.chrome);
  assert.equal(allControls.at(-1).target, 'toggleRibbon');
  assert.equal(allControls.at(-1).groupIndex, -1);
  assert.equal(allControls.at(-1).item, null);
  assert.deepEqual(controls.map(c => [c.kind, c.target]), [
    ['input', 'fontFamily'], ['menu', 'fontList'], ['command', 'fillColor'], ['menu', 'fillColor'],
    ['command', 'paste'], ['menu', 'paste'], ['menu', 'cellStyles'], ['menu', 'chartStyles'], ['command', 'fontDialog'],
  ]);
  assert.equal(controls[0].item, font); assert.equal(controls.at(-1).launcher, true); assert.equal(controls.at(-1).item, null);
});

test('한 탭의 반복 조작에도 서로 다른 ID/키를 주고 기존 prefix는 피한다', () => {
  const registry = createRibbonKeytipRegistry(fixture([{ type: 'btn', cmd: 'repeat' }, { type: 'btn', cmd: 'repeat' }]), { hz: ['legacyExtra', '기존'] });
  assert.deepEqual(registry.controls.filter(c => !c.chrome).map(c => c.id), ['home:command:repeat', 'home:command:repeat:2']);
  assert.ok(registry.controls.every(c => !c.path.startsWith('hz')));
  assert.equal(new Set(registry.controls.map(c => c.path)).size, 3);
  assert.deepEqual(registry.audit.prefixConflicts, []);
});

test('누락·잘못된 조작·경로 중복·terminal prefix 충돌을 감사에서 발견한다', () => {
  const controls = collectRibbonControls(fixture([{ type: 'btn', cmd: 'one' }, { type: 'btn', cmd: 'two' }])).filter(c => !c.chrome);
  const entries = [
    { path: 'ha', kind: 'command', target: 'one', tabId: 'home', controlId: controls[0].id },
    { path: 'hab', kind: 'menu', target: 'two', tabId: 'home', controlId: controls[1].id },
    { path: 'ha', kind: 'command', target: 'one', tabId: 'home' },
  ];
  const audit = auditRibbonKeytips(controls, entries, { h: 'home' });
  assert.deepEqual(audit.missingControlIds, [controls[1].id]);
  assert.deepEqual(audit.duplicatePaths, ['ha']); assert.deepEqual(audit.invalidEntries, ['hab']);
  assert.deepEqual(audit.prefixConflicts, [['ha', 'hab']]);
  assert.throws(() => createRibbonKeytipRegistry(fixture([{ type: 'btn', cmd: 'one' }]), { h: ['one', '잘못된 terminal'] }), /누락 또는 충돌|키를 배정하지 못했습니다/);
});

test('등록표 생성은 원본 정의와 기존 별칭을 변경하지 않고 결과도 재생성 가능하다', () => {
  const item = Object.freeze({ type: 'btn', cmd: 'one', title: '하나' });
  const source = fixture(Object.freeze([item])); Object.freeze(source[0].groups[0]);
  const aliases = Object.freeze({ ha: Object.freeze(['one', '하나']) });
  const a = createRibbonKeytipRegistry(source, aliases), b = createRibbonKeytipRegistry(source, aliases);
  assert.deepEqual(a.entries, b.entries);
  a.entries[0].label = '수정'; assert.equal(b.entries[0].label, '하나'); assert.equal(aliases.ha[1], '하나');
  assert.throws(() => collectRibbonControls(fixture([{ type: 'unknown', cmd: 'one' }])), /지원하지 않는/);
  assert.throws(() => createRibbonKeytipRegistry([{ id: 'futureTab', groups: [] }]), /탭 키가 없습니다/);
});
