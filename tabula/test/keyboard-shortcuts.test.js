import { test } from 'node:test';
import assert from 'node:assert/strict';
import { shortcutCode, appleTouchDevice, repeatFunctionKey, browserStructureCommand, excelDirectCommand } from '../src/keyboard-shortcuts.js';

test('physical shortcuts keep their position across Korean and Apple Option input', () => {
  for (const key of ['c', 'C', 'ㅊ', 'ç', 'Unidentified']) assert.equal(shortcutCode({ key, code: 'KeyC', keyCode: 67 }), 'KeyC');
  assert.equal(shortcutCode({ key: 'a', code: 'KeyQ', keyCode: 65 }), 'KeyQ');
  assert.equal(shortcutCode({ key: '@', code: 'Digit2' }), 'Digit2');
});
test('missing and unidentified physical codes use letters then legacy key codes', () => {
  for (const code of ['', 'Unidentified', undefined]) {
    assert.equal(shortcutCode({ key: 'C', code }), 'KeyC');
    assert.equal(shortcutCode({ key: 'ㅊ', code, keyCode: 67 }), 'KeyC');
    assert.equal(shortcutCode({ key: 'Unidentified', code, which: 86 }), 'KeyV');
    assert.equal(shortcutCode({ key: '3', code }), 'Digit3');
  }
  assert.equal(shortcutCode({ key: 'ㅊ' }), '');
  assert.equal(shortcutCode({ key: 'Control', keyCode: 17 }), '');
});
test('composition and dead keys never become spreadsheet shortcuts', () => {
  for (const extra of [{ isComposing: true }, { keyCode: 229 }, { key: 'Process' }, { key: 'Dead' }])
    assert.equal(shortcutCode({ key: 'c', code: 'KeyC', keyCode: 67, ...extra }), '');
});
test('iPad desktop user agent is independent of viewport and pointer settings', () => {
  assert.equal(appleTouchDevice({ userAgent: 'iPad', platform: 'iPad', maxTouchPoints: 5 }), true);
  assert.equal(appleTouchDevice({ userAgent: 'iPhone', maxTouchPoints: 5 }), true);
  assert.equal(appleTouchDevice({ userAgent: 'Macintosh', platform: 'MacIntel', maxTouchPoints: 5 }), true);
  assert.equal(appleTouchDevice({ userAgent: 'Macintosh', platform: 'MacIntel', maxTouchPoints: 0 }), false);
  assert.equal(appleTouchDevice({ userAgent: 'Windows', platform: 'Win32', maxTouchPoints: 10 }), false);
  assert.equal(appleTouchDevice({ userAgent: 'Android', platform: 'Linux', maxTouchPoints: 5 }), false);
  assert.equal(appleTouchDevice(null), false);
});


test('idle repeat recognizes physical F4 independently of IME key reporting', () => {
  for (const key of ['F4', 'Unidentified', 'Process']) for (const keyCode of [115, 229])
    assert.equal(repeatFunctionKey({ key, code:'F4', keyCode }), true);
  assert.equal(repeatFunctionKey({ key:'F4' }), true);
  assert.equal(repeatFunctionKey({ key:'Process', keyCode:229 }), false);
  for (const extra of [{ isComposing:true }, { ctrlKey:true }, { metaKey:true }, { altKey:true }, { shiftKey:true }, { getModifierState:k=>k==='AltGraph' }])
    assert.equal(repeatFunctionKey({ key:'F4', code:'F4', ...extra }), false);
  for (const key of ['a', 'Process', 'Unidentified']) assert.equal(repeatFunctionKey({ key, code:'KeyA' }), false);
});


test('browser structure menus require Ctrl or Meta plus Shift on physical main and keypad keys', () => {
  for (const modifier of [{ ctrlKey: true }, { metaKey: true }]) {
    for (const [code, command] of [['Equal', 'insertMenuKey'], ['NumpadAdd', 'insertMenuKey'], ['Minus', 'deleteMenuKey'], ['NumpadSubtract', 'deleteMenuKey']]) {
      for (const key of ['+', '=', '-', '_', 'Unidentified', 'ㅂ'])
        assert.equal(browserStructureCommand({ ...modifier, shiftKey: true, code, key }), command);
    }
  }
});

test('ordinary Ctrl or Meta zoom keys are left to the browser, including unshifted keypad keys', () => {
  for (const modifier of [{ ctrlKey: true }, { metaKey: true }]) {
    for (const code of ['Equal', 'Minus', 'NumpadAdd', 'NumpadSubtract', undefined, 'Unidentified'])
      for (const key of ['+', '=', '-', '_']) assert.equal(browserStructureCommand({ ...modifier, code, key }), '');
  }
});

test('structure modifier guards reject Alt, AltGr and combinations without Ctrl or Meta and Shift', () => {
  for (const ctrlKey of [false, true]) for (const metaKey of [false, true]) for (const shiftKey of [false, true]) {
    const expected = (ctrlKey || metaKey) && shiftKey ? 'insertMenuKey' : '';
    assert.equal(browserStructureCommand({ ctrlKey, metaKey, shiftKey, code: 'Equal', key: '+' }), expected);
  }
  for (const modifier of [{ ctrlKey: true }, { metaKey: true }]) {
    for (const extra of [{ altKey: true }, { getModifierState: key => key === 'AltGraph' }])
      for (const code of ['Equal', 'Minus', 'NumpadAdd', 'NumpadSubtract'])
        assert.equal(browserStructureCommand({ ...modifier, shiftKey: true, code, ...extra }), '');
  }
  assert.equal(browserStructureCommand(null), '');
});

test('legacy structure keys use only explicit plus equals minus or underscore when physical code is absent', () => {
  for (const code of [undefined, '', 'Unidentified']) for (const modifier of [{ ctrlKey: true }, { metaKey: true }]) {
    for (const [key, command] of [['+', 'insertMenuKey'], ['=', 'insertMenuKey'], ['-', 'deleteMenuKey'], ['_', 'deleteMenuKey']])
      assert.equal(browserStructureCommand({ ...modifier, shiftKey: true, code, key }), command);
    for (const key of ['Add', 'Subtract', '＋', '−', 'ArrowDown', undefined])
      assert.equal(browserStructureCommand({ ...modifier, shiftKey: true, code, key, keyCode: 187 }), '');
  }
  for (const code of ['KeyA', 'NumpadEqual', 'BracketRight', 'constructor', 'toString', '__proto__'])
    assert.equal(browserStructureCommand({ ctrlKey: true, shiftKey: true, code, key: '+' }), '');
});

test('structure commands never turn composition Process 229 or dead-key text into an insert or delete', () => {
  for (const modifier of [{ ctrlKey: true }, { metaKey: true }]) for (const code of ['Equal', 'Minus', 'NumpadAdd', 'NumpadSubtract', undefined])
    for (const extra of [{ isComposing: true }, { keyCode: 229 }, { which: 229 }, { key: 'Process' }, { key: 'Dead' }])
      assert.equal(browserStructureCommand({ ...modifier, shiftKey: true, code, key: '+', ...extra }), '');
});

test('explicit Ctrl Alt workbook zoom stays independent from browser structure shortcuts', () => {
  for (const modifier of [{ ctrlKey: true }, { metaKey: true }])
    for (const [code, command] of [['Equal', 'zoomIn'], ['NumpadAdd', 'zoomIn'], ['Minus', 'zoomOut'], ['NumpadSubtract', 'zoomOut']]) {
      const event = Object.freeze({ ...modifier, altKey: true, shiftKey: false, code });
      assert.equal(browserStructureCommand(event), '');
      assert.equal(excelDirectCommand(event), command);
    }
});
