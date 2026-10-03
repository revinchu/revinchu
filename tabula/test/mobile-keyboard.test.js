import { test } from 'node:test';
import assert from 'node:assert/strict';
import { MOBILE_KEYBOARD_KEY, mobileKeyboardPreference, hardwareKeyboardEvent, installMobileKeyboard, prepareKeyboardInput } from '../src/mobile-keyboard.js';

class KeyboardNode {
  constructor(tag = 'input', attrs = {}) {
    this.tagName = tag.toUpperCase(); this.attrs = new Map(Object.entries(attrs)); this.children = []; this.parentElement = null;
    this.isConnected = true; this.value = '한글 값'; this.readOnly = false; this.disabled = false; this.focusCalls = 0; this.inputEvents = 0;
  }
  get type() { return this.attrs.get('type') || 'text'; }
  get isContentEditable() { const v = this.attrs.get('contenteditable'); return v == null ? !!this.parentElement?.isContentEditable : v !== 'false'; }
  hasAttribute(k) { return this.attrs.has(k); }
  getAttribute(k) { return this.attrs.get(k) ?? null; }
  setAttribute(k, v) { this.attrs.set(k, String(v)); }
  removeAttribute(k) { this.attrs.delete(k); }
  append(node) { this.children.push(node); node.parentElement = this; }
  querySelectorAll() {
    const result = [];
    const walk = node => { for (const child of node.children) { if (['INPUT','TEXTAREA'].includes(child.tagName) || child.hasAttribute('contenteditable')) result.push(child); walk(child); } };
    walk(this); return result;
  }
  closest() { for (let n = this; n; n = n.parentElement) if (['INPUT','TEXTAREA','SELECT'].includes(n.tagName)) return n; return null; }
  focus() { this.focusCalls++; }
  dispatchEvent() { this.inputEvents++; }
}
function keyboardEnvironment(t, saved = null) {
  const savedGlobals = new Map(['document', 'localStorage', 'MutationObserver'].map(k => [k, Object.getOwnPropertyDescriptor(globalThis, k)]));
  const listeners = new Map(), values = new Map(saved == null ? [] : [[MOBILE_KEYBOARD_KEY, saved]]);
  const root = new KeyboardNode('html');
  const doc = { documentElement: root, querySelectorAll: (...args) => root.querySelectorAll(...args),
    addEventListener(name, fn, capture) { assert.equal(capture, true); if (!listeners.has(name)) listeners.set(name, new Set()); listeners.get(name).add(fn); },
    removeEventListener(name, fn) { listeners.get(name)?.delete(fn); } };
  const env = { root, doc, listeners, values, observer: null, controllers: [], mobile: true };
  globalThis.document = doc;
  globalThis.localStorage = { getItem: k => values.get(k) ?? null, setItem: (k,v) => values.set(k,v) };
  globalThis.MutationObserver = class {
    constructor(fn) { this.fn = fn; this.disconnected = false; env.observer = this; }
    observe(target, options) { this.target = target; this.options = options; }
    disconnect() { this.disconnected = true; }
  };
  env.add = node => { root.append(node); return node; };
  env.emit = (type, event = {}) => { const e = { target: root, isTrusted: true, preventDefault() { throw Error('정책이 기본 입력을 막았습니다'); }, ...event }; for (const fn of listeners.get(type) || []) fn(e); return e; };
  env.install = onChange => { const c = installMobileKeyboard({ isMobile: () => env.mobile, onChange }); env.controllers.push(c); return c; };
  env.mutate = records => env.observer.fn(records);
  t.after(() => {
    for (const c of env.controllers) c.dispose();
    for (const [k, d] of savedGlobals) { if (d) Object.defineProperty(globalThis, k, d); else delete globalThis[k]; }
  });
  return env;
}

test('keyboard preference accepts only explicit modes and defaults to auto', () => {
  for (const value of [null, undefined, '', 'on', 'Hardware', {}]) assert.equal(mobileKeyboardPreference(value), 'auto');
  for (const value of ['auto','hardware','screen']) assert.equal(mobileKeyboardPreference(value), value);
});

test('auto keyboard detection excludes synthetic, modifier, composition, dead keys and editable targets', () => {
  const target = new KeyboardNode('div');
  for (const key of ['a','가',' ','ArrowDown','Enter','F2','Tab','Escape']) assert.equal(hardwareKeyboardEvent({ isTrusted:true, target, key }), true, key);
  for (const key of ['Shift','Control','Alt','Meta','AltGraph','CapsLock','Dead','Process','Unidentified','']) assert.equal(hardwareKeyboardEvent({ isTrusted:true, target, key }), false, key);
  for (const extra of [{isTrusted:false},{isComposing:true},{keyCode:229},{defaultPrevented:true}]) assert.equal(hardwareKeyboardEvent({ isTrusted:true, target, key:'a', ...extra }), false);
  for (const node of [new KeyboardNode(), new KeyboardNode('textarea'), new KeyboardNode('select'), new KeyboardNode('div',{contenteditable:'true'})]) assert.equal(hardwareKeyboardEvent({isTrusted:true,target:node,key:'a'}),false);
});

test('mobile suppression restores absent, empty, explicit and original none input modes exactly', t => {
  const env = keyboardEnvironment(t, 'hardware');
  const originals = [null, '', 'decimal', 'TEXT', 'none'];
  const nodes = originals.map(value => env.add(new KeyboardNode('input', value == null ? {} : {inputmode:value})));
  const changes = [], c = env.install((next, prev) => { changes.push([next,prev]); if (next.suppressed) for (const node of nodes) assert.equal(node.getAttribute('inputmode'),'none'); });
  assert.equal(c.preference,'hardware'); assert.equal(c.suppressed,true);
  c.setPreference('screen');
  assert.deepEqual(nodes.map(n=>n.getAttribute('inputmode')), originals);
  assert.equal(nodes[0].hasAttribute('inputmode'),false); assert.equal(nodes[1].hasAttribute('inputmode'),true);
  assert.equal(env.values.get(MOBILE_KEYBOARD_KEY),'screen'); assert.equal(changes.length,2);
  for (const node of nodes) { assert.equal(node.value,'한글 값'); assert.equal(node.readOnly,false); assert.equal(node.disabled,false); assert.equal(node.focusCalls,0); assert.equal(node.inputEvents,0); }
});

test('hardware preference applies only on mobile and preserves text, form behavior and ignored input types', t => {
  const env=keyboardEnvironment(t,'hardware'); env.mobile=false;
  const nodes = ['text','search','url','tel','email','password','number'].map(type=>env.add(new KeyboardNode('input',{type,inputmode:'decimal'})));
  nodes.push(env.add(new KeyboardNode('textarea')),env.add(new KeyboardNode('div',{contenteditable:'plaintext-only'})));
  const others=['checkbox','radio','range','button','file','hidden','date','color'].map(type=>env.add(new KeyboardNode('input',{type,inputmode:'numeric'})));
  const noEdit=env.add(new KeyboardNode('div',{contenteditable:'false'})), c=env.install();
  assert.equal(c.suppressed,false); assert.equal(nodes[0].getAttribute('inputmode'),'decimal');
  env.mobile=true; c.refresh(); assert.equal(c.suppressed,true);
  for(const node of nodes) assert.equal(node.getAttribute('inputmode'),'none');
  for(const node of others) assert.equal(node.getAttribute('inputmode'),'numeric');
  assert.equal(noEdit.getAttribute('inputmode'),null);
  env.mobile=false; c.refresh(); assert.equal(c.suppressed,false); assert.equal(nodes[0].getAttribute('inputmode'),'decimal');
});

test('automatic detection is session-only, never inferred from a mouse/touch, and explicitly resets', t=>{
  const env=keyboardEnvironment(t), input=env.add(new KeyboardNode('textarea')), changes=[], c=env.install((next)=>changes.push(next));
  for(const pointerType of ['mouse','touch','pen']) env.emit('pointerdown',{pointerType,target:input});
  env.emit('keydown',{key:'a',target:input}); env.emit('keydown',{key:'a',isTrusted:false}); env.emit('keydown',{key:'Process',keyCode:229});
  assert.equal(c.suppressed,false);
  env.emit('keydown',{key:'a'}); assert.equal(c.suppressed,true); assert.equal(input.getAttribute('inputmode'),'none');
  assert.equal(c.preference,'auto'); assert.equal(env.values.has(MOBILE_KEYBOARD_KEY),false);
  env.mobile=false;c.refresh();assert.equal(c.suppressed,false);env.mobile=true;c.refresh();assert.equal(c.suppressed,true);
  c.setPreference('auto');assert.equal(c.suppressed,false);
  c.setPreference('screen'); env.emit('keydown',{key:'a'});assert.equal(c.suppressed,false);
  assert.equal(changes.length,6);
});

test('dynamic inputs are prepared before focus and original modes survive subtree insertion/removal', t=>{
  const env=keyboardEnvironment(t,'hardware'),c=env.install(), detached=new KeyboardNode('input',{inputmode:'email'});
  detached.isConnected=false; prepareKeyboardInput(detached); assert.equal(detached.getAttribute('inputmode'),'none'); assert.equal(detached.focusCalls,0);
  detached.isConnected=true;env.add(detached);env.mutate([{type:'childList',addedNodes:[detached]}]);
  const panel=new KeyboardNode('section'), nested=new KeyboardNode('textarea',{inputmode:''});panel.append(nested);env.add(panel);c.prepare(panel);assert.equal(nested.getAttribute('inputmode'),'none');
  const raw=env.add(new KeyboardNode('input',{inputmode:'tel'}));env.emit('pointerdown',{target:raw,pointerType:'touch'});assert.equal(raw.getAttribute('inputmode'),'none');
  detached.isConnected=false;env.mutate([{type:'childList',addedNodes:[],removedNodes:[detached]}]);assert.equal(detached.getAttribute('inputmode'),'email');
  c.setPreference('screen');assert.equal(nested.getAttribute('inputmode'),'');assert.equal(raw.getAttribute('inputmode'),'tel');
});

test('app inputmode/type changes restore the latest authored input mode without observer feedback', t=>{
  const env=keyboardEnvironment(t,'hardware'), input=env.add(new KeyboardNode()),c=env.install();
  input.setAttribute('inputmode','decimal');env.mutate([{type:'attributes',target:input,attributeName:'inputmode'}]);assert.equal(input.getAttribute('inputmode'),'none');
  env.mutate([{type:'attributes',target:input,attributeName:'inputmode'}]);
  c.setPreference('screen');assert.equal(input.getAttribute('inputmode'),'decimal');
  c.setPreference('hardware');input.setAttribute('type','checkbox');env.mutate([{type:'attributes',target:input,attributeName:'type'}]);assert.equal(input.getAttribute('inputmode'),'decimal');
});

test('policy restoration and callback wait until composition completes without changing IME text', async t=>{
  const env=keyboardEnvironment(t,'hardware'), input=env.add(new KeyboardNode('textarea',{inputmode:'text'})),changes=[],c=env.install((next,prev)=>changes.push([next,prev]));
  env.emit('compositionstart',{target:input}); input.value='한'; c.setPreference('screen');
  assert.equal(c.suppressed,true);assert.equal(input.getAttribute('inputmode'),'none');assert.equal(changes.length,1);
  c.refresh();env.emit('keydown',{key:'Process',keyCode:229,isComposing:true,target:input});
  input.value='한국';env.emit('compositionend',{target:input});
  c.refresh(); // Reentrant compositionend listener must also wait for final input.
  assert.equal(input.getAttribute('inputmode'),'none');assert.equal(changes.length,1);
  await Promise.resolve();assert.equal(c.suppressed,false);assert.equal(input.getAttribute('inputmode'),'text');assert.equal(changes.length,2);assert.equal(input.value,'한국');
  assert.equal(input.focusCalls,0);assert.equal(input.inputEvents,0);
});

test('composition defers enabling suppression and mobile mode changes, keeping last requested preference', async t=>{
  const env=keyboardEnvironment(t,'screen'),input=env.add(new KeyboardNode('textarea')), c=env.install();
  env.emit('compositionstart',{target:input});c.setPreference('hardware');assert.equal(input.getAttribute('inputmode'),null);assert.equal(c.suppressed,false);
  c.setPreference('auto');c.setPreference('hardware');env.emit('compositionend',{target:input});await Promise.resolve();assert.equal(c.suppressed,true);
  env.emit('compositionstart',{target:input});env.mobile=false;c.refresh();assert.equal(c.suppressed,true);
  env.emit('compositionend',{target:input});await Promise.resolve();assert.equal(c.suppressed,false);assert.equal(input.hasAttribute('inputmode'),false);
});

test('disposal restores attributes, releases observers/listeners and a new controller replaces the old policy', t=>{
  const env=keyboardEnvironment(t,'hardware'), input=env.add(new KeyboardNode('textarea')), c1=env.install(),obs=env.observer;
  c1.dispose();assert.equal(input.hasAttribute('inputmode'),false);assert.equal(obs.disconnected,true);for(const list of env.listeners.values())assert.equal(list.size,0);
  prepareKeyboardInput(input);assert.equal(input.hasAttribute('inputmode'),false);
  const c2=env.install();assert.equal(c2.suppressed,true);const c3=env.install();c2.setPreference('screen');assert.equal(input.getAttribute('inputmode'),'none');c3.dispose();assert.equal(input.hasAttribute('inputmode'),false);
});

test('blocked storage does not prevent changing policy or restoring inputs', t=>{
  const env=keyboardEnvironment(t);Object.defineProperty(globalThis,'localStorage',{configurable:true,get(){throw Error('denied');}});
  const input=env.add(new KeyboardNode('textarea')),c=env.install();c.setPreference('hardware');assert.equal(input.getAttribute('inputmode'),'none');c.setPreference('screen');assert.equal(input.hasAttribute('inputmode'),false);
});


test('removing a composing dialog releases pending policy without touching a connected composition', async t=>{
  const env=keyboardEnvironment(t,'hardware'),removed=env.add(new KeyboardNode('textarea',{inputmode:'text'})),live=env.add(new KeyboardNode('input',{inputmode:'decimal'})),c=env.install();
  env.emit('compositionstart',{target:removed});c.setPreference('screen');removed.isConnected=false;
  env.mutate([{type:'childList',addedNodes:[],removedNodes:[removed]}]);
  assert.equal(c.suppressed,false);assert.equal(removed.getAttribute('inputmode'),'text');assert.equal(live.getAttribute('inputmode'),'decimal');
  c.setPreference('hardware');env.emit('compositionstart',{target:live});env.mobile=false;c.refresh();
  env.mutate([{type:'childList',addedNodes:[],removedNodes:[]}]);assert.equal(c.suppressed,true);assert.equal(live.getAttribute('inputmode'),'none');
  env.emit('compositionend',{target:live});await Promise.resolve();assert.equal(c.suppressed,false);assert.equal(live.getAttribute('inputmode'),'decimal');
});


test('ordinary grid mutations do not scan input subtrees when suppression is inactive', t=>{
  const env=keyboardEnvironment(t),c=env.install(), grid=new KeyboardNode('div');
  grid.querySelectorAll=()=>{throw Error('화면 키보드 모드에서 격자 전체를 검색했습니다');};
  c.prepare(grid);env.mutate([{type:'childList',addedNodes:[grid]}]);
  assert.equal(c.suppressed,false);
});
