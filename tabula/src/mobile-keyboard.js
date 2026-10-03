// Device preference only. Keep native editing/composition; inputmode is a hint,
// not a hardware-keyboard detector or an editable/readonly switch.
export const MOBILE_KEYBOARD_KEY = 'wixel.mobile-keyboard.v1';
const KEYBOARD_INPUTS = 'input,textarea,[contenteditable]';
const KEYBOARD_TEXT_TYPES = new Set(['text', 'search', 'url', 'tel', 'email', 'password', 'number']);
let activeKeyboardController = null;

export function mobileKeyboardPreference(value) {
  return value === 'hardware' || value === 'screen' ? value : 'auto';
}

function keyboardInput(node) {
  const tag = node?.tagName?.toLowerCase();
  if (tag === 'textarea') return true;
  if (tag === 'input') return KEYBOARD_TEXT_TYPES.has((node.type || 'text').toLowerCase());
  const editable = node?.getAttribute?.('contenteditable');
  return editable != null && editable.toLowerCase() !== 'false';
}

export function hardwareKeyboardEvent(event) {
  if (!event?.isTrusted || event.defaultPrevented || event.isComposing || event.keyCode === 229) return false;
  const target = event.target;
  if (target?.isContentEditable || target?.closest?.('input,textarea,select')) return false;
  const key = event.key || '';
  // Dead/Process/Unidentified and modifier-only events cannot safely identify
  // a physical keyboard. In particular, never intercept the first IME 229.
  return key.length === 1 || /^(?:Arrow(?:Up|Down|Left|Right)|Home|End|PageUp|PageDown|Enter|Tab|Escape|Backspace|Delete|Insert|F(?:[1-9]|1\d|2[0-4]))$/.test(key);
}

/** Called by ui.el after all original attributes have been assigned. */
export function prepareKeyboardInput(node) {
  activeKeyboardController?.prepareInput(node);
  return node;
}

export function installMobileKeyboard({ isMobile = () => false, onChange } = {}) {
  activeKeyboardController?.dispose();
  const doc = globalThis.document;
  let storage;
  try { storage = globalThis.localStorage; } catch {}
  let preference = 'auto';
  try { preference = mobileKeyboardPreference(storage?.getItem(MOBILE_KEYBOARD_KEY)); } catch {}
  let sessionHardware = false, state = null, disposed = false, endQueued = false;
  const originals = new Map(), composing = new Set();
  const remember = node => ({ present: node.hasAttribute('inputmode'), value: node.getAttribute('inputmode') });
  const restore = node => {
    const original = originals.get(node);
    if (!original) return;
    originals.delete(node);
    if (original.present) node.setAttribute('inputmode', original.value);
    else node.removeAttribute('inputmode');
  };
  const prepareInput = node => {
    if (disposed || composing.has(node)) return;
    if (!state?.suppressed || !keyboardInput(node)) { restore(node); return; }
    if (!originals.has(node)) originals.set(node, remember(node));
    if (node.getAttribute('inputmode') !== 'none') node.setAttribute('inputmode', 'none');
  };
  const prepare = root => {
    if (!root || disposed || !state?.suppressed) return root;
    prepareInput(root);
    for (const node of root.querySelectorAll?.(KEYBOARD_INPUTS) || []) prepareInput(node);
    return root;
  };
  const refresh = () => {
    if (disposed || composing.size || endQueued) return state;
    const next = { preference, suppressed: !!isMobile() && (preference === 'hardware' || (preference === 'auto' && sessionHardware)) };
    const previous = state;
    state = next;
    if (!previous || previous.suppressed !== next.suppressed) {
      if (next.suppressed) prepare(doc);
      else for (const node of originals.keys()) restore(node);
    }
    if (!previous || previous.preference !== next.preference || previous.suppressed !== next.suppressed) onChange?.(next, previous);
    return next;
  };
  const setPreference = value => {
    if (disposed) return state;
    preference = mobileKeyboardPreference(value);
    // Choosing Automatic explicitly starts a fresh detection session. Merely
    // changing viewport/mobile mode does not erase an already detected keyboard.
    if (preference === 'auto') sessionHardware = false;
    try { storage?.setItem(MOBILE_KEYBOARD_KEY, preference); } catch {}
    return refresh();
  };
  const pointerDown = event => { refresh(); prepareInput(event.target); };
  const keyDown = event => {
    if (!composing.size && preference === 'auto' && isMobile() && hardwareKeyboardEvent(event)) sessionHardware = true;
    refresh(); prepareInput(event.target);
  };
  const focusIn = event => prepareInput(event.target);
  const compositionStart = event => composing.add(event.target);
  const compositionEnd = event => {
    composing.delete(event.target);
    if (endQueued) return;
    endQueued = true;
    // Leave the compositionend/input dispatch intact; no blur/refocus or text
    // replay. The app receives the callback only after the new policy is applied.
    queueMicrotask(() => { endQueued = false; refresh(); });
  };
  const listeners = [['pointerdown', pointerDown], ['keydown', keyDown], ['focusin', focusIn],
    ['compositionstart', compositionStart], ['compositionend', compositionEnd]];
  for (const [name, fn] of listeners) doc?.addEventListener(name, fn, true);
  const observer = typeof MutationObserver === 'function' && doc ? new MutationObserver(records => {
    if (disposed) return;
    // A dialog can be removed without compositionend. Detached input must not
    // leave the policy permanently locked; connected composition stays native.
    let abandoned = false;
    for (const node of composing) if (node.isConnected === false) { composing.delete(node); abandoned = true; }
    if (abandoned) refresh();
    // Grid virtualization replaces many nodes; when no keyboard policy owns
    // attributes there is no reason to scan those non-input subtrees.
    if (!state?.suppressed && !originals.size) return;
    for (const record of records) {
      if (record.type === 'attributes') {
        const node = record.target;
        if (node.isConnected === false) { restore(node); continue; }
        // An application may intentionally update inputmode after insertion.
        // Retain that value for restoration, while preserving the active policy.
        if (record.attributeName === 'inputmode' && originals.has(node) && node.getAttribute('inputmode') !== 'none') originals.set(node, remember(node));
        prepareInput(node);
      } else for (const node of record.addedNodes || []) prepare(node);
    }
    // Do not retain closed dialogs/menus for the rest of a hardware session.
    for (const node of originals.keys()) if (!node.isConnected && !composing.has(node)) restore(node);
  }) : null;
  observer?.observe(doc.documentElement || doc, { subtree: true, childList: true, attributes: true, attributeFilter: ['inputmode', 'type', 'contenteditable'] });
  const dispose = () => {
    if (disposed) return;
    disposed = true; observer?.disconnect();
    for (const [name, fn] of listeners) doc?.removeEventListener(name, fn, true);
    for (const node of originals.keys()) restore(node);
    composing.clear();
    if (activeKeyboardController === controller) activeKeyboardController = null;
  };
  const controller = { get preference() { return preference; }, get suppressed() { return !!state?.suppressed; },
    setPreference, refresh, prepare, prepareInput, dispose };
  activeKeyboardController = controller;
  refresh();
  return controller;
}
