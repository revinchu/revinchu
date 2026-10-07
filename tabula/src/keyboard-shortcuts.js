// Physical codes take precedence so switching the input language does not
// change spreadsheet shortcuts. Legacy codes are used only when code is absent.
export function shortcutCode(event) {
  if (!event || event.isComposing || event.keyCode === 229 || event.key === 'Process' || event.key === 'Dead') return '';
  if (event.code && event.code !== 'Unidentified') return event.code;
  const key = event.key || '';
  if (/^[a-z]$/i.test(key)) return `Key${key.toUpperCase()}`;
  if (/^[0-9]$/.test(key)) return `Digit${key}`;
  const legacy = event.keyCode || event.which;
  if (legacy >= 65 && legacy <= 90) return `Key${String.fromCharCode(legacy)}`;
  if (legacy >= 48 && legacy <= 57) return `Digit${legacy - 48}`;
  return '';
}

// A physical function key can be reported as Process/229 by an idle IME.
// Active text composition remains owned by the editor, never by repeat.
export function repeatFunctionKey(event) {
  if (!event || event.isComposing || event.ctrlKey || event.metaKey || event.altKey || event.shiftKey || event.getModifierState?.('AltGraph')) return false;
  return event.code === 'F4' || event.key === 'F4' && event.keyCode !== 229;
}

// iPadOS can identify itself as a Mac, and a wide iPad or attached pointer can
// select desktop layout. Clipboard behavior must not depend on layout density.
export function appleTouchDevice(platform = globalThis.navigator) {
  return !!platform && (/iPad|iPhone|iPod/.test(platform.userAgent || '')
    || (/^Mac/.test(platform.platform || '') && platform.maxTouchPoints > 1));
}

// Exact modifier combinations run before broader grid fallbacks. No aliases
// for unsupported desktop features (for example threaded comments or VBA).
export function excelDirectCommand(event) {
  if (!event || event.isComposing || event.keyCode === 229 || event.getModifierState?.('AltGraph')) return '';
  const code = shortcutCode(event), ctrl = !!(event.ctrlKey || event.metaKey), alt = !!event.altKey, shift = !!event.shiftKey;
  if (!ctrl && alt && shift && event.key === 'F1') return 'addSheet';
  if (ctrl && !alt && shift) {
    if (event.key === 'F1') return 'fullScreen';
    if (event.key === 'F12') return 'print';
    return { KeyG: 'workbookStats', KeyS: 'saveAs', KeyZ: 'redo' }[code] ?? '';
  }
  if (!alt && event.key === 'F12') {
    if (ctrl && !shift) return 'open';
    if (!ctrl && shift) return 'save';
  }
  if (ctrl && alt && !shift) {
    if (code === 'Equal' || code === 'NumpadAdd') return 'zoomIn';
    if (code === 'Minus' || code === 'NumpadSubtract') return 'zoomOut';
  }
  return '';
}
