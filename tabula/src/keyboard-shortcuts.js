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

// iPadOS can identify itself as a Mac, and a wide iPad or attached pointer can
// select desktop layout. Clipboard behavior must not depend on layout density.
export function appleTouchDevice(platform = globalThis.navigator) {
  return !!platform && (/iPad|iPhone|iPod/.test(platform.userAgent || '')
    || (/^Mac/.test(platform.platform || '') && platform.maxTouchPoints > 1));
}
