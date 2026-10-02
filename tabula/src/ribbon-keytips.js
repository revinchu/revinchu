// DOM-free inventory and KeyTip registry. Generated Zxx paths are WIXEL keys,
// not a claim that every Microsoft 365 version uses the same ribbon layout.
export const RIBBON_TAB_KEYS = {
  h: 'home', n: 'insert', p: 'layout', m: 'formulas', a: 'data', r: 'review', w: 'view', y: 'help',
  jt: 'tableDesign', jp: 'pivotAnalyze', jd: 'pivotDesign', jk: 'sparkTab', js: 'slicerTab', jo: 'objFormat', jq: 'pictureFormat', jc: 'chartDesign',
};

// Excel's documented HH opens the color picker. Requested compatibility keys
// have source=user; unverified size shortcuts remain explicitly WIXEL mappings.
export const PREFERRED_RIBBON_KEYTIPS = [
  { path: 'hj', tabId: 'home', kind: 'menu', target: 'cellStyles', source: 'user' },
  { path: 'hh', tabId: 'home', kind: 'menu', target: 'fillColor', source: 'excel' },
  { path: 'hfc', tabId: 'home', kind: 'menu', target: 'fontColor', source: 'user' },
  { path: 'hff', tabId: 'home', kind: 'input', target: 'fontFamily', source: 'user' },
  { path: 'hfs', tabId: 'home', kind: 'input', target: 'fontSize', source: 'user' },
  { path: 'hfp', tabId: 'home', kind: 'command', target: 'painter', source: 'user' },
  { path: 'hfg', tabId: 'home', kind: 'command', target: 'growFont', source: 'wixel' },
  { path: 'hfk', tabId: 'home', kind: 'command', target: 'shrinkFont', source: 'wixel' },
  { path: 'wvg', tabId: 'view', kind: 'command', target: 'toggleGrid', source: 'user' },
];

const INPUT_TYPES = new Set(['font', 'select', 'text', 'spin']);
const BUTTON_TYPES = new Set(['large', 'medium', 'btn', 'check', 'color']);
const SPLIT_COMMANDS = new Set(['borderLast', 'mergeCenter', 'paste']);
const LEGACY_MENUS = { tableStyleKey: 'tableStyles', condMenuKey: 'condFormat', shapesMenu: 'shapes' };
const labelOf = (item) => String(item.title ?? item.label ?? item.cmd ?? item.menu ?? '').replace(/\n/g, ' ').trim();

/** One record per real action: a split button's command and menu are separate.
 * item is the original TABS leaf reference, for renderer metadata matching.
 */
export function collectRibbonControls(tabs) {
  const controls = [], ids = new Map();
  for (const tab of tabs) {
    for (const [groupIndex, group] of (tab.groups ?? []).entries()) {
      const add = (item, kind, target, launcher = false) => {
        if (!target) throw new Error(`리본 조작 대상이 없습니다: ${tab.id}/${group.label}/${kind}`);
        const base = `${tab.id}:${kind}:${target}`, n = (ids.get(base) ?? 0) + 1;
        ids.set(base, n);
        const label = launcher ? `${group.label} 대화상자` : labelOf(item);
        controls.push({ id: base + (n > 1 ? `:${n}` : ''), tabId: tab.id, tabLabel: tab.label, context: tab.context ?? null,
          group: group.label, groupIndex, itemType: launcher ? 'launcher' : item.type, kind, target,
          label: kind === 'menu' ? `${label} 메뉴` : label, command: launcher ? target : item.cmd, menu: item?.menu,
          item: launcher ? null : item, launcher });
      };
      const walk = (item) => {
        if (item.type === 'row' || item.type === 'col') { for (const child of item.items ?? []) walk(child); return; }
        if (item.type === 'sep') return;
        if (INPUT_TYPES.has(item.type)) {
          add(item, 'input', item.cmd);
          if (item.menu) add(item, 'menu', item.menu);
          return;
        }
        if (item.type === 'combo' || item.type === 'gallery') { add(item, 'menu', item.menu); return; }
        if (!BUTTON_TYPES.has(item.type)) throw new Error(`지원하지 않는 리본 조작 종류: ${item.type}`);
        const menuOnly = item.menu && !item.split && !SPLIT_COMMANDS.has(item.cmd) && item.type !== 'color';
        if (!menuOnly) add(item, 'command', item.cmd);
        if (item.menu) add(item, 'menu', item.menu);
      };
      for (const item of group.items ?? []) walk(item);
      if (group.launcher) add(null, 'command', group.launcher, true);
    }
    if (!tab.file) controls.push({ id: `${tab.id}:command:toggleRibbon`, tabId: tab.id, tabLabel: tab.label, context: tab.context ?? null,
      group: '기타', groupIndex: -1, itemType: 'chrome', kind: 'command', target: 'toggleRibbon', label: '리본 축소', command: 'toggleRibbon',
      item: null, launcher: false, chrome: true });
  }
  return controls;
}

/** Checks registration independently of the fallback allocator. */
export function auditRibbonKeytips(controls, entries, tabs = {}) {
  const seen = new Set(), duplicatePaths = [], invalidEntries = [], prefixConflicts = [], covered = new Set();
  const controlMap = new Map(controls.map(c => [c.id, c]));
  for (const entry of entries) {
    if (seen.has(entry.path)) duplicatePaths.push(entry.path);
    seen.add(entry.path);
    if (!/^[a-z0-9]+$/.test(entry.path) || !['command', 'menu', 'input'].includes(entry.kind) || !entry.target) invalidEntries.push(entry.path);
    if (entry.controlId) {
      const control = controlMap.get(entry.controlId);
      if (!control || control.kind !== entry.kind || control.target !== entry.target || control.tabId !== entry.tabId) invalidEntries.push(entry.path);
      else covered.add(entry.controlId);
    }
  }
  for (const path of seen) for (const other of seen) if (path !== other && other.startsWith(path)) prefixConflicts.push([path, other]);
  // A terminal action cannot consume a prefix needed to enter a ribbon tab.
  for (const path of seen) for (const tab of Object.keys(tabs)) if (tab === path || tab.startsWith(path)) prefixConflicts.push([path, tab]);
  return { total: controls.length, covered: covered.size, missingControlIds: controls.filter(c => !covered.has(c.id)).map(c => c.id), duplicatePaths, prefixConflicts, invalidEntries };
}

export function createRibbonKeytipRegistry(tabs, legacy = {}) {
  const controls = collectRibbonControls(tabs);
  const knownTabs = new Set(tabs.map(t => t.id));
  const tabKeys = Object.fromEntries(Object.entries(RIBBON_TAB_KEYS).filter(([, id]) => knownTabs.has(id)));
  for (const tab of tabs) if (!tab.file && !Object.values(tabKeys).includes(tab.id)) throw new Error(`리본 탭 키가 없습니다: ${tab.id}`);
  const entries = [], byPath = new Map();
  const actionMatches = (a, b) => a.tabId === b.tabId && a.kind === b.kind && a.target === b.target;
  const tabFor = (path) => Object.entries(tabKeys).sort((a, b) => b[0].length - a[0].length).find(([prefix]) => path.startsWith(prefix))?.[1] ?? (path === 'f' ? 'file' : null);
  const add = (entry) => {
    const old = byPath.get(entry.path);
    if (old) {
      if (!actionMatches(old, entry)) throw new Error(`리본 키 충돌: ${entry.path}`);
      Object.assign(old, entry); return old;
    }
    entries.push(entry); byPath.set(entry.path, entry); return entry;
  };
  for (const [path, value] of Object.entries(legacy)) {
    const [command, label] = value;
    const menu = LEGACY_MENUS[command], tabId = tabFor(path), kind = menu ? 'menu' : 'command', target = menu ?? command;
    const control = controls.find(c => actionMatches(c, { tabId, kind, target }));
    add({ path, tabId, kind, target, label, controlId: control?.id ?? null, source: 'legacy', primary: false });
  }
  for (const preferred of PREFERRED_RIBBON_KEYTIPS) {
    const control = controls.find(c => actionMatches(c, preferred));
    if (control) add({ ...preferred, label: control.label, controlId: control.id, primary: false });
  }
  // Fixed-width codes avoid Z1 consuming Z10. Never silently drop a control.
  for (const control of controls) {
    let entry = entries.find(e => e.controlId === control.id && e.source !== 'legacy') ?? entries.find(e => e.controlId === control.id);
    if (!entry) {
      const prefix = Object.entries(tabKeys).find(([, id]) => id === control.tabId)?.[0];
      let path = null;
      for (const branch of ['z', 'x', 'q', 'u']) {
        for (let i = 1; i <= 99; i++) {
          const candidate = `${prefix}${branch}${String(i).padStart(2, '0')}`;
          if (entries.some(e => candidate.startsWith(e.path) || e.path.startsWith(candidate))) continue;
          if (Object.keys(tabKeys).some(p => p === candidate || p.startsWith(candidate))) continue;
          path = candidate; break;
        }
        if (path) break;
      }
      if (!path) throw new Error(`리본 키를 배정하지 못했습니다: ${control.id}`);
      entry = add({ path, tabId: control.tabId, kind: control.kind, target: control.target, label: control.label, controlId: control.id, source: 'wixel', primary: false });
    }
    entry.primary = true;
    control.path = entry.path;
  }
  const audit = auditRibbonKeytips(controls, entries, tabKeys);
  if (audit.missingControlIds.length || audit.duplicatePaths.length || audit.prefixConflicts.length || audit.invalidEntries.length) throw new Error('리본 키 등록표에 누락 또는 충돌이 있습니다.');
  return { controls, entries, tabs: tabKeys, audit };
}
