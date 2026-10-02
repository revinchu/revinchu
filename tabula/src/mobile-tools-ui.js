import { el, openDialog } from './ui.js';

/** A touch-size route to actions that compact chrome cannot show at once. */
export function openMobileTools(host) {
  let dialog;
  const invoke = action => () => { dialog.close(); action(); };
  const button = (label, action, attrs = {}) => el('button', { type: 'button', class: 'btn', ...attrs, onclick: invoke(action) }, label);
  const section = (label, children) => el('section', { class: 'mobile-tool-section' }, el('h3', {}, label), el('div', { class: 'mobile-tool-grid' }, children));
  const search = el('input', { type: 'search', disabled: true, placeholder: '명령 이름 검색', 'aria-label': '모바일 도구 명령 검색', autocomplete: 'off' });
  const results = el('div', { class: 'mobile-tool-grid mobile-tool-results', hidden: true });
  const commandButtons = host.commands.map(c => button(c.label, () => host.run(c.cmd), { 'data-mobile-command': c.cmd, disabled: !!c.disabled }));
  search.addEventListener('input', () => {
    const query = search.value.trim().toLocaleLowerCase();
    results.hidden = !query; results.replaceChildren();
    if (query) { const found = commandButtons.filter((b, i) => `${host.commands[i].label} ${host.commands[i].tab ?? ''}`.toLocaleLowerCase().includes(query)); results.append(...found); if (!found.length) results.append(el('p', {}, '일치하는 명령이 없습니다. 아래 메뉴 탭에서 찾아보세요.')); }
  });
  const body = el('div', { class: 'mobile-tools-body' },
    el('p', { class: 'mobile-tools-tip' }, '메뉴와 시트 탭을 좌우로 밀어보세요. 셀 한 번 누르기: 선택 · 두 번: 편집 · 길게 누른 뒤 끌기: 범위 · 두 손가락: 확대/축소'),
    search, results,
    section('메뉴 크기', [
      button('촘촘하게', () => host.setDensity('compact'), { 'aria-pressed': String(host.density === 'compact') }),
      button('여유롭게', () => host.setDensity('comfortable'), { 'aria-pressed': String(host.density === 'comfortable') }),
    ]),
    section('화면 맞춤', [
      button('화면에 맞추기', host.fit), button(`배율 ${host.zoom}% · 변경`, host.zoomDialog),
      button('선택 영역 맞춤', () => host.run('zoomSel')), button('리본 접기 / 펼치기', () => host.run('toggleRibbon')),
      button('자동 감지', () => host.preference('auto')), button('데스크톱 화면', () => host.preference('off')),
    ]),
    section('셀 작업', ['edit', 'goto', 'copy', 'cut', 'paste', 'pasteSpecial', 'undo', 'redo', 'find', 'replace', 'formatCells'].map(cmd => {
      const labels = { edit: '선택 셀 편집', goto: '셀 / 범위 선택', copy: '복사', cut: '잘라내기', paste: '붙여넣기', pasteSpecial: '선택하여 붙여넣기', undo: '실행 취소', redo: '다시 실행', find: '찾기', replace: '바꾸기', formatCells: '셀 서식' };
      return button(labels[cmd], () => cmd === 'edit' ? host.edit() : host.run(cmd), { disabled: !!host.disabled(cmd) });
    })),
    section('전체 메뉴 탭', host.tabs.map(t => button(t.label, () => host.tab(t.id), { 'data-mobile-tab': t.id }))),
    section('문서와 저장', [
      ...[['open', '파일 열기'], ['save', '파일 저장'], ['saveAs', '다른 이름으로 저장'], ['saveLocations', '저장 위치'], ['versionHistory', '버전 기록'], ['publish', '공유'], ['print', '인쇄 / PDF'], ['editComment', '메모'], ['calculationStatus', '계산 상태'], ['options', '설정']].map(([cmd, label]) => button(label, () => host.run(cmd), { disabled: !!host.disabled(cmd) })),
      button(`자동 저장 ${host.autosave ? '켜짐' : '꺼짐'} · 전환`, host.toggleAutosave),
    ]),
    section('빠른 실행 도구', host.quick.map(c => button(c.label, () => host.run(c.cmd), { disabled: !!c.disabled }))),
    el('p', { class: 'mobile-tools-state' }, host.status),
  );
  dialog = openDialog({ title: '모바일 작업 도구', width: 560, body, buttons: [{ label: '닫기' }] });
  // Opening the drawer must not summon the software keyboard before a user asks to type.
  search.disabled = false;
  dialog.root.focus({ preventScroll: true });
  return dialog;
}
