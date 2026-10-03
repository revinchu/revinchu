import { el, openDialog } from './ui.js';
import { shortcutCode } from './keyboard-shortcuts.js';

const CHECK_ACTIONS = [['KeyC', '복사'], ['KeyX', '잘라내기'], ['KeyV', '붙여넣기'], ['KeyZ', '실행 취소']];
const CHECK_MODIFIERS = [['ctrlKey', 'Ctrl'], ['altKey', 'Alt/Option'], ['metaKey', 'Command/Meta'], ['shiftKey', 'Shift']];
const CHECK_CODES = /^(?:Key[A-Z]|Digit[0-9]|(?:Control|Alt|Meta|Shift)(?:Left|Right)|Arrow(?:Up|Down|Left|Right)|F(?:[1-9]|1[0-9]|2[0-4])|Numpad(?:[0-9]|Add|Subtract|Multiply|Divide|Decimal|Enter)|Space|Enter|Backspace|Delete|Home|End|PageUp|PageDown|CapsLock|NumLock|ContextMenu|Insert)$/;

/** 키 내용·클립보드를 읽거나 저장하지 않는, 현재 창 안의 키 수신 확인. */
export function openKeyboardCheck() {
  const modifierNodes = CHECK_MODIFIERS.map(([, label]) => el('span', {
    style: { display: 'inline-block', padding: '3px 6px', border: '1px solid #cbd5d1', borderRadius: '4px' },
  }, `${label}: 꺼짐`));
  const physical = el('div', { 'data-keyboard-check-physical': '', style: { overflowWrap: 'anywhere' } }, '아직 키를 받지 않았습니다.');
  const statuses = new Map(CHECK_ACTIONS.map(([code]) => [code, el('span', { 'data-keyboard-check-action': code, style: { overflowWrap: 'anywhere' } }, '수신 대기')]));
  const pad = el('div', {
    tabindex: '0', role: 'group', 'aria-label': '키보드 시험 영역', 'data-keyboard-check-pad': '', 'data-access-key': 'none',
    style: { padding: '12px', border: '2px solid #217346', borderRadius: '6px', background: '#f2faf5', color: '#173b28', cursor: 'pointer', userSelect: 'none', touchAction: 'manipulation' },
    onclick: () => pad.focus({ preventScroll: true }),
  }, '이 영역에서 Ctrl·Alt·Command를 각각 누른 뒤 C·X·V·Z와 함께 눌러 보세요.');
  const modifiers = (event) => {
    CHECK_MODIFIERS.forEach(([key, label], index) => {
      const down = !!event?.[key], node = modifierNodes[index];
      node.textContent = `${label}: ${down ? '켜짐' : '꺼짐'}`;
      node.style.background = down ? '#d7f0df' : 'transparent';
      node.style.fontWeight = down ? '700' : '400';
    });
  };
  const onKey = (event) => {
    // IME 조합·죽은 키는 검사도, 기본 동작 차단도 하지 않는다.
    if (event.isComposing || event.keyCode === 229 || event.key === 'Process' || event.key === 'Dead') return;
    if (event.key === 'Escape' || event.key === 'Tab') return;
    modifiers(event);
    // event.key의 입력 글자는 출력하지 않는다. code도 허용된 물리 키 이름만 표시한다.
    const nativeCode = CHECK_CODES.test(event.code || '') ? event.code : '';
    const fallback = !event.code || event.code === 'Unidentified' ? shortcutCode(event) : '';
    const code = nativeCode || (CHECK_CODES.test(fallback) ? fallback : '');
    if (event.type === 'keydown') {
      const side = event.location === 1 ? ' · 왼쪽' : event.location === 2 ? ' · 오른쪽' : event.location === 3 ? ' · 숫자 키패드' : '';
      physical.textContent = code ? `${nativeCode ? '물리 키' : '대체 판정(물리 키 미제공)'}: ${code}${side}` : '키 수신됨 · 물리 키 코드 미제공/지원 밖';
      const legacy = Number(event.keyCode || event.which || 0);
      if (Number.isInteger(legacy) && legacy > 0 && legacy <= 255) physical.textContent += ' · 레거시 코드 ' + legacy;
      const status = statuses.get(code);
      if (status && (event.ctrlKey || event.altKey || event.metaKey)) {
        const combo = CHECK_MODIFIERS.filter(([key]) => event[key]).map(([, label]) => label).join(' + ');
        const supported = (event.ctrlKey || event.metaKey) && !event.altKey && !event.shiftKey;
        status.textContent = `${combo} + ${code.slice(3)} 수신 · ${supported ? '기본 조합' : '다른 조합'} · 실행 안 함`;
      }
    }
    // 진단 영역에서 브라우저 기본 단축키·격자 명령·클립보드 작업을 실행하지 않는다.
    event.preventDefault();
    event.stopPropagation();
  };
  pad.addEventListener('keydown', onKey);
  pad.addEventListener('keyup', onKey);
  pad.addEventListener('blur', () => modifiers(null));
  for (const type of ['copy', 'cut', 'paste']) pad.addEventListener(type, (event) => {
    event.preventDefault(); event.stopPropagation(); // clipboardData에는 접근하지 않는다.
  });
  const external = (label, href) => el('a', { href, target: '_blank', rel: 'noopener noreferrer', tabindex: '0', 'data-access-key': 'none' }, label);
  const body = el('div', { class: 'keyboard-check', style: { display: 'grid', gap: '9px', fontSize: '13px', lineHeight: '1.45', minWidth: '0' } },
    el('p', { style: { margin: '0' } }, 'iPad·Mac의 기본 복사/붙여넣기는 Command(⌘)+C·V입니다. 위셀 셀 선택에서는 Ctrl+C·V도 지원합니다. 키에 적힌 이름과 실제 전달되는 Ctrl·Alt·Command가 같은지 확인하세요.'),
    pad,
    el('div', { 'data-keyboard-check-modifiers': '', style: { display: 'flex', flexWrap: 'wrap', gap: '4px' } }, modifierNodes),
    el('div', { role: 'status', 'aria-live': 'polite', 'aria-atomic': 'true' }, physical),
    el('div', { style: { display: 'grid', gridTemplateColumns: '72px minmax(0,1fr)', gap: '5px 8px' } }, CHECK_ACTIONS.map(([code, name]) => [el('strong', {}, name), statuses.get(code)])),
    el('p', { class: 'muted', style: { margin: '0' } }, '시험 영역에서는 문서·클립보드를 바꾸지 않습니다. 입력 글자와 키 이력은 저장·전송하지 않습니다. OS가 먼저 처리하여 브라우저에 전달하지 않은 키는 확인할 수 없습니다.'),
    el('p', { class: 'muted', style: { margin: '0' } }, 'Tab으로 링크·닫기 버튼 이동 · Escape로 닫기. 키 조합을 다시 확인하려면 시험 영역을 누르세요.'),
    el('div', { style: { display: 'flex', flexWrap: 'wrap', gap: '6px 12px' } },
      external('Apple iPad 단축키 안내', 'https://support.apple.com/en-us/102393'),
      external('Logitech Keys-To-Go 2 안내', 'https://support.logi.com/hc/en-us/articles/22241277704215-Getting-Started-Keys-To-Go-2'),
      external('iPad용 모델 안내', 'https://support.logi.com/hc/en-us/articles/22241608375959-Getting-Started-Keys-To-Go-2-for-iPad')));
  return openDialog({ title: '키보드 단축키 확인', width: 500, body, initialFocus: pad,
    buttons: [{ label: '닫기', primary: true, accessKey: 'none' }],
    onOpen: (dialog) => {
      // 진단 중 Alt+D 등도 관찰할 수 있도록 이 창의 자동 접근키만 비활성화한다.
      const head = dialog.querySelector('[data-dialog-close-head]');
      head?.removeAttribute('data-dialog-close-head');
      head?.setAttribute('data-access-key', 'none');
    },
  });
}
