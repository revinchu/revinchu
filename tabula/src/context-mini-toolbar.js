// 우클릭 미니 서식 도구 모음. 명령/선택/보호 판단은 app.js의 기존 경로를 사용한다.
import { el, icon, toast } from './ui.js';
import { FONTS, FONT_SIZES } from './ribbon.js';
import { createFormatValuePicker } from './format-value-picker.js';

/** onCommand(cmd,value), openNamedMenu(name,anchor), getStyle()->현재 서식, readonly:boolean|()=>boolean */
export function createContextMiniToolbar({ onCommand, openNamedMenu, getStyle = () => ({}), readonly = false } = {}) {
  const root = el('div', { class: 'context-mini-toolbar cell-mini-toolbar', role: 'toolbar', 'aria-label': '미니 서식 도구 모음' });
  const controls = [], toggles = [];
  const locked = () => typeof readonly === 'function' ? !!readonly() : !!readonly;
  const invoke = async (cmd, value) => {
    if (locked()) return;
    try { await onCommand?.(cmd, value); }
    catch (error) { toast(`서식을 적용하지 못했습니다: ${error?.message ?? String(error)}`); }
    finally { root.refresh(); }
  };
  const button = (label, content, cmd, { menu, pressed } = {}) => {
    const node = el('button', { type: 'button', class: 'mini-format-button', title: label, 'aria-label': label, 'data-mini-command': cmd ?? menu, 'aria-haspopup': menu ? 'menu' : null,
      onclick: (event) => {
        if (locked()) return;
        if (menu) {
          try { openNamedMenu?.(menu, event.currentTarget); } catch (error) { toast(`서식 메뉴를 열지 못했습니다: ${error?.message ?? String(error)}`); }
        } else invoke(cmd);
      },
    }, typeof content === 'string' ? icon(content) : content);
    if (pressed) toggles.push([node, pressed]);
    controls.push(node); return node;
  };
  const fontPicker = createFormatValuePicker({ label: '글꼴', className: 'mini-font-picker', inputClass: 'mini-font', values: FONTS, disabled: locked, onChange: (v) => invoke('fontFamily', v), onOpen: openNamedMenu ? (anchor) => openNamedMenu('fontList', anchor) : undefined });
  const sizePicker = createFormatValuePicker({ label: '글꼴 크기', className: 'mini-size-picker', inputClass: 'mini-size', values: FONT_SIZES, number: true, disabled: locked, onChange: (v) => invoke('fontSize', v) });
  controls.push(fontPicker.input, fontPicker.button, sizePicker.input, sizePicker.button);
  root.append(el('div', { class: 'mini-format-row' }, fontPicker.root, sizePicker.root,
    button('글꼴 크기 크게', el('span', { class: 'mini-grow' }, '가', el('sup', {}, '▲')), 'growFont'),
    button('글꼴 크기 작게', el('span', { class: 'mini-grow' }, '가', el('sup', {}, '▼')), 'shrinkFont'),
    button('서식 복사', 'painter', 'painter', { pressed: (style) => !!style.painter })));
  const fill = button('채우기 색', 'fill', null, { menu: 'fillColor' }); fill.classList.add('mini-fill');
  const color = button('글꼴 색', 'fontColor', null, { menu: 'fontColor' }); color.classList.add('mini-color');
  root.append(el('div', { class: 'mini-format-row' },
    button('굵게', el('b', {}, '가'), 'bold', { pressed: (style) => !!style.bold }),
    button('기울임꼴', el('i', {}, '가'), 'italic', { pressed: (style) => !!style.italic }),
    button('밑줄', el('u', {}, '가'), 'underline', { pressed: (style) => !!style.underline }),
    button('가운데 맞춤', 'alignCenter', 'alignCenter', { pressed: (style) => style.align === 'center' }), fill, color,
    button('테두리', 'border', null, { menu: 'borders' }),
    button('표시 형식', 'currency', null, { menu: 'numFormats' }),
    button('백분율 스타일', 'percent', 'fmtPercent'), button('쉼표 스타일', 'comma', 'fmtComma'),
    button('자릿수 늘림', 'incDecimal', 'incDecimal'), button('자릿수 줄임', 'decDecimal', 'decDecimal'),
    button('병합하고 가운데 맞춤', 'merge', 'mergeCenter', { pressed: (style) => !!style.merged })));
  root.refresh = () => {
    const style = getStyle() ?? {}, disabled = locked();
    fontPicker.update(style.font || '맑은 고딕'); sizePicker.update(style.size || 11);
    for (const control of controls) control.disabled = disabled;
    for (const [node, read] of toggles) { const pressed = read(style); node.classList.toggle('on', pressed); node.setAttribute('aria-pressed', String(pressed)); }
    root.style.setProperty('--mini-fill', style.lastFill || style.fill || '#ffff00');
    root.style.setProperty('--mini-color', style.lastFont || style.color || '#ff0000');
    root.setAttribute('aria-disabled', String(disabled));
  };
  root.refresh(); return root;
}
