import { el } from './ui.js';
import { shapeSvg } from './shapes.js';
import { SHAPE_DASH_OPTIONS, SHAPE_ARROW_OPTIONS, SHAPE_PATTERN_OPTIONS, shapeSizePatch, shapeArrowEnd, shapeGradientStops, shapeGradientStopPatch, addShapeGradientStop } from './shape-format.js';

/** 모델리스 즉시 적용 패널. onChange는 한 편집을 한 트랜잭션으로 기록한다. */
export function createShapeFormatPanel({ getShape, onChange, isLine, fontName = '맑은 고딕', initialTab = '채우기 및 선' }) {
  const body = el('div', { class: 'cfp shape-format-pane' });
  const tabs = el('div', { class: 'format-pane-tabs shape-format-categories', role: 'tablist', 'aria-label': '도형 서식 범주' });
  const options = el('div', { class: 'format-pane-tabs shape-format-options', role: 'tablist', 'aria-label': '서식 옵션' });
  const navigation = el('div', { class: 'shape-format-navigation' }, options, tabs);
  const expanded = new Map();
  let selected = initialTab === '텍스트 옵션' ? '텍스트 및 글꼴' : initialTab, stopIndex = 0, textInput, linePreview;
  let lastShape = '채우기 및 선', lastText = '텍스트 및 글꼴';
  const preview = () => {
    const shape = getShape(); if (!linePreview || !shape) return;
    // 실제 도형 렌더러를 재사용하므로 대시·복합 선·화살표가 본문과 같은 방식으로 보인다.
    const sample = el('div', { html: shapeSvg({ ...shape, id: 'shape-format-line-preview', kind: 'line', w: 180, h: 0,
      path: undefined, customGeometry: undefined, fill: null, grad: undefined, pattern: undefined, shadow: undefined, glow: undefined, soft: undefined }) });
    linePreview.replaceChildren(sample);
  };
  const up = patch => { if (!getShape()) { draw(); return; } if (onChange(patch) === false) draw(); else preview(); };
  const textUp = patch => up({ ...patch, paras: undefined });
  const row = (name, input) => { input?.setAttribute('aria-label', name); return el('label', { class: 'cfp-row' }, el('span', {}, name), input); };
  const sec = (name, ...children) => {
    const details = el('details', { class: 'cfp-sec', open: expanded.get(name) !== false, 'data-shape-section': name }, el('summary', {}, name), ...children);
    details.addEventListener('toggle', () => { if (details.isConnected) expanded.set(name, details.open); }); return details;
  };
  const modes = (name, value, choices, change) => el('div', { class: 'shape-mode-options', role: 'radiogroup', 'aria-label': name },
    choices.map(([key, label]) => {
      const input = el('input', { type: 'radio', name: 'shape-format-' + name, value: key, checked: value === key, 'aria-label': label });
      input.addEventListener('change', () => { if (input.checked) change(key); });
      return el('label', {}, input, el('span', {}, label));
    }));
  const controlKey = node => node?.getAttribute('data-shape-focus') || node?.getAttribute('aria-label') ||
    (node?.tagName === 'SUMMARY' ? 'section:' + node.parentElement.dataset.shapeSection : node?.tagName === 'BUTTON' ? 'button:' + node.textContent.trim() : null);
  const num = (value, fn, min = 0, max = 100, step = 1, disabled = false) => {
    const input = el('input', { type: 'number', value: Number.isFinite(Number(value)) ? Math.round(Number(value) * 10000) / 10000 : min, min, max, step, disabled });
    input.addEventListener('change', () => { const n = Number(input.value); if (input.value === '' || !Number.isFinite(n)) return; const v = Math.max(min, Math.min(max, n)); input.value = v; fn(v); }); return input;
  };
  const check = (value, fn) => { const input = el('input', { type: 'checkbox', checked: !!value }); input.addEventListener('change', () => fn(input.checked)); return input; };
  const color = (value, fn) => { const input = el('input', { type: 'color', value: /^#[0-9a-f]{6}$/i.test(value ?? '') ? value : '#000000' }); input.addEventListener('change', () => fn(input.value)); return input; };
  const choose = (value, options, fn) => {
    const current = String(value ?? '');
    const list = options.some(([v]) => String(v) === current) ? options : [...options, [current, `현재 값 (${current})`]];
    const input = el('select', {}, list.map(([v, name]) => el('option', { value: v, selected: current === String(v) }, name)));
    input.addEventListener('change', () => fn(input.value)); return input;
  };
  const field = (value, fn) => { const input = el('input', { type: 'text', value: value ?? '' }); input.addEventListener('change', () => fn(input.value)); return input; };
  const draw = () => {
    // 전체 재구성 전 현재 펼침·초점·스크롤을 읽어 모델 새로 고침/Undo도 같은 편집 위치를 유지한다.
    for (const section of body.querySelectorAll('details[data-shape-section]')) expanded.set(section.dataset.shapeSection, section.open);
    const active = body.contains(document.activeElement) ? document.activeElement : null;
    const focusKey = controlKey(active), selection = active && typeof active.selectionStart === 'number' ? [active.selectionStart, active.selectionEnd] : null;
    const scroll = body.closest('.dialog-body'), scrollTop = scroll?.scrollTop, scrollLeft = scroll?.scrollLeft;
    const sh = getShape();
    if (!sh) { body.replaceChildren(el('p', { class: 'muted', role: 'status' }, '도형이 삭제되었거나 다른 문서·시트로 이동했습니다. 이 창을 닫고 도형을 다시 선택하세요.')); return; }
    const line = isLine(sh), pages = new Map();
    const stops = shapeGradientStops(sh); stopIndex = Math.min(stopIndex, stops.length - 1);
    const stop = stops[stopIndex];
    const mode = sh.pattern ? 'pattern' : sh.grad ? 'gradient' : sh.fill ? 'solid' : 'none';
    const fillPage = !line ? sec('채우기',
      modes('채우기', mode, [['none', '채우기 없음'], ['solid', '단색 채우기'], ['gradient', '그라데이션 채우기'], ['pattern', '무늬 채우기']], v => {
        const current = getShape();
        up({ fill: v === 'none' ? null : current.fill || '#4472c4', grad: v === 'gradient' ? current.grad ?? { ang: 90, stops: shapeGradientStops(current) } : undefined,
          pattern: v === 'pattern' ? current.pattern ?? { preset: 'pct25', fg: current.fill || '#4472c4', bg: '#ffffff' } : undefined }); draw();
      }),
      mode === 'solid' ? row('채우기 색', color(sh.fill, v => up({ fill: v }))) : null,
      mode !== 'none' ? row('채우기 투명도(%)', num((1 - (sh.fillOpacity ?? 1)) * 100, v => up({ fillOpacity: 1 - v / 100 }))) : null,
      mode === 'pattern' ? row('무늬 종류', choose(sh.pattern.preset, SHAPE_PATTERN_OPTIONS, v => up({ pattern: { ...getShape().pattern, preset: v } }))) : null,
      mode === 'pattern' ? row('무늬 색', color(sh.pattern.fg, v => up({ pattern: { ...getShape().pattern, fg: v } }))) : null,
      mode === 'pattern' ? row('무늬 배경색', color(sh.pattern.bg, v => up({ pattern: { ...getShape().pattern, bg: v } }))) : null,
      mode === 'gradient' ? el('div', {},
        row('그라데이션 방향(°)', num(sh.grad.ang ?? 90, v => up({ grad: { ...getShape().grad, ang: v } }), 0, 360)),
        row('그라데이션 중지점', choose(stopIndex, stops.map((s, i) => [i, `${i + 1} · ${Math.round(s[0] * 100)}%`]), v => { stopIndex = Number(v); draw(); })),
        el('div', { style: { display: 'flex', gap: '6px', margin: '6px 0' } },
          el('button', { class: 'btn', disabled: stops.length >= 16, onclick: () => { const next = addShapeGradientStop(getShape()); stopIndex = next.index; up({ grad: next.grad }); draw(); } }, '중지점 추가'),
          el('button', { class: 'btn', disabled: stops.length <= 2, onclick: () => { const current = getShape(); const next = shapeGradientStops(current); next.splice(stopIndex, 1); stopIndex = Math.min(stopIndex, next.length - 1); up({ grad: { ...current.grad, stops: next } }); draw(); } }, '중지점 제거')),
        row('중지점 위치(%)', num(stop[0] * 100, v => { up(shapeGradientStopPatch(getShape(), stopIndex, { position: v / 100 })); draw(); })),
        row('중지점 색', color(stop[1], v => up(shapeGradientStopPatch(getShape(), stopIndex, { color: v })))),
        row('중지점 투명도(%)', num((1 - (stop[2] ?? 1)) * 100, v => up(shapeGradientStopPatch(getShape(), stopIndex, { opacity: 1 - v / 100 }))))) : null) : null;
    const arrowRows = (end, label) => {
      const state = shapeArrowEnd(sh, end);
      const update = patch => up({ [end]: { ...shapeArrowEnd(getShape(), end), ...patch } });
      return [row(`${label} 화살표 종류`, choose(state.type, SHAPE_ARROW_OPTIONS, v => { update({ type: v }); draw(); })),
        ...(state.type !== 'none' ? [row(`${label} 화살표 너비`, choose(state.w, [['sm', '좁게'], ['med', '중간'], ['lg', '넓게']], v => update({ w: v }))),
          row(`${label} 화살표 길이`, choose(state.len, [['sm', '짧게'], ['med', '중간'], ['lg', '길게']], v => update({ len: v })))] : [])];
    };
    linePreview = el('div', { class: 'shape-line-preview', role: 'img', 'aria-label': '선 서식 미리 보기' });
    pages.set('채우기 및 선', el('div', {}, fillPage, sec('선',
      modes('선', sh.stroke ? 'solid' : 'none', [['none', '선 없음'], ['solid', '실선']], v => { up({ stroke: v === 'solid' ? getShape().stroke || '#2f528f' : null }); draw(); }),
      sh.stroke ? linePreview : null,
      ...(sh.stroke ? [row('선 색', color(sh.stroke, v => up({ stroke: v }))),
        row('선 투명도(%)', num((1 - (sh.strokeOpacity ?? 1)) * 100, v => up({ strokeOpacity: 1 - v / 100 }))),
        row('너비(pt)', num((sh.strokeWidth ?? 1) * .75, v => up({ strokeWidth: v / .75 }), .25, 100, .25)),
        row('복합 종류', choose(sh.compound ?? 'sng', [['sng', '단일선'], ['dbl', '이중선'], ['tri', '삼중선']], v => up({ compound: v }))),
        row('대시 종류', choose(sh.dash ?? '', SHAPE_DASH_OPTIONS, v => up({ dash: v || undefined }))),
        row('끝 모양', choose(sh.lineCap ?? 'flat', [['flat', '평평하게'], ['rnd', '둥글게'], ['sq', '사각형']], v => up({ lineCap: v }))),
        row('연결 모양', choose(sh.lineJoin ?? 'round', [['round', '둥글게'], ['bevel', '빗면'], ['miter', '각지게']], v => up({ lineJoin: v }))),
        ...(line ? [...arrowRows('headEnd', '시작'), ...arrowRows('tailEnd', '끝')] : [])] : []))));
    const sd = typeof sh.shadow === 'object' ? sh.shadow : {};
    pages.set('효과', el('div', {}, sec('그림자',
      row('그림자', check(sh.shadow, v => { up({ shadow: v ? { dx: 3, dy: 3, blur: 3, opacity: .4 } : undefined }); draw(); })),
      ...(sh.shadow ? [row('그림자 색', color(sd.color, v => up({ shadow: { ...getShape().shadow, color: v } }))),
        row('그림자 투명도(%)', num((1 - (sd.opacity ?? .4)) * 100, v => up({ shadow: { ...getShape().shadow, opacity: 1 - v / 100 } }))),
        row('그림자 흐리게(px)', num(sd.blur ?? 3, v => up({ shadow: { ...getShape().shadow, blur: v } }), 0, 50)),
        row('그림자 가로 거리(px)', num(sd.dx ?? 3, v => up({ shadow: { ...getShape().shadow, dx: v } }), -100, 100)),
        row('그림자 세로 거리(px)', num(sd.dy ?? 3, v => up({ shadow: { ...getShape().shadow, dy: v } }), -100, 100))] : [])),
      sec('네온 및 부드러운 가장자리',
        row('네온', check(sh.glow, v => { up({ glow: v ? { color: '#4472c4', size: 6, opacity: .6 } : undefined }); draw(); })),
        ...(sh.glow ? [row('네온 색', color(sh.glow.color, v => up({ glow: { ...getShape().glow, color: v } }))),
          row('네온 크기(px)', num(sh.glow.size ?? 6, v => up({ glow: { ...getShape().glow, size: v } }), 0, 50)),
          row('네온 투명도(%)', num((1 - (sh.glow.opacity ?? .6)) * 100, v => up({ glow: { ...getShape().glow, opacity: 1 - v / 100 } })))] : []),
        row('부드러운 가장자리(px)', num(sh.soft ?? 0, v => up({ soft: v || undefined }), 0, 30)))));
    pages.set('크기 및 속성', el('div', {}, sec('크기와 위치',
      row('가로 세로 비율 고정', check(sh.lockAspect, v => up({ lockAspect: v }))),
      row('너비(px)', num(sh.w, v => { up(shapeSizePatch(getShape(), 'w', v)); draw(); }, line ? 0 : 1, 10000, 1, sh.noMove)),
      row('높이(px)', num(sh.h, v => { up(shapeSizePatch(getShape(), 'h', v)); draw(); }, line ? 0 : 1, 10000, 1, sh.noMove)),
      row('가로 위치(px)', num(sh.x, v => up({ x: v }), 0, 1000000, 1, sh.noMove)),
      row('세로 위치(px)', num(sh.y, v => up({ y: v }), 0, 10000000, 1, sh.noMove)),
      row('회전(°)', num(sh.rot ?? 0, v => up({ rot: v || undefined }), -360, 360)),
      row('좌우 대칭', check(sh.flip, v => up({ flip: v || undefined }))),
      row('상하 대칭', check(sh.flipV, v => up({ flipV: v || undefined })))),
      sec('속성', row('셀에 따른 배치', choose(sh.placement ?? 'twoCell', [['twoCell', '위치와 크기 변함'], ['oneCell', '위치만 변함'], ['absolute', '위치와 크기 고정']], v => up({ placement: v }))),
        row('크기 조정 및 이동 사용 안 함', check(sh.noMove, v => { up({ noMove: v || undefined }); draw(); })),
        row('개체 인쇄', check(!sh.noPrint, v => up({ noPrint: !v || undefined }))),
        row('잠금', check(sh.locked !== false, v => up({ locked: v }))),
        el('p', { class: 'muted' }, '잠금은 시트 보호 중 적용됩니다.')),
      sec('대체 텍스트', row('설명', field(sh.alt, v => up({ alt: v || undefined }))))));
    if (!line) {
      textInput = el('textarea', { rows: 4, 'aria-label': '도형 텍스트', style: { width: '100%', boxSizing: 'border-box' } }, sh.text ?? '');
      textInput.addEventListener('change', () => textUp({ text: textInput.value }));
      const pad = sh.pad ?? [4.8, 9.6, 4.8, 9.6];
      pages.set('텍스트 및 글꼴', el('div', {}, sec('텍스트 및 글꼴', textInput,
        row('도형 글꼴', field(sh.font ?? fontName, v => textUp({ font: v || undefined }))),
        row('글자 크기(pt)', num(sh.size ?? 11, v => textUp({ size: v }), 6, 400)),
        row('글자 색', color(sh.color, v => textUp({ color: v }))),
        row('굵게', check(sh.bold, v => textUp({ bold: v || undefined }))),
        row('기울임꼴', check(sh.italic, v => textUp({ italic: v || undefined }))),
        row('밑줄', check(sh.underline, v => textUp({ underline: v || undefined }))))));
      pages.set('텍스트 상자', el('div', {}, sec('텍스트 상자', row('가로 맞춤', choose(sh.align ?? 'center', [['left', '왼쪽'], ['center', '가운데'], ['right', '오른쪽'], ['justify', '양쪽 맞춤']], v => textUp({ align: v }))),
          row('세로 맞춤', choose(sh.valign ?? (sh.kind === 'textbox' ? 'top' : 'middle'), [['top', '위쪽'], ['middle', '가운데'], ['bottom', '아래쪽']], v => up({ valign: v }))),
          row('텍스트 회전', choose(sh.textRot ?? 0, [[0, '가로'], [90, '시계 방향 90°'], [270, '시계 방향 270°']], v => up({ textRot: Number(v) }))),
          row('텍스트 자동 맞춤', choose(sh.textFit ?? 'none', [['none', '자동 맞춤 안 함'], ['shrink', '넘치면 텍스트 축소']], v => up({ textFit: v }))),
          row('도형에서 텍스트 줄 바꿈', check(!sh.nowrap, v => up({ nowrap: !v || undefined }))),
          ...['위쪽 여백(px)', '오른쪽 여백(px)', '아래쪽 여백(px)', '왼쪽 여백(px)'].map((name, i) => row(name, num(pad[i], v => { const next = [...(getShape().pad ?? pad)]; next[i] = v; up({ pad: next }); }, 0, 200, .5))))));
    }
    const shapeNames = ['채우기 및 선', '효과', '크기 및 속성'], textNames = ['텍스트 및 글꼴', '텍스트 상자'];
    if (!pages.has(selected)) selected = shapeNames[0];
    const isText = name => textNames.includes(name);
    const makeTab = (name, activate, names, owner, current, prefix) => el('button', { type: 'button', role: 'tab',
      'aria-label': name, 'data-shape-focus': prefix + name, 'aria-selected': String(name === current), tabindex: name === current ? 0 : -1,
      onclick: () => { activate(name); [...owner.children].find(node => node.dataset.shapeFocus === prefix + name)?.focus({ preventScroll: true }); }, onkeydown: event => {
        if (!['ArrowLeft', 'ArrowRight', 'Home', 'End'].includes(event.key)) return; event.preventDefault(); event.stopPropagation();
        const at = names.indexOf(name), next = event.key === 'Home' ? 0 : event.key === 'End' ? names.length - 1 : (at + (event.key === 'ArrowRight' ? 1 : -1) + names.length) % names.length;
        activate(names[next]); owner.children[next]?.focus();
      }
    }, name);
    const show = name => {
      selected = name; const text = isText(name); if (text) lastText = name; else lastShape = name;
      for (const [key, page] of pages) { page.hidden = key !== name; page.setAttribute('role', 'tabpanel'); page.setAttribute('aria-label', key); }
      const groups = line ? ['도형 옵션'] : ['도형 옵션', '텍스트 옵션'], group = text ? '텍스트 옵션' : '도형 옵션';
      options.replaceChildren(...groups.map(label => makeTab(label, label => show(label === '텍스트 옵션' ? lastText : lastShape), groups, options, group, 'option:')));
      const names = text ? textNames : shapeNames;
      tabs.setAttribute('aria-label', text ? '텍스트 서식 범주' : '도형 서식 범주');
      tabs.replaceChildren(...names.map(label => makeTab(label, show, names, tabs, selected, 'category:')));
    };
    body.replaceChildren(navigation, el('p', { class: 'muted shape-format-note' }, '변경 내용은 즉시 적용됩니다. 닫기는 취소가 아니며 실행 취소(Ctrl+Z)로 되돌릴 수 있습니다.'), ...pages.values());
    show(selected); preview();
    if (focusKey) {
      const next = [...body.querySelectorAll('input,select,textarea,button,summary')].find(node => controlKey(node) === focusKey && !node.disabled && !node.closest('[hidden]')) || tabs.querySelector('[aria-selected="true"]');
      if (next) { next.focus({ preventScroll: true }); if (selection && typeof next.setSelectionRange === 'function') next.setSelectionRange(...selection); }
    }
    if (scroll) { scroll.scrollTop = scrollTop; scroll.scrollLeft = scrollLeft; }

  };
  // 모델리스 창이 열린 채 문서를 바꾼 뒤의 늦은 change/click은 이전 도형에 쓰지 않는다.
  for (const name of ['change', 'click']) body.addEventListener(name, event => { if (!getShape()) { event.preventDefault(); event.stopImmediatePropagation(); draw(); } }, true);
  draw();
  return { body, refresh: draw, getSelectedTab: () => selected, focusText: () => { textInput?.focus(); textInput?.setSelectionRange(textInput.value.length, textInput.value.length); } };
}
