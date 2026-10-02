import { el } from './ui.js';
const drawingIcons = {
  select:'<path d="M5 3v16l4-5 5 7 3-2-5-7 6-1Z"/>',
  pen:'<path d="m5 15 9-11 5 4-9 11-6 1Z"/><path d="m12 6 5 4M5 15l5 4"/>',
  marker:'<path d="m7 13 7-10 6 5-8 9-5-4ZM7 13l-3 5 6 2 2-3M3 22h17"/>',
  line:'<path d="m4 20 16-16"/>',rect:'<rect x="4" y="5" width="16" height="14" rx="1"/>',
  ellipse:'<circle cx="12" cy="12" r="8"/>',arrow:'<path d="M4 19 19 4M9 4h10v10"/>',
  eraser:'<path d="m3 14 9-10 9 8-8 9H9Z"/><path d="m7 10 9 8M13 21h9"/>',
};
const drawingIcon = key => `<svg viewBox="0 0 24 24" width="22" height="22" fill="none" stroke="currentColor" stroke-width="1.6" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">${drawingIcons[key]}</svg>`;
export function createDrawingPalette({ onTool, onSettings, onUndo, onRedo, onClose }) {
  const tools = [['select','선택','selectAll'],['pen','펜','painter'],['marker','형광펜','fill'],['line','선','border'],['rect','사각형','shapes'],['ellipse','원','shapes'],['arrow','화살표','bringForward'],['eraser','획 지우개','clear']];
  const buttons = [], state = { color: '#185c45', width: 3, opacity: 1 };
  const group = el('div', { class: 'draw-tools', role: 'toolbar', 'aria-label': '그리기 도구' }, tools.map(([key,label,icon]) => {
    const button = el('button', { type:'button', title:label, 'aria-label':label, 'aria-pressed':'false', onclick:() => onTool(key) }, el('span',{html:drawingIcon(key)}),el('small',{},key==='eraser'?'지우개':label)); buttons.push([key,button]); return button;
  }));
  const color = el('input',{type:'color',value:state.color,'aria-label':'펜 색',oninput:() => { state.color=color.value; emit(); }});
  const width = el('input',{type:'range',min:1,max:30,step:1,value:3,'aria-label':'펜 굵기',oninput:() => { state.width=Number(width.value); size.textContent=`${state.width} px`; emit(); }});
  const opacity = el('input',{type:'range',min:10,max:100,step:5,value:100,'aria-label':'펜 불투명도',oninput:() => { state.opacity=Number(opacity.value)/100; alpha.textContent=`${opacity.value}%`; emit(); }});
  const size=el('output',{},'3 px'),alpha=el('output',{},'100%');
  const emit=() => onSettings({...state});
  const root=el('aside',{class:'drawing-palette','aria-label':'그리기 팔레트'},
    el('header',{},el('b',{},'그리기'),el('button',{type:'button','aria-label':'그리기 팔레트 닫기',onclick:onClose},'×')),group,
    el('div',{class:'draw-swatches'},['#185c45','#2456a6','#7c3aed','#dc3545','#f59e0b','#111827','#ffffff'].map(c=>el('button',{type:'button',style:{background:c},'aria-label':`펜 색 ${c}`,onclick:()=>{state.color=c;color.value=c;emit();}})),color),
    el('label',{},'굵기',width,size),el('label',{},'불투명도',opacity,alpha),
    el('footer',{},el('button',{type:'button',class:'btn',onclick:onUndo},'실행 취소'),el('button',{type:'button',class:'btn',onclick:onRedo},'다시 실행')),
    el('p',{class:'muted'},'펜·터치로 연속 그리기 · Esc: 선택으로 · 지우개는 클릭한 획만 삭제'));
  document.body.append(root); emit();
  return { root, select(key){for(const[k,b]of buttons){b.classList.toggle('on',k===key);b.setAttribute('aria-pressed',String(k===key));}}, close(){root.remove();} };
}
