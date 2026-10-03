// 이름 있는 표·피벗·슬라이서 스타일의 요소별 편집기. 취소 전까지 문서에는 쓰지 않습니다.
import { el } from './ui.js';
import { tableCellStyle } from './tables.js';
import { slicerColors } from './slicerstyle.js';
import { fontCss, borderCss, gradientCss, patternCss } from './view.js';
import { formatValue } from './format.js';

const TABLE_ELEMENTS = [
  ['wholeTable', '전체 표'], ['headerRow', '머리글 행'], ['totalRow', '요약 행'],
  ['firstColumn', '첫째 열'], ['lastColumn', '마지막 열'],
  ['firstRowStripe', '첫째 행 줄무늬'], ['secondRowStripe', '둘째 행 줄무늬'],
  ['firstColumnStripe', '첫째 열 줄무늬'], ['secondColumnStripe', '둘째 열 줄무늬'],
  ['firstHeaderCell', '첫째 머리글 셀'], ['lastHeaderCell', '마지막 머리글 셀'],
  ['firstTotalCell', '첫째 요약 셀'], ['lastTotalCell', '마지막 요약 셀'],
];
const PIVOT_ELEMENTS = [
  ['wholeTable', '전체 표'], ['headerRow', '머리글 행'], ['firstColumn', '첫째 열'],
  ['firstRowStripe', '첫째 행 줄무늬'], ['secondRowStripe', '둘째 행 줄무늬'],
  ['firstColumnStripe', '첫째 열 줄무늬'], ['secondColumnStripe', '둘째 열 줄무늬'],
  ['firstHeaderCell', '첫째 머리글 셀'], ['firstSubtotalRow', '첫째 부분합 행'],
  ['secondSubtotalRow', '둘째 부분합 행'], ['thirdSubtotalRow', '셋째 부분합 행'],
  ['firstSubtotalColumn', '첫째 부분합 열'], ['secondSubtotalColumn', '둘째 부분합 열'],
  ['thirdSubtotalColumn', '셋째 부분합 열'], ['firstRowSubheading', '첫째 행 하위 제목'],
  ['secondRowSubheading', '둘째 행 하위 제목'], ['thirdRowSubheading', '셋째 행 하위 제목'],
  ['firstColumnSubheading', '첫째 열 하위 제목'], ['secondColumnSubheading', '둘째 열 하위 제목'],
  ['thirdColumnSubheading', '셋째 열 하위 제목'], ['blankRow', '빈 행'],
  ['totalRow', '총합계 행'], ['lastColumn', '총합계 열'],
  ['pageFieldLabels', '보고서 필터 필드 레이블'], ['pageFieldValues', '보고서 필터 필드 값'],
];
const SLICER_ELEMENTS = [
  ['wholeTable', '전체 슬라이서'], ['headerRow', '머리글'],
  ['selectedItemWithData', '선택한 항목 · 데이터 있음'], ['selectedItemWithNoData', '선택한 항목 · 데이터 없음'],
  ['unselectedItemWithData', '선택하지 않은 항목 · 데이터 있음'], ['unselectedItemWithNoData', '선택하지 않은 항목 · 데이터 없음'],
  ['hoveredSelectedItemWithData', '가리킨 선택 항목 · 데이터 있음'], ['hoveredSelectedItemWithNoData', '가리킨 선택 항목 · 데이터 없음'],
  ['hoveredUnselectedItemWithData', '가리킨 미선택 항목 · 데이터 있음'], ['hoveredUnselectedItemWithNoData', '가리킨 미선택 항목 · 데이터 없음'],
];
const FORMAT_KEYS = new Set(['font', 'size', 'color', 'bold', 'italic', 'underline', 'strike', 'fill', 'pattern', 'patternColor', 'gradient', 'numFmt', 'code', 'decimals', 'thousands', 'negative', 'currency', 'symbol', 'bt', 'btc', 'bts', 'bb', 'bbc', 'bbs', 'bl', 'blc', 'bls', 'br', 'brc', 'brs', 'du', 'duc', 'dus', 'dd', 'ddc', 'dds', 'bh', 'bhc', 'bhs', 'bv', 'bvc', 'bvs']);
const cleanFormat = style => Object.fromEntries(Object.entries(style ?? {}).filter(([key, value]) => FORMAT_KEYS.has(key)));
function cellCss(st) {
  const border=(side,key)=>st[key]?borderCss(side,st[key+'s'],st[key+'c']).split(':').slice(1).join(':'):'';
  return {background:st.gradient?gradientCss(st.gradient):st.pattern?patternCss(st.pattern,st.patternColor||'#000',st.fill):st.fill||'#fff',
    color:st.color||'#222',font:fontCss({...st,size:Math.min(st.size||10,18)}),textDecoration:[st.underline?'underline':'',st.strike?'line-through':''].join(' '),
    borderTop:border('top','bt'),borderBottom:border('bottom','bb'),borderLeft:border('left','bl'),borderRight:border('right','br')};
}
/** onFormat은 기존 셀 서식 창으로 연결합니다. 저장은 호출자가 한 트랜잭션에서 처리합니다. */
export function objectStyleEditor({kind, definition, onFormat, defaultStyle = false}) {
  const draft = structuredClone(definition), elements = draft.elements ??= [];
  const known = kind === 'slicer' ? SLICER_ELEMENTS : kind === 'pivot' ? PIVOT_ELEMENTS : TABLE_ELEMENTS;
  const rows = [...known, ...elements.filter(e => !known.some(([type]) => type === e.type)).map(e => [e.type, e.type])];
  const name = el('input', {type:'text', value:draft.name || '', maxlength:255, 'aria-label':'스타일 이름', autocomplete:'off'});
  const defaults = el('input', {type:'checkbox', checked:defaultStyle, 'aria-label':'이 문서의 기본 스타일로 설정'});
  const list = el('div', {class:'object-style-elements', role:'listbox', 'aria-label':'스타일 요소'});
  const preview = el('div', {class:'object-style-preview', 'aria-label':'스타일 미리 보기'});
  const elementPreview = el('div', {class:'object-style-element-preview','aria-label':'선택한 요소 미리 보기'});
  const detail = el('div', {class:'object-style-description'});
  const size = el('input', {type:'number', value:1, min:1, max:9, 'aria-label':'줄무늬 크기'});
  const sizeRow = el('label', {class:'object-style-stripe'}, el('span', {}, '줄무늬 크기'), size);
  const clear = el('button', {type:'button', class:'btn'}, '요소 서식 지우기(C)');
  const format = el('button', {type:'button', class:'btn'}, '서식(F)...');
  let selected = rows[0][0];
  const element = type => elements.find(e => e.type === type);
  const refresh = () => {
    for (const b of list.children) { const on=b.dataset.element===selected; b.setAttribute('aria-selected',String(on)); b.tabIndex=on?0:-1; b.classList.toggle('on',on); b.classList.toggle('defined',!!element(b.dataset.element)); }
    const entry=element(selected), st=entry?.style ?? {};
    sizeRow.hidden=!/Stripe$/.test(selected); size.value=String(entry?.size ?? 1); clear.disabled=!entry;
    const descriptions=[];
    if(st.bold)descriptions.push('굵게'); if(st.italic)descriptions.push('기울임'); if(st.color)descriptions.push(`글자 ${st.color}`);
    if(st.fill)descriptions.push(`채우기 ${st.fill}`); if(st.gradient)descriptions.push('그라데이션');
    if(st.bt||st.bb||st.bl||st.br)descriptions.push('테두리'); if(st.numFmt)descriptions.push(`표시 형식 ${st.numFmt}`);
    detail.textContent=descriptions.join(' · ')||'지정된 서식 없음';
    elementPreview.removeAttribute('style');Object.assign(elementPreview.style,cellCss(st));
    try {elementPreview.textContent=`가나다 Aa · ${formatValue(1234.5,st).text}`;}catch {elementPreview.textContent='가나다 Aa · 1,234.50';}
    if(kind==='slicer') {
      const c=slicerColors({style:draft.name,styleElements:elements});
      preview.replaceChildren(el('div',{class:'object-style-slicer',style:{background:c.frame,border:`1px solid ${c.border}`}},
        el('b',{style:{color:c.head,borderBottom:`1px solid ${c.headLine||c.border}`}},'지역'),
        ...[true,false,true,false].map((on,i)=>el('div',{style:{background:on?c.selFill:c.item,color:on?c.selText:c.itemText,border:`1px solid ${on?c.selBorder:c.itemBorder}`}},['서울','부산','대구','제주'][i]))));
    } else {
      const t={style:draft.name,styleElements:elements,r1:0,c1:0,r2:5,c2:2,header:true,totals:true,banded:true,firstCol:true};
      const grid = Array.from({length:6},(_,r) => el('tr', {}, Array.from({length:3},(_,c) => {
        const text = r===0 ? ['지역','수량','매출'][c] : r===5 ? (c===0?'합계':c===1?'30':'90,000') : c===0 ? ['서울','부산','대구','제주'][r-1] : c===1 ? String(r*3) : `${r*9},000`;
        return el('td', {style:cellCss(tableCellStyle(t,r,c)||{})}, text);
      })));
      preview.replaceChildren(el('table',{},el('tbody',{},grid)));
    }
  };
  const choose = (type, focus=false) => {selected=type;refresh();if(focus){const b=[...list.children].find(n=>n.dataset.element===type);b.focus({preventScroll:true});b.scrollIntoView({block:'nearest'});}};
  for(const [type,label] of rows) list.append(el('button',{type:'button',role:'option','data-element':type,'data-access-key':'none',onclick:()=>choose(type)},label));
  list.addEventListener('keydown',e=>{if(!['ArrowUp','ArrowDown','Home','End'].includes(e.key))return;e.preventDefault();e.stopPropagation();const at=rows.findIndex(([t])=>t===selected),i=e.key==='Home'?0:e.key==='End'?rows.length-1:Math.max(0,Math.min(rows.length-1,at+(e.key==='ArrowDown'?1:-1)));choose(rows[i][0],true);});
  size.addEventListener('change',()=>{const value=Number(size.value);if(!Number.isInteger(value)||value<1||value>9){size.value=String(element(selected)?.size??1);return;}let entry=element(selected);if(!entry){entry={type:selected,style:{}};elements.push(entry);}entry.size=value;refresh();});
  clear.addEventListener('click',()=>{const at=elements.findIndex(e=>e.type===selected);if(at>=0)elements.splice(at,1);refresh();});
  format.addEventListener('click',()=>{
    const type=selected,own=element(type)?.style??{};
    const parent=type.startsWith('hovered')?type.replace(/^hovered([A-Z])/,(_,c)=>c.toLowerCase()):type.includes('NoData')?type.replace('NoData','Data'):/(?:first|last)HeaderCell/.test(type)?'headerRow':/(?:first|last)TotalCell/.test(type)?'totalRow':null;
    const inherited=type==='wholeTable'?{}:{...(element('wholeTable')?.style??{}),...(element(parent)?.style??{})};
    onFormat(structuredClone({...inherited,...own}), style=>{
      const patch=cleanFormat(style);if(!Object.keys(patch).length)return;
      let entry=element(type);if(!entry){entry={type,style:{}};elements.push(entry);}
      entry.style={...entry.style,...patch};for(const key of Object.keys(entry.style))if(entry.style[key]===undefined)delete entry.style[key];refresh();
    });
  });
  refresh();
  const body=el('div',{class:'object-style-editor'},el('label',{class:'object-style-name'},el('span',{},'이름(N)'),name),
    el('div',{class:'object-style-columns'},el('div',{},el('b',{},'스타일 요소(E)'),list),el('div',{class:'object-style-detail'},el('b',{},'미리 보기'),preview,elementPreview,sizeRow,detail,el('div',{class:'object-style-actions'},format,clear))),
    el('label',{class:'fc-check'},defaults,'이 문서의 기본 스타일로 설정(D)'));
  return {body, nameInput:name, read:()=>({...structuredClone(draft),name:name.value.trim()}), isDefault:()=>defaults.checked};
}
