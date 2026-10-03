// 통합 문서의 이름 있는 표·피벗·슬라이서 스타일. DOM/전역 테마/통합 문서를 변경하지 않는다.
export const TABLE_STYLE_ELEMENTS = ['wholeTable','headerRow','totalRow','firstColumn','lastColumn','firstRowStripe','secondRowStripe','firstColumnStripe','secondColumnStripe','firstHeaderCell','lastHeaderCell','firstTotalCell','lastTotalCell','firstSubtotalRow','secondSubtotalRow','thirdSubtotalRow','firstSubtotalColumn','secondSubtotalColumn','thirdSubtotalColumn','firstRowSubheading','secondRowSubheading','thirdRowSubheading','firstColumnSubheading','secondColumnSubheading','thirdColumnSubheading','pageFieldLabels','pageFieldValues','blankRow'];
export const SLICER_STYLE_ELEMENTS = ['wholeTable','headerRow','selectedItemWithData','selectedItemWithNoData','unselectedItemWithData','unselectedItemWithNoData','hoveredSelectedItemWithData','hoveredSelectedItemWithNoData','hoveredUnselectedItemWithData','hoveredUnselectedItemWithNoData'];
export const objectStyleKey = name => String(name ?? '').trim().toLowerCase();
const clone = value => structuredClone(value);
function definition(value, slicer) {
  if (!value || !objectStyleKey(value.name)) return null;
  const name = String(value.name).trim();
  if (name.length > 255) throw new Error('스타일 이름은 255자 이내로 입력하세요.');
  const elements = [], seen = new Set();
  for (const e of value.elements ?? []) {
    if (!e || typeof e.type !== 'string' || seen.has(e.type)) continue;
    seen.add(e.type);
    const entry = {type:e.type,style:clone(e.style ?? {})};
    if (Number.isInteger(e.size) && e.size > 0 && e.size <= 4294967295) entry.size=e.size;
    // 읽은 표준 DXF의 지원 밖 속성과 테마 참조도 서식을 편집하기 전까지 보존한다.
    if (e.sourceDxf && e.sourceStyle) {entry.sourceDxf=clone(e.sourceDxf);entry.sourceStyle=clone(e.sourceStyle);}
    elements.push(entry);
  }
  return {name,...(slicer?{}:{table:value.table!==false,pivot:value.pivot!==false}),elements};
}
export function normalizeObjectStyles(registry) {
  const out={tables:[],slicers:[]};
  for (const group of ['tables','slicers']) {
    const seen=new Set();
    for(const value of registry?.[group] ?? []) {const d=definition(value,group==='slicers');if(!d||seen.has(objectStyleKey(d.name)))continue;seen.add(objectStyleKey(d.name));out[group].push(d);}
  }
  for(const key of ['defaultTableStyle','defaultPivotStyle','defaultSlicerStyle'])if(typeof registry?.[key]==='string')out[key]=registry[key];
  return out;
}
export function findObjectStyle(registry,kind,name) {
  const group=kind==='slicer'?'slicers':'tables',key=objectStyleKey(name);
  return registry?.[group]?.find(d=>objectStyleKey(d.name)===key&&(kind==='slicer'||d[kind]!==false))??null;
}
export function upsertObjectStyle(registry,kind,value,options={}) {
  const out=normalizeObjectStyles(registry),group=kind==='slicer'?'slicers':'tables';
  const d=definition({...value,...(kind==='slicer'?{}:{table:value.table??kind==='table',pivot:value.pivot??kind==='pivot'})},kind==='slicer');
  if(!d)throw new Error('스타일 이름을 입력하세요.');
  const foreign=group==='tables'?out.slicers:out.tables.filter(x=>x.table!==false||x.pivot!==false);
  if(foreign.some(x=>objectStyleKey(x.name)===objectStyleKey(d.name)))throw new Error('다른 종류에 같은 이름의 스타일이 이미 있습니다.');
  const i=out[group].findIndex(x=>objectStyleKey(x.name)===objectStyleKey(d.name));
  if(i>=0&&!options.replace)throw new Error('같은 이름의 스타일이 이미 있습니다.');
  if(i<0)out[group].push(d);else out[group][i]=d;
  return out;
}
export function objectStylePatch(kind,style) {
  if(!style)return clearObjectStyle(kind);
  return {style:style.name,styleElements:clone(style.elements??[]),styleDef:undefined,...(kind==='slicer'?{custom:undefined,color:undefined}:{})};
}
export function clearObjectStyle(kind) {
  return {style:'None',styleElements:undefined,styleDef:undefined,...(kind==='slicer'?{custom:undefined,color:undefined}:{})};
}

/** 한 스타일 원소의 채우기는 전체 채우기 채널을 대체한다. null은 명시적인 채우기 없음이다. */
export function mergeObjectStyle(...styles) {
  const out={};
  for(const st of styles)if(st){
    if(st.fill!==undefined||st.gradient!==undefined||st.pattern!==undefined)for(const k of ['fill','gradient','pattern','patternColor'])delete out[k];
    Object.assign(out,st);
  }
  return out;
}
