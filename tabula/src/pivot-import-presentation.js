import { pivotPageLayout, pivotSourceData, normalizeDef } from './pivot.js';
import { importedPivotSourceReference, pivotSourceReferenceCurrent, pivotCacheLayoutBinding } from './pivot-source-reference.js';
import { capturePivotCellFormat } from './pivot-style-format.js';
import { columnKeysInRange } from './column-keys-in-range.js';
import { blockValue } from './block.js';

const areaKeys = ['r1', 'c1', 'r2', 'c2'];
const validArea = area => area && areaKeys.every(key => Number.isSafeInteger(area[key]) && area[key] >= 0) && area.r1 <= area.r2 && area.c1 <= area.c2;
const inside = (area, r, c) => area && r >= area.r1 && r <= area.r2 && c >= area.c1 && c <= area.c2;
const sourceKey = def => {
  const ref = importedPivotSourceReference(def);
  return JSON.stringify([def.source == null ? null : String(def.source).toLowerCase(), def.table == null ? null : String(def.table).toLowerCase(),
    def.range ? areaKeys.map(key => def.range[key]) : null, def.cacheItemsId ?? null, def.snapshotId ?? null,
    ref ? [ref.external ?? null, ref.sheet ?? null, ref.name ?? null, ref.ref ?? null] : null]);
};
/** XLSX reader만 초기 정규화 완료 후 이 메타데이터를 확정합니다. */
export function bindImportedPivotPresentation(def, presentation) {
  presentation.source = sourceKey(def);
  presentation.binding = pivotCacheLayoutBinding(def);
  presentation.relativeArea = validArea(def.area) ? relativeArea(def) : null;
}
export function importedPivotPresentationCurrent(def, presentation = def.importedPresentation) {
  if (presentation?.kind !== 'xlsx-pivot-output' || presentation.source !== sourceKey(def) || !validArea(def.area)
    || !Array.isArray(presentation.relativeArea) || !relativeArea(def).every((value,i)=>value===presentation.relativeArea[i])) return false;
  const reference = importedPivotSourceReference(def);
  if (reference && !pivotSourceReferenceCurrent(def)) return false;
  const binding = pivotCacheLayoutBinding(def);
  return binding.fields === presentation.binding?.fields && binding.query === presentation.binding?.query;
}
/** 별도 배열을 만들지 않고 필드 항목 순번의 의미/순서를 비교합니다. */
export function pivotItemOrderSignature(keys) {
  let a=2166136261,b=2654435769,n=0;
  for (const key of keys) {
    const text = typeof key + ':' + String(key), token = text.length + ':' + text + ';';
    for (let i=0;i<token.length;i++) { const c=token.charCodeAt(i); a=Math.imul(a^c,16777619);b=Math.imul(b^c,2246822507); }
    n++;
  }
  return n + ':' + (a>>>0).toString(16) + ':' + (b>>>0).toString(16);
}
const relativeArea = def => {
  const a = def.area, top = def.top ?? 0, left = def.left ?? 0;
  return [a.r1 - top, a.c1 - left, a.r2 - top, a.c2 - left];
};

/** 기존 쓰기와 가져온 머리글에 같은 필터 단추 판정을 적용합니다. */
export function pivotCellButtons(cd, r, c, pm, d, { includeToggle = true } = {}) {
  if (!cd) return [];
  const out = [];
  if (cd.role === 'rowHead:0' && d.rows.length) out.push({ r, c, kind: 'rows', ...(d.layout !== 'compact' ? { field: d.rows[0] } : {}) });
  else if (/^rowHead:\d+$/.test(cd.role) && d.layout !== 'compact' && d.rows[+cd.role.split(':')[1]]) out.push({ r, c, kind: 'rows', field: d.rows[+cd.role.split(':')[1]] });
  else if (d.valuesOnRows && d.values.length > 1 && cd.role === `rowHead:${d.rows.length}`) out.push({ r, c, kind: 'rows', sigma: true });
  else if (cd.role === 'colHead' && cd.raw) {
    const level = c - (pm.left ?? 0) - pm.labelCols, multi = d.values.length > 1 && !d.valuesOnRows;
    const vp = multi ? Math.min(d.cols.length, d.valuesPos ?? d.cols.length) : -1;
    const sigma = multi && (!d.cols.length || d.layout !== 'compact' && level === vp);
    const field = d.layout !== 'compact' && !sigma ? d.cols[vp >= 0 && level > vp ? level - 1 : level] : null;
    out.push({ r, c, kind: 'cols', ...(field ? { field } : {}), ...(sigma ? { sigma: true } : {}) });
  } else if (cd.role === 'pageValue') out.push({ r, c, kind: 'page', field: cd.field ?? d.pages[r - (pm.top ?? 0)] });
  if (includeToggle && cd.toggle) out.push({ r, c, kind: 'toggle', field: cd.toggle.field, item: cd.toggle.item, collapsed: cd.toggle.collapsed });
  return out;
}

/** 저장된 본문을 계산하지 않고 위치 메타데이터/원본 머리글만 준비합니다. */
export function pivotImportedPresentation(book, targetSi, def) {
  const saved=def.importedPresentation;
  if (!def.captureFmt || def.needsRender || def.refreshOnOpen === true || !validArea(def.area)
    || !importedPivotPresentationCurrent(def,saved) || !validArea(saved.body)) return null;
  const snap=def.snapshotId&&book.pivotSnapshots?.get(def.snapshotId);
  if (!snap||!book.pivotSnapshotCurrent({...snap},def)) return null;
  const src=pivotSourceData(book,def,{preserveSnapshot:true});
  if (!src||src.metadataOnly) return null;
  const d=normalizeDef(def,src.cube.header),top=def.top??0,left=def.left??0;
  const dr=top-saved.top,dc=left-saved.left,body={r1:saved.body.r1+dr,c1:saved.body.c1+dc,r2:saved.body.r2+dr,c2:saved.body.c2+dc};
  const headerRows=Number(saved.location.firstDataRow),labelCols=Number(saved.location.firstDataCol);
  if (!Number.isSafeInteger(headerRows)||headerRows<1||headerRows>Math.max(2,d.cols.length+3)
    ||!Number.isSafeInteger(labelCols)||labelCols<0||labelCols>body.c2-body.c1+1
    ||!inside(def.area,body.r1,body.c1)||!inside(def.area,body.r2,body.c2)) return null;
  const cellAt=(r,c)=>book.getCell(targetSi,r,c);
  const hasRaw=cell=>cell?.raw!==undefined&&cell.raw!==null&&cell.raw!=='';
  let hasBody=false;
  for(let c=body.c1;c<=body.c2&&!hasBody;c++){
    if(hasRaw(cellAt(body.r1+headerRows,c))){hasBody=true;break;}
    const col=book.sheets[targetSi]?.cells.col(c);
    if(!col)continue;
    if(col.storageEntries){
      for(const [r,cell,count] of col.storageEntries())if(r<=body.r2&&r+count>body.r1+headerRows&&hasRaw(cell)){hasBody=true;break;}
    }else for(const r of columnKeysInRange(col,body.r1+headerRows,body.r2))if(hasRaw(col.get(r))){hasBody=true;break;}
  }
  if(!hasBody)for(const block of book.sheets[targetSi]?.blocks??[]){
    for(let c=Math.max(body.c1,block.c0);c<=Math.min(body.c2,block.c0+block.cols.length-1)&&!hasBody;c++)
      for(let r=Math.max(body.r1+headerRows,block.r0);r<=Math.min(body.r2,block.r0+block.n-1);r++){
        if(blockValue(block,r,c)!==null&&hasRaw(cellAt(r,c))){hasBody=true;break;}
      }
    if(hasBody)break;
  }
  if(!hasBody)return null;
  const t=book.sheets[targetSi],page=pivotPageLayout(d),pageRows=body.r1-top;
  if(pageRows!==(page.height?page.height+1:0))return null;
  const grid=Array.from({length:pageRows+headerRows},()=>[]);
  const add=(r,c,role,field)=>{
    if(!inside(def.area,r,c))return;
    const cell=cellAt(r,c);if(!cell)return;
    grid[r-top][c-left]={raw:cell.raw,role,...(field?{field}:{}),style:{}};
  };
  for(const point of page.fields){add(top+point.r,left+point.c,'pageLabel',point.field);add(top+point.r,left+point.c+1,'pageValue',point.field);}
  if(d.showHeaders!==false){
    for(let c=0;c<labelCols;c++)add(body.r1+headerRows-1,body.c1+c,'rowHead:'+c);
    const multi=d.values.length>1&&!d.valuesOnRows;
    const levels=d.cols.length+(multi?1:0);
    if(d.cols.length||multi&&d.valuesHeadRow){
      const count=d.layout==='compact'?1:levels;
      for(let c=0;c<count;c++)add(body.r1,body.c1+labelCols+c,'colHead');
    }
  }
  const marker={source:sourceKey(def),binding:pivotCacheLayoutBinding(def),relativeArea:relativeArea(def),
    bodyRow:body.r1-top,bodyColumnStart:body.c1-left,bodyColumnEnd:body.c2-left,pageFields:page.fields.map(({r,c})=>({r,c}))};
  return {marker,plan:{src,d,t,top,left,grid,cellAt,pm:{pageRows,headerRows,labelCols}}};
}

/** 문서별 마커는 지우지 않아 Undo에서 복원된 원본 정의도 같은 소유권을 갖습니다. */
export function pivotImportedPresentationCurrent(def, marker) {
  if (!marker || !def.captureFmt || def.needsRender || !validArea(def.area) || marker.source !== sourceKey(def)) return false;
  if (importedPivotSourceReference(def) && !pivotSourceReferenceCurrent(def)) return false;
  if (!relativeArea(def).every((value, i) => value === marker.relativeArea[i])) return false;
  const binding = pivotCacheLayoutBinding(def);
  return binding.fields === marker.binding.fields && binding.query === marker.binding.query;
}

/** 보고서 필터 사이 빈 열/본문 위의 빈 줄은 사용자 영역입니다. */
export function pivotImportedOwnsCell(def, marker, r, c) {
  if (!inside(def.area, r, c)) return false;
  const dr = r - (def.top ?? 0), dc = c - (def.left ?? 0);
  if (dr >= marker.bodyRow && dc >= marker.bodyColumnStart && dc <= marker.bodyColumnEnd) return true;
  return marker.pageFields.some(point => dr === point.r && (dc === point.c || dc === point.c + 1));
}

/** 기존 역할별 투표 규칙을 유지하면서 가져온 원본 범위 밖 서식을 읽지 않습니다. */
export function captureImportedPivotFormats(sheet, def, plan, captureBase, { area = null, headersOnly = false } = {}) {
  const fmt = { ...(def.cellFmt ?? {}) }, votes = new Map(), voteMemo = new WeakMap();
  const { grid, pm, top, left } = plan;
  grid.forEach((row, r) => row.forEach((cd, c) => {
    if (!cd?.role || cd.role === 'empty' || def.cellFmt?.[cd.role] || area && !inside(area, top + r, left + c)
      || headersOnly && r >= pm.pageRows + pm.headerRows) return;
    const fc = plan.cellAt ? plan.cellAt(top+r,left+c) : sheet.cells.getRC(top + r, left + c), own = fc?.style;
    let f = null, key;
    if (own && Object.keys(own).length) {
      const generatedKey = JSON.stringify(cd.style ?? {});
      let pair = voteMemo.get(own);
      if (!pair) voteMemo.set(own, (pair = new Map()));
      let hit = pair.get(generatedKey);
      if (!hit) { const ff = capturePivotCellFormat(own, cd.style, captureBase); hit = { f: ff, key: JSON.stringify(ff) }; pair.set(generatedKey, hit); }
      f = hit.f; key = hit.key;
    } else if (fc && fc.raw !== '' && cd.style?.numFmt) f = { numFmt: 'general' };
    else if (!fc) return;
    key ??= JSON.stringify(f);
    let m = votes.get(cd.role);
    if (!m) votes.set(cd.role, (m = new Map()));
    const e = m.get(key);
    if (e) e.n++; else m.set(key, { n: 1, f });
  }));
  for (const [role, m] of votes) {
    let best = null;
    for (const e of m.values()) if (!best || e.n > best.n) best = e;
    if (best?.f) fmt[role] = best.f;
  }
  return fmt;
}

/** 본문 값은 쓰지 않고 원본 안의 헤더·보고서 필터 단추만 준비합니다. */
export function pivotImportedButtons(sheet, def, plan) {
  const buttons = [], { grid, pm, d, top, left } = plan, position = { ...pm, top, left };
  grid.forEach((row, r) => row.forEach((cd, c) => {
    if (r >= pm.pageRows + pm.headerRows || !inside(def.area, top + r, left + c) || !(plan.cellAt ? plan.cellAt(top+r,left+c) : sheet.cells.getRC(top+r,left+c))) return;
    buttons.push(...pivotCellButtons(cd, top + r, left + c, position, d, { includeToggle: false }));
  }));
  const saved=def.importedPresentation;
  if(d.showExpand!==false&&importedPivotPresentationCurrent(def,saved)){
    const dr=top-saved.top,dc=left-saved.left;
    for(const point of saved.toggles??[]){
      const r=point.r+dr,c=point.c+dc,cell=plan.cellAt?plan.cellAt(r,c):sheet.cells.getRC(r,c);
      if(inside(def.area,r,c)&&cell?.raw===point.raw)
        buttons.push({r,c,kind:'toggle',field:point.field,item:point.item,collapsed:point.collapsed});
    }
  }
  return buttons;
}

/** 원본 소유 영역에서 새 본문/필터가 쓰지 않는 사각형만 블록 contents clear에 넘깁니다. */
export function pivotImportedClearAreas(def, marker, plan) {
  const top=def.top??0,left=def.left??0,a=def.area;
  const body={r1:Math.max(a.r1,top+marker.bodyRow),c1:Math.max(a.c1,left+marker.bodyColumnStart),r2:a.r2,c2:Math.min(a.c2,left+marker.bodyColumnEnd)};
  let areas=[body,...marker.pageFields.map(p=>({r1:top+p.r,c1:left+p.c,r2:top+p.r,c2:left+p.c+1}))].filter(validArea);
  const next=[];
  if(plan.bodyColsN>0&&plan.grid.length>plan.pm.pageRows)next.push({r1:plan.top+plan.pm.pageRows,c1:plan.left,r2:plan.area.r2,c2:plan.left+plan.bodyColsN-1});
  for(const p of plan.pm.pageFields??[])next.push({r1:plan.top+p.r,c1:plan.left+p.c,r2:plan.top+p.r,c2:plan.left+p.c+1});
  for(const n of next)areas=areas.flatMap(o=>{
    const r1=Math.max(o.r1,n.r1),r2=Math.min(o.r2,n.r2),c1=Math.max(o.c1,n.c1),c2=Math.min(o.c2,n.c2);
    if(r1>r2||c1>c2)return[o];
    return[{r1:o.r1,c1:o.c1,r2:r1-1,c2:o.c2},{r1:r2+1,c1:o.c1,r2:o.r2,c2:o.c2},
      {r1,c1:o.c1,r2,c2:c1-1},{r1,c1:c2+1,r2,c2:o.c2}].filter(validArea);
  });
  return areas;
}
