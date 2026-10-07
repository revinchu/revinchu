// 대량 슬라이서도 보이는 항목만 DOM으로 만든다. 전체 항목/선택값은 모델에 그대로 남긴다.
const mounted = new WeakMap();
// startItem은 sharedItems의 x가 아니라 표시 목록의 0 기반 항목 번호입니다.
export function slicerStartItem(value, count = Infinity) {
  const n = typeof value === 'number' ? value : typeof value === 'string' && /^\+?\d+$/.test(value.trim()) ? Number(value) : 0;
  const index = Number.isInteger(n) && n >= 0 && n <= 4294967295 ? n : 0;
  return Number.isFinite(count) ? Math.min(index, Math.max(0, Math.trunc(count) - 1)) : index;
}
export function slicerStartScrollTop(count, columns, height, gap, viewport, startItem) {
  const win = slicerWindow(count, columns, height, gap, viewport, 0);
  const logical = Math.min(win.logicalMax, Math.floor(slicerStartItem(startItem, count) / win.columns) * win.pitch);
  return win.logicalMax ? logical * win.max / win.logicalMax : 0;
}
export function slicerWindow(count, columns, height, gap, viewport, top) {
  columns=Math.max(1,Math.trunc(columns)||1);height=Math.max(1,Number(height)||24);gap=Math.max(0,Number(gap)||0);
  const rows=Math.ceil(count/columns),pitch=height+gap,total=Math.max(8,8+rows*pitch-gap),physical=Math.min(8000000,total);
  const max=Math.max(0,physical-viewport),logicalMax=Math.max(0,total-viewport);
  const scroll=Math.max(0,Math.min(max,top)),logical=max?scroll*logicalMax/max:0;
  const start=Math.max(0,Math.floor(Math.max(0,logical-4)/pitch)-3),end=Math.min(rows,Math.ceil((logical+viewport)/pitch)+3);
  return {first:start*columns,last:Math.min(count,end*columns),top:4+start*pitch-logical+scroll,total,physical,pitch,columns,logical,max,logicalMax};
}
export function mountSlicerWindow(list, model) {
  const previous=mounted.get(list);
  if(previous?.model===model)return;
  previous?.destroy();
  const savedTop=previous ? list.scrollTop : slicerStartScrollTop(model.items.length,model.columns??1,model.height??24,model.gap??3,list.clientHeight||model.viewport||240,model.startItem),savedLeft=list.scrollLeft;
  const {items,columns=1,height=24,gap=3,width}=model,content=document.createElement('div'),grid=document.createElement('div');
  list.style.display='block';list.style.position='relative';list.style.padding='0';list.tabIndex=0;
  content.style.position='relative';grid.style.cssText='position:absolute;left:5px;right:5px;display:grid;align-content:start;';
  grid.style.gridTemplateColumns=`repeat(${Math.max(1,columns)}, ${width?width+'px':'minmax(0, 1fr)'})`;grid.style.gap=gap+'px';
  if(width)content.style.minWidth=(columns*width+Math.max(0,columns-1)*gap+10)+'px';
  content.style.height=slicerWindow(items.length,columns,height,gap,list.clientHeight||model.viewport||240,savedTop).physical+'px';
  content.append(grid);list.replaceChildren(content);list.scrollTop=savedTop;list.scrollLeft=savedLeft;let first=-1,last=-1,frame=0;
  const viewport=()=>list.clientHeight||model.viewport||240;
  const render=()=>{
    frame=0;const win=slicerWindow(items.length,columns,height,gap,viewport(),list.scrollTop);
    content.style.height=win.physical+'px';grid.style.top=win.top+'px';
    if(first===win.first&&last===win.last)return;
    const active=document.activeElement,focusIndex=active?.closest?.('.sl-items')===list?active.dataset.slIndex:null;
    first=win.first;last=win.last;const fragment=document.createDocumentFragment();
    for(let i=first;i<last;i++){
      const item=items[i],button=document.createElement('button');button.type='button';
      button.className='sl-item'+(item.selected?' on':'')+(item.hasData?'':' nodata');
      button.dataset.k=item.key;button.dataset.slIndex=String(i);button.title=item.text;button.textContent=item.text;
      button.style.height=height+'px';fragment.append(button);
    }
    grid.replaceChildren(fragment);model.onRender?.(grid);
    if(focusIndex!==null&&focusIndex!==undefined)grid.querySelector(`[data-sl-index="${focusIndex}"]`)?.focus({preventScroll:true});
  };
  const scroll=()=>{if(!frame)frame=requestAnimationFrame(render);};
  const key=event=>{
    const active=event.target.closest?.('.sl-item'),current=active?Number(active.dataset.slIndex):first;
    const page=Math.max(1,Math.floor(viewport()/(height+gap)))*columns;
    const offsets={ArrowDown:columns,ArrowUp:-columns,ArrowRight:1,ArrowLeft:-1,PageDown:page,PageUp:-page};
    let next;if(event.key==='Home')next=0;else if(event.key==='End')next=items.length-1;else if(Object.hasOwn(offsets,event.key))next=current+offsets[event.key];else return;
    if(!items.length)return;event.preventDefault();event.stopPropagation();next=Math.max(0,Math.min(items.length-1,next));
    const win=slicerWindow(items.length,columns,height,gap,viewport(),list.scrollTop),y=4+Math.floor(next/columns)*win.pitch;
    let logical=win.logical;if(y<logical)logical=y;else if(y+height>logical+viewport())logical=y+height-viewport();
    list.scrollTop=win.logicalMax?Math.max(0,Math.min(win.logicalMax,logical))*win.max/win.logicalMax:0;
    render();grid.querySelector(`[data-sl-index="${next}"]`)?.focus({preventScroll:true});
  };
  list.addEventListener('scroll',scroll,{passive:true});list.addEventListener('keydown',key);render();
  mounted.set(list,{model,destroy(){if(frame)cancelAnimationFrame(frame);list.removeEventListener('scroll',scroll);list.removeEventListener('keydown',key);}});
}
