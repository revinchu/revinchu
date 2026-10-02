import { el, openDialog } from './ui.js';
import { setSafeHtml } from './safe-html.js';
import { SMARTART_CATEGORIES, SMARTART_LAYOUTS, SMARTART_PALETTES, SMARTART_STYLES, SMARTART_MAX_NODES, normalizeSmartArt, editSmartArt, smartArtPicture } from './smartart.js';
import { smartArtSvg } from './smartart-render.js';

/** 편집 초안만 변경한다. 확인 시 호출자가 보호 검사와 한 번의 transaction을 수행한다. */
export function openSmartArtEditor({ shape, onCommit, onConvert }) {
  let draft = { ...structuredClone(shape), smartArt: normalizeSmartArt(shape.smartArt) }, selected = draft.smartArt.nodes[0].id;
  let category = SMARTART_LAYOUTS.find(x => x.id === draft.smartArt.layout)?.category ?? '목록', closed = false, pictureRead = 0, pictureBusy = false;
  const preview = el('div', { class: 'sa-preview', 'aria-label': 'SmartArt 미리 보기' });
  const gallery = el('div', { class: 'sa-gallery', role: 'group', 'aria-label': 'SmartArt 레이아웃' });
  const list = el('div', { class: 'sa-text-list', 'aria-label': '텍스트 창' });
  const notice = el('div', { class: 'muted sa-note', role: 'status' });
  const resultCount = el('span', { class: 'sa-result-count', role: 'status' });
  const layoutName = el('strong', { class: 'sa-selected-name' }), description = el('p', { class: 'sa-description' }), count = el('span', { class: 'sa-node-count' });
  const search = el('input', { type: 'search', placeholder: '이름이나 용도로 레이아웃 검색', 'aria-label': 'SmartArt 레이아웃 검색' });
  const categories = el('div', { class: 'sa-categories', role: 'tablist', 'aria-label': 'SmartArt 범주', 'aria-orientation': 'vertical' });
  const commands = [['add','항목 추가'],['before','앞에 추가'],['child','하위 항목 추가'],['delete','항목 삭제'],['promote','수준 올리기'],['demote','수준 내리기'],['up','위로 이동'],['down','아래로 이동']];
  const buttons = new Map(), paletteButtons = [], styleButtons = [];
  const focusNode = (selectAll = false) => { const input = list.querySelector(`[data-node-index="${draft.smartArt.nodes.findIndex(n => n.id === selected)}"] textarea`); input?.focus(); if (selectAll) input?.select(); };
  const markSelected = () => {
    const at = draft.smartArt.nodes.findIndex(n => n.id === selected), node = draft.smartArt.nodes[at];
    list.querySelectorAll('.sa-text-row').forEach(row => row.classList.toggle('on', row.dataset.nodeIndex === String(at)));
    preview.querySelectorAll('[data-smartart-node]').forEach(part => part.classList.toggle('sa-node-selected', part.dataset.smartartNode === selected));
    count.textContent = `항목 ${at + 1} / ${draft.smartArt.nodes.length} · 수준 ${(node?.level ?? 0) + 1}`;
    for (const [action, button] of buttons) {
      if (['add','before','child'].includes(action)) button.disabled = draft.smartArt.nodes.length >= SMARTART_MAX_NODES || action === 'child' && node.level >= 5;
      else if (action === 'delete') button.disabled = draft.smartArt.nodes.length <= 1;
      else { const after = editSmartArt(draft.smartArt, selected, action); button.disabled = after.nodes.every((n,i) => n.id === draft.smartArt.nodes[i].id && n.level === draft.smartArt.nodes[i].level); }
    }
    removePicture.disabled = !node?.picture;
  };
  const redrawPreview = () => {
    setSafeHtml(preview, smartArtSvg({ ...draft, w: 560, h: 300 }));
    for (const part of preview.querySelectorAll('[data-smartart-node]')) part.addEventListener('click', () => { selected = part.dataset.smartartNode; markSelected(); focusNode(); });
    const item = SMARTART_LAYOUTS.find(x => x.id === draft.smartArt.layout);
    layoutName.textContent = item.name; description.textContent = item.description;
    markSelected();
  };
  const apply = (action, focus = false) => {
    try {
      const before = draft.smartArt.nodes, at = before.findIndex(n => n.id === selected);
      draft.smartArt = editSmartArt(draft.smartArt, selected, action);
      if (['add','before','child'].includes(action)) selected = draft.smartArt.nodes.find(n => !before.some(o => o.id === n.id))?.id ?? selected;
      if (!draft.smartArt.nodes.some(n => n.id === selected)) selected = draft.smartArt.nodes[Math.min(at, draft.smartArt.nodes.length - 1)].id;
      drawText(); redrawPreview(); notice.textContent = ''; if (focus) focusNode(true);
    } catch (e) { notice.textContent = e.message; }
  };
  const drawText = () => {
    const scroll = list.scrollTop;
    list.replaceChildren(...draft.smartArt.nodes.map((node, i) => {
      const input = el('textarea', { rows: 2, maxlength: 500, 'aria-label': `항목 ${i + 1} 텍스트` }, node.text);
      input.addEventListener('focus', () => { selected = node.id; markSelected(); });
      input.addEventListener('input', () => { draft.smartArt = editSmartArt(draft.smartArt, node.id, 'text', input.value); redrawPreview(); });
      input.addEventListener('keydown', event => {
        if (event.isComposing) return;
        let action;
        if (event.key === 'Tab' && !event.ctrlKey && !event.altKey && !event.metaKey) action = event.shiftKey ? 'promote' : 'demote';
        else if (event.key === 'Enter' && !event.shiftKey && !event.ctrlKey && !event.altKey && !event.metaKey) action = 'add';
        else if (event.altKey && event.shiftKey && event.key === 'ArrowUp') action = 'up';
        else if (event.altKey && event.shiftKey && event.key === 'ArrowDown') action = 'down';
        if (action) { event.preventDefault(); event.stopPropagation(); selected = node.id; apply(action, true); }
      });
      return el('div', { class: 'sa-text-row', 'data-node-index': i, style: { marginLeft: `${node.level * 14}px` } },
        el('span', { class: 'sa-node-number', title: `수준 ${node.level + 1}` }, `${i + 1}.`), input);
    }));
    list.scrollTop = scroll; markSelected();
  };
  const sample = (layout, style = draft.smartArt.style) => ({ ...draft, id: `thumb-${layout.id}-${style}`, w: 140, h: 82,
    smartArt: { ...draft.smartArt, style, layout: layout.id, nodes: Array.from({ length: layout.category === '계층' ? 5 : 4 }, (_,i) => ({ id:`sample${i}`, text:'', level:layout.category === '계층' ? (i === 0 ? 0 : i === 2 ? 2 : 1) : 0 })) } });
  const pickLayout = layout => {
    draft.smartArt.layout = layout.id; drawGallery(); drawAppearance(); redrawPreview();
    gallery.querySelector(`[data-layout="${layout.id}"]`)?.focus({ preventScroll: true });
  };
  const drawGallery = () => {
    const query = search.value.trim().toLocaleLowerCase(), layouts = SMARTART_LAYOUTS.filter(x => (category === '전체' || x.category === category) && (!query || `${x.name} ${x.category} ${x.description}`.toLocaleLowerCase().includes(query)));
    resultCount.textContent = `${layouts.length}개 레이아웃`;
    gallery.replaceChildren(...layouts.map(layout => el('button', { type: 'button', class: `sa-layout${layout.id === draft.smartArt.layout ? ' on' : ''}`, 'data-layout': layout.id,
      'data-access-key': 'none', 'aria-label': layout.name, 'aria-pressed': String(layout.id === draft.smartArt.layout), title: layout.description, tabindex: layout.id === draft.smartArt.layout || !layouts.some(x => x.id === draft.smartArt.layout) && layout === layouts[0] ? 0 : -1,
      onclick: () => pickLayout(layout), onkeydown: event => {
        if (!['ArrowLeft','ArrowRight','ArrowUp','ArrowDown','Home','End'].includes(event.key)) return;
        event.preventDefault(); const items = [...gallery.querySelectorAll('.sa-layout')], index = items.indexOf(event.currentTarget);
        let columns = items.filter(item => Math.abs(item.getBoundingClientRect().top - items[0].getBoundingClientRect().top) < 2).length || 1;
        const next = event.key === 'Home' ? 0 : event.key === 'End' ? items.length - 1 : Math.max(0,Math.min(items.length-1,index+(event.key === 'ArrowLeft'?-1:event.key === 'ArrowRight'?1:event.key === 'ArrowUp'?-columns:columns)));
        items.forEach((item,i) => item.tabIndex = i === next ? 0 : -1); items[next]?.focus();
      }
    }, el('span', { class:'sa-thumbnail', html: smartArtSvg(sample(layout)) }), el('span', { class:'sa-layout-name' }, layout.name))));
    if (!layouts.length) gallery.append(el('p', { class: 'sa-no-results', role:'status' }, '일치하는 레이아웃이 없습니다. 검색어를 줄이거나 전체 범주를 선택하세요.'));
    for (const button of categories.children) { const on = button.dataset.category === category; button.setAttribute('aria-selected',String(on)); button.tabIndex=on?0:-1; }
  };
  const categoryNames = ['전체', ...SMARTART_CATEGORIES];
  categories.append(...categoryNames.map((name,i) => el('button', { type:'button',role:'tab','data-category':name,'data-access-key':'none','aria-label':name,
    onclick:() => { category=name;drawGallery(); },onkeydown:event => {
      if(!['ArrowUp','ArrowDown','Home','End'].includes(event.key))return;event.preventDefault();
      const next=event.key==='Home'?0:event.key==='End'?categoryNames.length-1:Math.max(0,Math.min(categoryNames.length-1,i+(event.key==='ArrowDown'?1:-1)));
      category=categoryNames[next];drawGallery();categories.children[next].focus();
    }
  },name)));
  search.addEventListener('input',()=>{category='전체';drawGallery();});
  search.addEventListener('keydown',event=>{if(event.key==='ArrowDown'){event.preventDefault();gallery.querySelector('.sa-layout')?.focus();}});
  const palettes = el('div',{class:'sa-palette-list',role:'group','aria-label':'SmartArt 색 변경'});
  SMARTART_PALETTES.forEach((colors,i)=>{
    const button=el('button',{type:'button',class:'sa-palette','aria-label':`색: ${colors[0]}`,'data-access-key':'none',onclick:()=>{draft.smartArt.palette=colors.slice(1);drawAppearance();drawGallery();redrawPreview();}},
      el('span',{class:'sa-palette-colors','aria-hidden':'true'},colors.slice(1).map(color=>el('i',{style:{background:color}}))),el('span',{},colors[0]));paletteButtons.push(button);palettes.append(button);
  });
  const styles=el('div',{class:'sa-style-list',role:'group','aria-label':'SmartArt 스타일'});
  SMARTART_STYLES.forEach(([key,name])=>{const button=el('button',{type:'button',class:'sa-style','aria-label':`스타일: ${name}`,'data-access-key':'none',onclick:()=>{draft.smartArt.style=key;drawAppearance();drawGallery();redrawPreview();}},el('span',{class:'sa-style-sample'}),el('span',{},name));styleButtons.push(button);styles.append(button);});
  const drawAppearance=()=>{
    paletteButtons.forEach((button,i)=>button.setAttribute('aria-pressed',String(JSON.stringify(draft.smartArt.palette)===JSON.stringify(SMARTART_PALETTES[i].slice(1)))));
    styleButtons.forEach((button,i)=>{button.setAttribute('aria-pressed',String(draft.smartArt.style===SMARTART_STYLES[i][0]));setSafeHtml(button.firstElementChild,smartArtSvg(sample(SMARTART_LAYOUTS.find(x=>x.id===draft.smartArt.layout),SMARTART_STYLES[i][0])));});
  };
  const file=el('input',{type:'file',accept:'image/png,image/jpeg','aria-label':'선택 항목 그림 파일',style:{display:'none'}});
  const choosePicture=el('button',{type:'button',class:'btn',onclick:()=>file.click()},'그림 선택…');
  const removePicture=el('button',{type:'button',class:'btn',onclick:()=>{pictureRead++;pictureBusy=false;const node=draft.smartArt.nodes.find(n=>n.id===selected);if(node){delete node.picture;redrawPreview();}file.value='';}},'그림 제거');
  file.addEventListener('change',async()=>{
    const f=file.files?.[0],target=selected,version=++pictureRead;pictureBusy=false;if(!f)return;
    if(f.size>2*1024*1024||!['image/png','image/jpeg'].includes(f.type)){notice.textContent='2MB 이하 PNG/JPEG 그림을 선택하세요.';file.value='';return;}
    pictureBusy=true;notice.textContent='그림을 읽고 있습니다…';
    try{
      const src=await new Promise((resolve,reject)=>{const reader=new FileReader();reader.onload=()=>resolve(reader.result);reader.onerror=()=>reject(new Error('그림 파일을 읽지 못했습니다.'));reader.readAsDataURL(f);});
      if(closed||version!==pictureRead)return;if(!smartArtPicture(src))throw new Error('PNG/JPEG 그림을 확인하세요.');
      const model=structuredClone(draft.smartArt),node=model.nodes.find(n=>n.id===target);if(!node)return;node.picture=src;draft.smartArt=normalizeSmartArt(model);
      if(!draft.smartArt.layout.startsWith('picture')){draft.smartArt.layout='pictureCards';category='그림';search.value='';drawGallery();}drawAppearance();redrawPreview();notice.textContent='';
    }catch(e){if(!closed&&version===pictureRead)notice.textContent=e.message;}finally{if(version===pictureRead){pictureBusy=false;file.value='';}}
  });
  const toolbar=el('div',{class:'sa-toolbar',role:'toolbar','aria-label':'SmartArt 항목 편집'},commands.map(([action,title])=>{const button=el('button',{type:'button',class:'btn',onclick:()=>apply(action,true)},title);buttons.set(action,button);return button;}));
  const appearance=el('details',{class:'sa-appearance'},el('summary',{},'색 및 스타일'),el('div',{class:'sa-appearance-body'},palettes,styles));
  const body=el('div',{class:'sa-editor sa-editor-expanded'},
    el('div',{class:'sa-search-row'},search,resultCount),el('div',{class:'sa-library'},categories,gallery),
    el('div',{class:'sa-layout-info'},layoutName,description),appearance,
    el('div',{class:'sa-workarea'},el('section',{class:'sa-text-editor'},el('div',{class:'sa-text-heading'},el('strong',{},'텍스트 창'),count),toolbar,list),
      el('section',{class:'sa-preview-area'},preview,el('div',{class:'sa-picture'},el('span',{},'선택 항목 그림'),choosePicture,removePicture,file))),notice,
    el('p',{class:'muted sa-note'},'Enter: 항목 추가 · Shift+Enter: 줄 바꿈 · Tab/Shift+Tab: 수준 변경 · Alt+Shift+↑/↓: 순서 변경. 최대 60개이며, 항목이 많으면 글자가 작아집니다.'),
    el('p',{class:'muted sa-note'},'Excel에는 편집 가능한 일반 도형 그룹으로 저장합니다. Excel의 기본 SmartArt와 동일한 개체 형식은 아닙니다.'));
  drawText();drawGallery();drawAppearance();redrawPreview();
  const confirm=convert=>{if(pictureBusy){notice.textContent='그림을 읽은 뒤 다시 확인하세요.';return false;}return (convert?onConvert:onCommit)({...draft,smartArt:normalizeSmartArt(draft.smartArt)});};
  return openDialog({title:'SmartArt 그래픽',body,width:1080,onClose:()=>{closed=true;pictureRead++;},buttons:[
    ...(onConvert?[{label:'일반 도형으로 변환',action:()=>confirm(true)}]:[]),
    {label:'확인',primary:true,action:()=>confirm(false)},{label:'취소'}]});
}
