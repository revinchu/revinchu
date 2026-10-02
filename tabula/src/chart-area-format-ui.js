import { el } from './ui.js';
import { normalizeChartAreaFormat } from './chart-area-format.js';
import { SHAPE_DASH_OPTIONS, SHAPE_PATTERN_OPTIONS, addShapeGradientStop, shapeGradientStopPatch } from './shape-format.js';

let areaPanelId = 0;
export function createChartAreaFormatPanel({ getFormat, getIdentity = () => null, onChange, kind = 'chart', initialTab = 'fillLine', onTab = () => {} }) {
  const body = el('div', { class: 'chart-area-format' }), uid = 'chart-area-' + (++areaPanelId);
  let tab = initialTab, stopIndex = 0, lineStopIndex = 0;
  const row = (name, input) => { input.setAttribute('aria-label', name); return el('label', { class: 'cfp-row' }, el('span', {}, name), input); };
  const section = (name, ...nodes) => el('details', { class: 'cfp-sec', open: true }, el('summary', {}, name), nodes);
  const button = (text, action) => el('button', { type: 'button', class: 'btn small', onclick: action }, text);
  const color = (value, action) => { const i = el('input', { type: 'color', value }); i.addEventListener('change', () => action(i.value)); return i; };
  const num = (value, action, min = 0, max = 100, step = 1) => { const i = el('input', { type: 'number', value: Math.round(value * 100) / 100, min, max, step }); i.addEventListener('change', () => { if (i.value !== '' && Number.isFinite(i.valueAsNumber)) action(Math.min(max, Math.max(min, i.valueAsNumber))); }); return i; };
  const choose = (value, options, action) => { const i = el('select', {}, options.map(([v, text]) => el('option', { value: v, selected: value === v }, text))); i.addEventListener('change', () => action(i.value)); return i; };
  const check = (value, action) => { const i = el('input', { type: 'checkbox', checked: value }); i.addEventListener('change', () => action(i.checked)); return i; };
  const current = () => normalizeChartAreaFormat(getFormat(), kind);
  const up = patch => { onChange({ ...current(), ...patch }); draw(); };
  const modes = (name, value, options, action) => el('div', { class: 'shape-mode-options', role: 'radiogroup', 'aria-label': name }, options.map(([v, text]) => {
    const input = el('input', { type: 'radio', name: uid + name, value: v, checked: v === value, 'aria-label': text }); input.addEventListener('change', () => { if (input.checked) action(v); }); return el('label', {}, input, el('span', {}, text));
  }));
  const texture = (name) => {
    const canvas=document.createElement('canvas');canvas.width=canvas.height=128;const ctx=canvas.getContext('2d');ctx.scale(8,8);
    ctx.fillStyle=name==='linen'?'#f5f0e6':name==='paper'?'#f5e4be':'#dceaf7';ctx.fillRect(0,0,16,16);ctx.strokeStyle=name==='linen'?'#d3c8b7':name==='paper'?'#d9bf90':'#a8c4df';ctx.lineWidth=name==='woven'?2:.5;ctx.beginPath();
    const line=(x,y,x2,y2)=>{ctx.moveTo(x,y);ctx.lineTo(x2,y2);};
    if(name==='linen')for(let i=2;i<16;i+=4){line(0,i,16,i);line(i,0,i,16);}else if(name==='paper'){line(0,3,16,2);line(0,10,16,12);line(3,0,5,16);line(12,0,10,16);}else{line(0,0,16,16);line(-8,0,8,16);line(8,0,24,16);line(0,16,16,0);}ctx.stroke();return canvas.toDataURL('image/png');
  };
  const decodePicture = src => new Promise((resolve,reject)=>{const image=new Image();image.onload=()=>{const width=image.naturalWidth,height=image.naturalHeight;if(!width||!height||width*height>16777216){reject(Error('그림 해상도는 1,600만 화소 이하로 선택하세요.'));return;}if(/^data:image\/(svg\+xml|webp)/i.test(src)){const canvas=document.createElement('canvas');canvas.width=width;canvas.height=height;try{canvas.getContext('2d').drawImage(image,0,0);src=canvas.toDataURL('image/png');}catch(e){reject(e);return;}}resolve({src,width,height});};image.onerror=reject;image.src=src;});
  const gradientRows = (f, key, prefix) => {
    const g = f[key], index = Math.min(key === 'grad' ? stopIndex : lineStopIndex, g.stops.length - 1), stop = g.stops[index];
    const change = patch => up({ [key]: { ...current()[key], ...patch } });
    const changeStop = patch => up({ [key]: shapeGradientStopPatch({ grad: current()[key] }, index, patch).grad });
    return [row(prefix + ' 종류', choose(g.type, [['linear', '선형'], ['radial', '방사형']], type => change({type}))), ...(g.type === 'linear' ? [row(prefix + ' 각도(°)', num(g.ang, ang => change({ang}), 0, 360))] : []),
      row(prefix + ' 중지점', choose(String(index), g.stops.map((s, i) => [String(i), `${i + 1} · ${Math.round(s[0] * 100)}%`]), v => { if (key === 'grad') stopIndex = +v; else lineStopIndex = +v; draw(); })),
      el('div', { class: 'chart-area-buttons' }, button('중지점 추가', () => { const next = addShapeGradientStop({ grad: current()[key] }); if (key === 'grad') stopIndex = next.index; else lineStopIndex = next.index; up({[key]:next.grad}); }), el('button', { type: 'button', class: 'btn small', disabled: g.stops.length <= 2, onclick: () => { const stops = current()[key].stops.map(s => [...s]); stops.splice(index, 1); up({[key]:{...g,stops}}); } }, '중지점 제거')),
      row(prefix + ' 위치(%)', num(stop[0] * 100, n => changeStop({position:n / 100}))), row(prefix + ' 색', color(stop[1], c => changeStop({color:c}))), row(prefix + ' 투명도(%)', num((1 - stop[2]) * 100, n => changeStop({opacity:1 - n / 100})))];
  };
  const expanded = new Map();
  function draw() {
    for (const d of body.querySelectorAll('details')) expanded.set(d.querySelector('summary')?.textContent, d.open);
    const focused = body.contains(document.activeElement) ? document.activeElement.getAttribute('aria-label') : null;
    const scrollBox = body.closest('.dialog-body'), scroll = scrollBox?.scrollTop;
    const f = current(), sections = [];
    const tabs = el('div', { class: 'format-pane-tabs chart-area-tabs', role: 'tablist', 'aria-label': '영역 서식 범주' }, [['fillLine','채우기 및 선'],['effects','효과'],['threeD','3차원 서식']].map(([value,label]) => el('button', { type: 'button', role: 'tab', class: value === tab ? 'on' : '', 'aria-selected': String(value === tab), onclick: () => { tab = value; onTab(value); draw(); } }, label)));
    tabs.addEventListener('keydown', event => { if (!['ArrowLeft','ArrowRight','Home','End'].includes(event.key)) return; event.preventDefault(); const buttons = [...tabs.children], at = buttons.indexOf(document.activeElement), i = event.key === 'Home' ? 0 : event.key === 'End' ? buttons.length - 1 : (at + (event.key === 'ArrowRight' ? 1 : -1) + buttons.length) % buttons.length; buttons[i].click(); body.querySelector('[role="tab"][aria-selected="true"]')?.focus(); });
    if (tab === 'fillLine') {
      const fillRows = [modes('채우기', f.fillMode, [['none','채우기 없음'],['solid','단색 채우기'],['gradient','그라데이션 채우기'],['picture','그림 또는 질감 채우기'],['pattern','패턴 채우기'],['auto','자동 채우기']], fillMode => up({fillMode}))];
      if (f.fillMode === 'solid') fillRows.push(row('채우기 색', color(f.fill, fill => up({fill}))));
      if (f.fillMode === 'gradient') fillRows.push(...gradientRows(f,'grad','채우기 그라데이션'));
      if (f.fillMode === 'pattern') fillRows.push(row('패턴', choose(f.pattern.preset, SHAPE_PATTERN_OPTIONS, preset => up({pattern:{...f.pattern,preset}}))), row('패턴 전경색', color(f.pattern.fg, fg => up({pattern:{...f.pattern,fg}}))), row('패턴 배경색', color(f.pattern.bg, bg => up({pattern:{...f.pattern,bg}}))));
      if (f.fillMode === 'picture') {
        const file = el('input', {type:'file',accept:'image/png,image/jpeg,image/webp,image/gif,image/svg+xml','aria-label':'영역 채우기 그림 파일'}), status = el('p',{role:'status',class:'muted'});
        file.addEventListener('change', async () => { const chosen = file.files?.[0], identity = getIdentity(); if (!chosen) return; if (chosen.size > 8 * 1024 * 1024 || !/^image\/(png|jpeg|webp|gif|svg\+xml)$/.test(chosen.type)) { status.textContent = '8MB 이하 PNG·JPEG·WebP·GIF·SVG 그림을 선택하세요.'; return; } try { const src = await new Promise((resolve,reject) => { const r = new FileReader(); r.onload = () => resolve(r.result); r.onerror = reject; r.readAsDataURL(chosen); }); const image=await decodePicture(src);if (!body.isConnected || getIdentity() !== identity) return; up({picture:{...current().picture,...image}}); } catch(error) { status.textContent = error?.message||'그림을 읽지 못했습니다.'; } });
        fillRows.push(row('그림 삽입',file),status,row('질감',choose('',[['','질감 선택…'],['linen','리넨'],['paper','종이'],['woven','직물']], name => { if (name) up({picture:{...f.picture,src:texture(name),width:128,height:128,mode:'tile'}}); })),row('그림 배치',choose(f.picture.mode,[['stretch','영역에 맞게 늘이기'],['tile','질감으로 바둑판식 채우기']],mode=>up({picture:{...f.picture,mode}}))));
        if (f.picture.mode === 'tile') fillRows.push(row('타일 크기(%)',num(f.picture.scale*100,n=>up({picture:{...f.picture,scale:n/100}}),5,1000)));
        if (f.picture.src) fillRows.push(el('img',{src:f.picture.src,alt:'영역 채우기 그림 미리 보기',class:'chart-area-image-preview'}));
      }
      if (!['none','auto'].includes(f.fillMode)) fillRows.push(row('채우기 투명도(%)',num((1-f.fillOpacity)*100,n=>up({fillOpacity:1-n/100}))));
      sections.push(section('채우기',...fillRows));
      const lineRows = [modes('테두리',f.lineMode,[['none','선 없음'],['solid','실선'],['gradient','그라데이션 선'],['auto','자동 선']],lineMode=>up({lineMode}))];
      if (['solid','gradient'].includes(f.lineMode)) {
        if (f.lineMode==='solid') lineRows.push(row('테두리 색',color(f.stroke,stroke=>up({stroke})))); else lineRows.push(...gradientRows(f,'strokeGrad','선 그라데이션'));
        lineRows.push(row('선 투명도(%)',num((1-f.strokeOpacity)*100,n=>up({strokeOpacity:1-n/100}))),row('선 너비(pt)',num(f.strokeWidth*.75,n=>up({strokeWidth:n/.75}),0,30,.25)),row('겹선 종류',choose(f.compound,[['sng','단선'],['dbl','겹선'],['thickThin','굵고 가는 선'],['thinThick','가늘고 굵은 선'],['tri','삼중선']],compound=>up({compound}))),row('대시 종류',choose(f.dash,SHAPE_DASH_OPTIONS,dash=>up({dash}))),row('끝 모양',choose(f.lineCap,[['flat','평면'],['rnd','원형'],['sq','사각형']],lineCap=>up({lineCap}))),row('연결 종류',choose(f.lineJoin,[['round','원형'],['bevel','빗면'],['miter','각진']],lineJoin=>up({lineJoin}))));
      }
      sections.push(section('테두리',...lineRows));
    } else if (tab === 'effects') {
      const shadow = f.shadow ?? {dx:3,dy:3,blur:4,color:'#000000',opacity:.3,scale:1}, glow=f.glow??{size:6,color:'#ffc000',opacity:.5};
      const angle=(Math.atan2(shadow.dy,shadow.dx)*180/Math.PI+360)%360,distance=Math.hypot(shadow.dx,shadow.dy);
      const polar=(degrees,dist)=>up({shadow:{...shadow,dx:Math.cos(degrees*Math.PI/180)*dist,dy:Math.sin(degrees*Math.PI/180)*dist}});
      const shadowPresets={bottomRight:{dx:4,dy:4,blur:6,scale:1},bottomLeft:{dx:-4,dy:4,blur:6,scale:1},bottom:{dx:0,dy:5,blur:8,scale:1},center:{dx:0,dy:0,blur:8,scale:1.04}};
      sections.push(section('그림자',row('그림자 미리 설정',choose(f.shadow?'custom':'none',[['none','그림자 없음'],['custom','사용자 지정'],['bottomRight','바깥쪽 · 오른쪽 아래'],['bottomLeft','바깥쪽 · 왼쪽 아래'],['bottom','바깥쪽 · 아래'],['center','바깥쪽 · 가운데']],v=>{if(v==='none')up({shadow:null});else if(shadowPresets[v])up({shadow:{...shadow,...shadowPresets[v]}});})),row('그림자 사용',check(!!f.shadow,v=>up({shadow:v?shadow:null}))),...(f.shadow?[row('그림자 색',color(shadow.color,color=>up({shadow:{...shadow,color}}))),row('그림자 투명도(%)',num((1-shadow.opacity)*100,n=>up({shadow:{...shadow,opacity:1-n/100}}))),row('그림자 크기(%)',num(shadow.scale*100,n=>up({shadow:{...shadow,scale:n/100}}),5,400)),row('그림자 흐리게(pt)',num(shadow.blur*.75,n=>up({shadow:{...shadow,blur:n/.75}}),0,75,.25)),row('그림자 각도(°)',num(angle,n=>polar(n,distance),0,360)),row('그림자 간격(pt)',num(distance*.75,n=>polar(angle,n/.75),0,75,.25))]:[])));
      sections.push(section('네온',row('네온 미리 설정',choose(f.glow?'custom':'none',[['none','네온 없음'],['custom','사용자 지정'],['blue','파랑 · 5pt'],['orange','주황 · 8pt'],['green','초록 · 11pt']],v=>{const presets={blue:{color:'#4472c4',size:20/3},orange:{color:'#ed7d31',size:32/3},green:{color:'#70ad47',size:44/3}};if(v==='none')up({glow:null});else if(presets[v])up({glow:{...glow,...presets[v]}});})),row('네온 사용',check(!!f.glow,v=>up({glow:v?glow:null}))),...(f.glow?[row('네온 색',color(glow.color,color=>up({glow:{...glow,color}}))),row('네온 크기(pt)',num(glow.size*.75,n=>up({glow:{...glow,size:n/.75}}),0,75,.25)),row('네온 투명도(%)',num((1-glow.opacity)*100,n=>up({glow:{...glow,opacity:1-n/100}})))]:[])));
      sections.push(section('부드러운 가장자리',row('가장자리 미리 설정',choose(f.soft?'custom':'0',[['0','가장자리 없음'],['custom','사용자 지정'],['1','1pt'],['2.5','2.5pt'],['5','5pt'],['10','10pt']],v=>{if(v!=='custom')up({soft:+v/.75});})),row('가장자리 크기(pt)',num(f.soft*.75,n=>up({soft:n/.75}),0,37.5,.25))));
    } else {
      const d=f.threeD, set=patch=>up({threeD:{...current().threeD,...patch}});
      for(const [key,title] of [['bevelTop','위쪽 입체'],['bevelBottom','아래쪽 입체']]) {const b=d[key];sections.push(section(title,row(title+' 종류',choose(b.type,[['none','없음'],['angle','각진'],['circle','둥근'],['convex','볼록']],type=>set({[key]:{...b,type}}))),...(b.type!=='none'?[row(title+' 너비(pt)',num(b.w*.75,n=>set({[key]:{...b,w:n/.75}}),0,54,.25)),row(title+' 높이(pt)',num(b.h*.75,n=>set({[key]:{...b,h:n/.75}}),0,54,.25))]:[])));}
      sections.push(section('깊이 및 외형선',row('깊이(pt)',num(d.depth*.75,n=>set({depth:n/.75}),0,75,.25)),row('외형선 크기(pt)',num(d.contourWidth*.75,n=>set({contourWidth:n/.75}),0,15,.25)),row('외형선 색',color(d.contourColor,contourColor=>set({contourColor})))));
      sections.push(section('재질 및 조명',row('재질',choose(d.material,[['plastic','플라스틱'],['matte','무광'],['metal','금속']],material=>set({material}))),row('조명',choose(d.lightRig,[['threePt','세 점'],['balanced','균형'],['soft','부드럽게']],lightRig=>set({lightRig}))),row('조명 각도(°)',num(d.lightAngle,lightAngle=>set({lightAngle}),0,360))));
      sections.push(el('p',{class:'muted'},'영역의 입체 테두리와 깊이는 2차원 음영으로 표시합니다. Excel의 3차원 카메라·재질과 광택은 다르게 보일 수 있습니다.'));
    }
    body.replaceChildren(tabs,...sections,button('이 영역 서식 초기화',()=>{onChange(null);draw();}));
    for(const d of body.querySelectorAll('details')) if(expanded.has(d.querySelector('summary')?.textContent)) d.open=expanded.get(d.querySelector('summary')?.textContent);
    if(focused) [...body.querySelectorAll('[aria-label]')].find(n=>n.getAttribute('aria-label')===focused)?.focus({preventScroll:true});
    if(scrollBox&&scroll!==undefined)scrollBox.scrollTop=scroll;
  }
  draw();return {body,refresh:draw};
}
