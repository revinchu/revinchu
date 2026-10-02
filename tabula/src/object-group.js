import { pictureSvg } from './picture-render.js';
import { esc } from './xml.js';

export function makeObjectGroup(objects, id) {
  if (objects.length < 2) throw new Error('그룹화할 개체를 두 개 이상 선택하세요.');
  let x=Infinity,y=Infinity,right=-Infinity,bottom=-Infinity;
  for (const o of objects) { x=Math.min(x,o.x);y=Math.min(y,o.y);right=Math.max(right,o.x+o.w);bottom=Math.max(bottom,o.y+o.h); }
  const w=Math.max(1,right-x),h=Math.max(1,bottom-y);
  return {id,kind:'group',name:'그룹',x,y,w,h,groupSize:{w,h},groupItems:objects.map(o=>({...structuredClone(o),x:o.x-x,y:o.y-y}))};
}
export function ungroupObjects(group) {
  if(group.kind!=='group') return [];
  const sx=group.w/(group.groupSize?.w||group.w),sy=group.h/(group.groupSize?.h||group.h),rad=(group.rot||0)*Math.PI/180,co=Math.cos(rad),sn=Math.sin(rad),fx=group.flip?-1:1,fy=group.flipV?-1:1;
  return (group.groupItems??[]).map((o,i)=>{
    const dx=((o.x+o.w/2)*sx-group.w/2)*fx,dy=((o.y+o.h/2)*sy-group.h/2)*fy;
    const center={x:group.x+group.w/2+co*dx-sn*dy,y:group.y+group.h/2+sn*dx+co*dy};
    const a=(o.rot||0)*Math.PI/180,ax=sx*Math.cos(a),ay=sy*Math.sin(a),bx=-sx*Math.sin(a),by=sy*Math.cos(a);
    // 비균등 크기와 회전으로 전단이 생기면 한 항목 그룹으로 정확한 변환을 보존한다.
    if(Math.abs(ax*bx+ay*by)>1e-7) return {...structuredClone(group),id:o.id||group.id+'-'+i,groupItems:[structuredClone(o)],name:o.name||'그룹 항목'};
    const w=o.w*Math.hypot(ax,ay),h=o.h*Math.hypot(bx,by);
    return {...structuredClone(o),x:center.x-w/2,y:center.y-h/2,w,h,rot:((group.rot||0)+Math.atan2(ay*fx*fy,ax)*180/Math.PI+360)%360,flip:!!o.flip!==!!group.flip||undefined,flipV:!!o.flipV!==!!group.flipV||undefined};
  });
}
function groupText(o) {
  const paragraphs=o.paras??String(o.text??'').split('\n').map(t=>({runs:[{t}]}));
  if(!paragraphs.length) return '';
  const baseSize=(o.size??11)*4/3, pad=o.pad??[4.8,9.6,4.8,9.6], width=Math.max(1,o.w-pad[1]-pad[3]),lines=[];
  for(const p of paragraphs) {
    let runs=[],used=0,height=baseSize*1.2;
    const finish=()=>{lines.push({runs,height,align:p.align??o.align??(o.kind==='textbox'?'left':'center')});runs=[];used=0;height=baseSize*1.2;};
    for(const r of p.runs??[]) {
      let text=''; const fs=(r.sz??o.size??11)*4/3;
      const flush=()=>{if(text)runs.push({...r,t:text});text='';};
      for(const ch of String(r.t??r.text??'')) {
        const cw=/[^\u0000-\u00ff]/.test(ch)?fs:fs*.56;
        if(ch==='\n'||(!o.nowrap&&used&&used+cw>width)){flush();finish();if(ch==='\n')continue;}
        text+=ch;used+=cw;height=Math.max(height,fs*1.2);
      }
      flush();
    }
    finish();
  }
  const total=lines.reduce((n,l)=>n+l.height,0), available=Math.max(1,o.h-pad[0]-pad[2]),scale=o.textFit==='shrink'?Math.min(1,available/total):1;
  const valign=o.valign??(o.kind==='textbox'?'top':'middle');
  let y=valign==='top'?pad[0]:valign==='bottom'?o.h-pad[2]-total*scale:pad[0]+(available-total*scale)/2;
  const body=lines.map(l=>{
    y+=l.height*scale;
    const x=l.align==='left'?pad[3]:l.align==='right'?o.w-pad[1]:(pad[3]+o.w-pad[1])/2;
    return `<text x="${x}" y="${y-l.height*.2*scale}" font-family="${esc(o.font??'맑은 고딕')}" fill="${esc(o.color??'#000000')}" font-size="${baseSize*scale}" text-anchor="${l.align==='left'?'start':l.align==='right'?'end':'middle'}"${o.bold?' font-weight="bold"':''}${o.italic?' font-style="italic"':''}>${l.runs.map(r=>`<tspan${r.b!==undefined?` font-weight="${r.b?'bold':'normal'}"`:''}${r.i!==undefined?` font-style="${r.i?'italic':'normal'}"`:''}${r.sz?` font-size="${r.sz*4/3*scale}"`:''}${r.color?` fill="${esc(r.color)}"`:''}${r.font?` font-family="${esc(r.font)}"`:''}${r.u||r.s?` text-decoration="${r.u?'underline ':''}${r.s?'line-through':''}"`:''}>${esc(r.t)}</tspan>`).join('')}</text>`;
  }).join('');
  return [90,270].includes(Number(o.textRot))?`<g transform="rotate(${o.textRot},${o.w/2},${o.h/2})">${body}</g>`:body;
}

export function groupSvg(group, renderer) {
  const w=group.groupSize?.w||group.w,h=group.groupSize?.h||group.h;
  const flip=`translate(${group.flip?w:0},${group.flipV?h:0}) scale(${group.flip?-1:1},${group.flipV?-1:1})`;
  const parts=(group.groupItems??[]).filter(o=>!o.hidden).map(o=>{
    const transform=`translate(${o.x},${o.y}) rotate(${o.rot||0},${o.w/2},${o.h/2})`;
    let body;
    if(o.kind==='picture') {
      body=pictureSvg(o, String(group.id)+'_'+String(o.id));
    } else body=renderer({...o,id:String(group.id)+'_'+String(o.id)})+(o.kind==='group'?'':groupText(o));
    return `<g transform="${transform}"${o.hyperlink?.target ? ` data-object-link="${esc(o.hyperlink.target)}" data-object-link-mode="${esc(o.hyperlink.targetMode || '')}" style="cursor:pointer"` : ''}>${o.hyperlink?.target ? `<title>${esc(o.hyperlink.tooltip || o.hyperlink.target)}</title>` : ''}${body}</g>`;
  }).join('');
  return `<svg xmlns="http://www.w3.org/2000/svg" width="${group.w}" height="${group.h}" viewBox="0 0 ${w} ${h}" preserveAspectRatio="none" aria-label="그룹" style="overflow:visible"><g transform="${flip}">${parts}</g></svg>`;
}
