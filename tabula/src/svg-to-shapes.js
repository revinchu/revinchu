// Editable, standard DrawingML freeforms from a conservative SVG subset.
// Unsupported paint servers/effects/resources are rejected, never silently lost.
import { parseXml } from './xml.js';
import { normalizedPathBounds } from './shape-path.js';
import { parseSvgPath, parseSvgTransform, transformCommands, multiplyMatrix, flattenCommands } from './geometry-path.js';
import { arrangeRegions, geometryShape, geometryComponents } from './shape-boolean.js';

const fail=detail=>{throw new Error(`이 SVG는 도형으로 변환할 수 없습니다: ${detail}. 원본 그림은 유지됩니다.`);};
const finite=(value,fallback=0)=>{if(value===undefined)return fallback;if(!/^[+-]?(?:\d+\.?\d*|\.\d+)(?:e[+-]?\d+)?(?:px)?$/i.test(String(value)))fail('지원하지 않는 길이 단위');const n=Number(String(value).replace(/px$/i,''));if(!Number.isFinite(n)||Math.abs(n)>1e7)fail('올바르지 않은 좌표');return n;};
const opacity=value=>{const n=value===undefined?1:finite(value);if(n<0||n>1)fail('올바르지 않은 투명도');return n;};
const NAMES={black:'#000000',white:'#ffffff',red:'#ff0000',green:'#008000',blue:'#0000ff',yellow:'#ffff00',gray:'#808080',grey:'#808080',silver:'#c0c0c0',orange:'#ffa500',purple:'#800080',pink:'#ffc0cb',navy:'#000080',teal:'#008080',aqua:'#00ffff',lime:'#00ff00',maroon:'#800000',olive:'#808000',fuchsia:'#ff00ff',rebeccapurple:'#663399'};
function paint(value,current='#000000'){
  const v=String(value??'#000000').trim().toLowerCase();if(v==='none')return {color:null,alpha:1};if(v==='transparent')return {color:'#000000',alpha:0};if(v==='currentcolor')return paint(current);
  if(NAMES[v])return {color:NAMES[v],alpha:1};
  if(/^#[0-9a-f]{3,4}$/i.test(v))return paint('#'+v.slice(1).split('').map(c=>c+c).join(''));
  if(/^#[0-9a-f]{6}(?:[0-9a-f]{2})?$/i.test(v))return {color:v.slice(0,7),alpha:v.length===9?parseInt(v.slice(7),16)/255:1};
  const m=/^rgba?\(([^)]+)\)$/.exec(v);if(m){const a=m[1].split(/\s*[,/]\s*|\s+/).filter(Boolean);if(a.length===3||a.length===4){const ch=a.slice(0,3).map(s=>Number(s.replace('%',''))*(s.endsWith('%')?2.55:1));if(ch.every(n=>Number.isFinite(n)&&n>=0&&n<=255)){const alpha=a.length===4?Number(a[3].replace('%',''))/(a[3].endsWith('%')?100:1):1;if(alpha>=0&&alpha<=1)return {color:'#'+ch.map(n=>Math.round(n).toString(16).padStart(2,'0')).join(''),alpha};}}}
  fail('지원하지 않는 색 또는 그라데이션');
}
const STYLES=new Set(['fill','stroke','stroke-width','fill-rule','fill-opacity','stroke-opacity','opacity','color','stroke-linecap','stroke-linejoin','stroke-miterlimit','stroke-dasharray','stroke-dashoffset','display','visibility']);
function styleFor(node,parent){
  const s={...parent,opacity:'1'};
  for(const [k,v] of Object.entries(node.attrs)){if(STYLES.has(k))s[k]=v;if(/^on/i.test(k)||/href$|filter|clip-path|mask|vector-effect|marker|mix-blend-mode/i.test(k))fail('이벤트·외부 참조·필터·마스크·마커');}
  if(node.attrs.style)for(const declaration of node.attrs.style.split(';')){if(!declaration.trim())continue;const at=declaration.indexOf(':'),k=declaration.slice(0,at).trim(),v=declaration.slice(at+1).trim();if(at<0||!STYLES.has(k)||v.includes('!important'))fail('지원하지 않는 CSS');s[k]=v;}
  for(const k of STYLES)if(s[k]==='inherit')s[k]=parent[k];
  if(String(s.color).toLowerCase()==='currentcolor')s.color=parent.color??'#000000';
  if(s['stroke-dasharray']&&s['stroke-dasharray']!=='none')fail('사용자 지정 점선');
  if(s['stroke-dashoffset']&&finite(s['stroke-dashoffset'])!==0)fail('점선 오프셋');
  if(s['fill-rule']&&!['nonzero','evenodd'].includes(s['fill-rule']))fail('채우기 규칙');
  return s;
}
function primitive(node){
  const a=node.attrs,n=k=>finite(a[k]);
  if(node.name==='path')return a.d?.trim()?parseSvgPath(a.d):null;
  if(node.name==='line')return [['M',n('x1'),n('y1')],['L',n('x2'),n('y2')]];
  if(node.name==='polyline'||node.name==='polygon'){
    const source=(a.points??'').trim();if(!source)return null;const p=source.split(/[\s,]+/).map(v=>finite(v));if(p.length<4||p.length%2)fail('꼭짓점 목록');
    const c=[];for(let i=0;i<p.length;i+=2)c.push([i?'L':'M',p[i],p[i+1]]);if(node.name==='polygon')c.push(['Z']);return c;
  }
  if(node.name==='rect'){
    const x=n('x'),y=n('y'),w=n('width'),h=n('height');if(w<0||h<0)fail('음수 크기');if(!w||!h)return null;
    const rx=Math.min(w/2,finite(a.rx,finite(a.ry))),ry=Math.min(h/2,finite(a.ry,finite(a.rx)));if(rx<0||ry<0)fail('음수 모서리 반지름');
    if(!rx||!ry)return [['M',x,y],['L',x+w,y],['L',x+w,y+h],['L',x,y+h],['Z']];
    return parseSvgPath(`M${x+rx} ${y}H${x+w-rx}A${rx} ${ry} 0 0 1 ${x+w} ${y+ry}V${y+h-ry}A${rx} ${ry} 0 0 1 ${x+w-rx} ${y+h}H${x+rx}A${rx} ${ry} 0 0 1 ${x} ${y+h-ry}V${y+ry}A${rx} ${ry} 0 0 1 ${x+rx} ${y}Z`);
  }
  if(node.name==='circle'||node.name==='ellipse'){
    const x=n('cx'),y=n('cy'),rx=node.name==='circle'?n('r'):n('rx'),ry=node.name==='circle'?n('r'):n('ry');if(rx<0||ry<0)fail('음수 반지름');if(!rx||!ry)return null;
    return parseSvgPath(`M${x-rx} ${y}A${rx} ${ry} 0 1 0 ${x+rx} ${y}A${rx} ${ry} 0 1 0 ${x-rx} ${y}Z`);
  }fail(`지원하지 않는 요소 ${node.name}`);
}
function commandShape(commands,style,id){
  const b=normalizedPathBounds({paths:[{commands}]});if(!b)fail('올바르지 않은 경로');
  const x=b.minX,y=b.minY,w=Math.max(1,b.maxX-x),h=Math.max(1,b.maxY-y);
  const normalized=transformCommands(commands,[1/w,0,0,1/h,-x/w,-y/h]);
  return {id,kind:'freeform',x,y,w,h,...style,text:'',path:{paths:[{commands:normalized,fill:!!style.fill,stroke:!!style.stroke}]}};
}
export function svgToEditableGroup(svgText,picture={}){
  if(typeof svgText!=='string'||svgText.length>1000000||/<!DOCTYPE|<!ENTITY|<\w+:/i.test(svgText))fail('큰 파일·DTD·네임스페이스 요소');
  validateXml(svgText);
  if(picture.crop&&Object.values(picture.crop).some(v=>Number(v)!==0)||picture.shadow||picture.radius||picture.border||(picture.opacity??1)!==1)fail('그림 자르기·효과를 먼저 재설정하세요');
  const root=parseXml(svgText);if(root.name!=='svg'||root.attrs.xmlns&&root.attrs.xmlns!=='http://www.w3.org/2000/svg')fail('SVG 문서가 아닙니다');
  let vb=root.attrs.viewBox?.trim().split(/[\s,]+/).map(v=>finite(v));
  if(!vb)vb=[0,0,finite(root.attrs.width),finite(root.attrs.height)];if(vb.length!==4||vb[2]<=0||vb[3]<=0)fail('viewBox 또는 크기');
  const w=picture.w??vb[2],h=picture.h??vb[3];if(![w,h,picture.x??0,picture.y??0].every(Number.isFinite)||w<=0||h<=0)fail('그림 크기');
  const aspect=(root.attrs.preserveAspectRatio??'xMidYMid meet').trim();let sx=w/vb[2],sy=h/vb[3],tx=-vb[0]*sx,ty=-vb[1]*sy;
  if(aspect!=='none'){
    const match=/^(xMin|xMid|xMax)(YMin|YMid|YMax)(?:\s+meet)?$/.exec(aspect);if(!match)fail('자르기를 포함한 preserveAspectRatio');
    sx=sy=Math.min(sx,sy);tx=-vb[0]*sx+(w-vb[2]*sx)*({xMin:0,xMid:.5,xMax:1}[match[1]]);ty=-vb[1]*sy+(h-vb[3]*sy)*({YMin:0,YMid:.5,YMax:1}[match[2]]);
  }
  const base=[sx,0,0,sy,tx,ty],items=[],id=picture.id||'svg'+Date.now().toString(36)+Math.random().toString(36).slice(2,7);let nodes=0,commandsCount=0;
  const visit=(node,parent,m,depth)=>{
    if(++nodes>512||depth>32)fail('너무 복잡한 SVG');
    if(['title','desc','metadata'].includes(node.name))return;
    if(node.name==='style'&&/^(?:\s*\.[\w-]+\s*\{\s*\}\s*)*$/.test(node.text))return; // Office icon bundles contain inert theme-class declarations.
    if(node.attrs.xmlns&&node.attrs.xmlns!=='http://www.w3.org/2000/svg')fail('다른 네임스페이스');
    const style=styleFor(node,parent);if(style.display==='none')return;
    const transform=multiplyMatrix(m,parseSvgTransform(node.attrs.transform));
    if(node===root||node.name==='g'){
      if(opacity(style.opacity)!==1)fail('그룹 투명도');
      for(const c of node.children)visit(c,style,transform,depth+1);return;
    }
    if(style.visibility==='hidden'||style.visibility==='collapse')return;
    if(node.name==='defs'&&!node.children.length)return;
    if(node.children.some(c=>!['title','desc'].includes(c.name)))fail('도형 내부의 중첩 요소');
    const commands=primitive(node);if(!commands)return;
    commandsCount+=commands.length;if(commandsCount>10000)fail('경로 점 수 제한');
    const fill=paint(style.fill,style.color),stroke=paint(style.stroke??'none',style.color),o=opacity(style.opacity);
    const sw=finite(style['stroke-width'],1),xScale=Math.hypot(transform[0],transform[1]),yScale=Math.hypot(transform[2],transform[3]);
    if(sw<0)fail('음수 선 두께');
    if(stroke.color&&sw&&((Math.abs(xScale-yScale)>1e-7*Math.max(xScale,yScale))||Math.abs(transform[0]*transform[2]+transform[1]*transform[3])>1e-7*xScale*yScale))fail('비균등 확대·전단된 윤곽선');
    const cap={butt:'flat',round:'rnd',square:'sq'}[style['stroke-linecap']??'butt'],join=style['stroke-linejoin']??'miter';if(!cap||!['miter','round','bevel'].includes(join))fail('선 끝 또는 모서리 형식');
    if(style['stroke-miterlimit']&&finite(style['stroke-miterlimit'])!==4)fail('사용자 지정 이음 한도');
    const s={fill:node.name==='line'?null:fill.color,stroke:sw?stroke.color:null,strokeWidth:sw*xScale,fillOpacity:fill.alpha*opacity(style['fill-opacity'])*o,strokeOpacity:stroke.alpha*opacity(style['stroke-opacity'])*o,lineCap:cap,lineJoin:join};
    // SVG element opacity composites fill+stroke together. Per-paint alpha alone
    // would darken their overlap; this model has no isolated-compositing property.
    if(o<1&&s.fill&&s.stroke&&s.fillOpacity>0&&s.strokeOpacity>0)fail('채우기와 선을 함께 합성하는 개체 투명도');
    if(!s.fill&&!s.stroke)return;
    const transformed=transformCommands(commands,transform);
    const b=normalizedPathBounds({paths:[{commands:transformed}]});
    const margin=s.stroke?s.strokeWidth*(join==='miter'?2:cap==='sq'?Math.SQRT1_2:.5):0;
    if(b.minX-margin<-.01||b.minY-margin<-.01||b.maxX+margin>w+.01||b.maxY+margin>h+.01)fail('SVG 화면 밖으로 잘리는 경로');
    const sections=[];for(const c of transformed){if(c[0]==='M')sections.push([]);sections.at(-1).push(c);}
    if((style['fill-rule']==='evenodd'||sections.length>1)&&s.fill){
      const rings=flattenCommands(transformed,.2),groups=arrangeRegions([[{rings,evenodd:style['fill-rule']==='evenodd'}]]);
      for(const g of groups)for(const component of geometryComponents(g.rings))items.push(geometryShape(component,{...s,stroke:null,text:''},id+'-'+items.length));
      if(s.stroke)for(const section of sections)items.push(commandShape(section,{...s,fill:null},id+'-'+items.length));
    }else for(const section of sections)items.push(commandShape(section,s,id+'-'+items.length));
    if(items.length>256)fail('변환 결과가 256개 도형을 초과합니다');
  };
  visit(root,{fill:'#000000',stroke:'none',color:'#000000'},base,0);
  if(!items.length)fail('변환할 도형이 없습니다');
  return {id,kind:'group',name:picture.name??'변환한 SVG',x:picture.x??0,y:picture.y??0,w,h,groupSize:{w,h},groupItems:items,
    ...(picture.rot?{rot:picture.rot}:{}),...(picture.flip?{flip:true}:{}),...(picture.flipV?{flipV:true}:{}),...(picture.hidden?{hidden:true}:{}),...(picture.hyperlink?{hyperlink:structuredClone(picture.hyperlink)}:{}),...(picture.alt?{alt:picture.alt}:{})};
}

function validateXml(text){
  const token=/<!--[\s\S]*?-->|<\?xml[\s\S]*?\?>|<(\/?)([A-Za-z][\w.-]*)((?:\s+[\w:.-]+\s*=\s*(?:"[^"]*"|'[^']*'))*)\s*(\/?)>/g;
  const stack=[];let last=0,count=0,rootCount=0,m;
  while((m=token.exec(text))){const gap=text.slice(last,m.index);if(gap.includes('<')||(!stack.length&&gap.trim()))fail('잘못된 XML 문법');last=token.lastIndex;if(!m[2])continue;
    if(++count>1024)fail('너무 많은 SVG 요소');
    if(m[1]){if(m[3].trim()||m[4]||stack.pop()!==m[2])fail('맞지 않는 닫는 태그');}
    else {if(!stack.length&&++rootCount>1)fail('중복 루트 요소');const names=new Set();for(const a of m[3].matchAll(/([\w:.-]+)\s*=\s*(?:"[^"]*"|'[^']*')/g)){if(names.has(a[1]))fail('중복 속성');names.add(a[1]);}if(!m[4])stack.push(m[2]);}
  }
  if(stack.length||text.slice(last).trim()||!rootCount)fail('완성되지 않은 XML');
}
