// Standalone, self-contained SVG and raster exports. Never changes workbook objects.
import { shapeSvg, isShapeLine, shapeLineEnds } from './shapes.js';
import { groupSvg } from './object-group.js';
import { pictureSvg } from './picture-render.js';
import { pictureExportBounds } from './picture-export.js';
import { normalizePictureVisual } from './picture-filter.js';
import { normalizedPathBounds } from './shape-path.js';
import { isSmartArt, smartArtParts } from './smartart.js';
import { smartArtSvg } from './smartart-render.js';
import { safeUrl } from './safe-html.js';
import { esc, decodeEntities } from './xml.js';

export const OBJECT_IMAGE_MAX_SIDE = 4096;
export const OBJECT_IMAGE_MAX_PIXELS = 16000000;
const MAX_XML = 64 * 1024 * 1024;
const SVG_NS = 'http://www.w3.org/2000/svg';
const SVG_TAGS = new Set('svg g path rect circle ellipse line polyline polygon text tspan defs linearGradient radialGradient stop filter clipPath mask marker pattern use image title desc feDropShadow feGaussianBlur feColorMatrix feOffset feFlood feComposite feMerge feMergeNode feMorphology feComponentTransfer feFuncR feFuncG feFuncB feFuncA feBlend feTile feTurbulence feDisplacementMap feConvolveMatrix feDiffuseLighting feSpecularLighting feDistantLight fePointLight feSpotLight'.split(' '));
const CSS_PROPS = new Set('fill fill-opacity fill-rule stroke stroke-width stroke-opacity stroke-linecap stroke-linejoin stroke-dasharray stroke-dashoffset stroke-miterlimit color opacity font-family font-size font-weight font-style text-anchor dominant-baseline alignment-baseline text-decoration white-space overflow display visibility filter clip-path mask marker-start marker-mid marker-end paint-order'.split(' '));
const PAINT = new Set('fill stroke filter clip-path mask marker marker-start marker-mid marker-end cursor'.split(' '));
const fail = message => { throw new Error(message); };
const n = (value, fallback=0) => Number.isFinite(Number(value)) ? Number(value) : fallback;
const dimensions = o => ({ w:Math.max(1,n(o.w,1)), h:Math.max(1,n(o.h,1)) });
const union = (a,b) => !a ? {...b} : {x:Math.min(a.x,b.x),y:Math.min(a.y,b.y),w:Math.max(a.x+a.w,b.x+b.w)-Math.min(a.x,b.x),h:Math.max(a.y+a.h,b.y+b.h)-Math.min(a.y,b.y)};
function transformBounds(b, fn) {
  let result=null;
  for(const [x,y] of [[b.x,b.y],[b.x+b.w,b.y],[b.x,b.y+b.h],[b.x+b.w,b.y+b.h]]) { const p=fn(x,y); result=union(result,{x:p[0],y:p[1],w:0,h:0}); }
  return result;
}
function rotated(b,o) {
  const {w,h}=dimensions(o),a=n(o.rot)*Math.PI/180,snap=v=>Math.abs(v)<1e-12?0:Math.abs(v-1)<1e-12?1:Math.abs(v+1)<1e-12?-1:v,co=snap(Math.cos(a)),sn=snap(Math.sin(a));
  return transformBounds(b,(x,y)=>[w/2+(x-w/2)*co-(y-h/2)*sn,h/2+(x-w/2)*sn+(y-h/2)*co]);
}
function checkedTree(object) {
  let count=0;
  const walk=(o,depth)=>{
    if(!o||typeof o!=='object'||++count>2000||depth>32)fail('그림으로 저장할 개체가 너무 복잡합니다.');
    for(const key of ['x','y','w','h','rot']) if(o[key]!==undefined&&(!Number.isFinite(Number(o[key]))||Math.abs(Number(o[key]))>1000000))fail('개체의 위치 또는 크기가 올바르지 않습니다.');
    if(n(o.w)<0||n(o.h)<0)fail('음수 크기의 개체는 먼저 크기를 조정하세요.');
    if(o.groupItems!==undefined&&!Array.isArray(o.groupItems))fail('그룹 구성이 올바르지 않습니다.');
    if(o.groupSize&&['w','h'].some(k=>!Number.isFinite(o.groupSize[k])||o.groupSize[k]<=0))fail('그룹의 기준 크기가 올바르지 않습니다.');
    for(const child of o.groupItems??[])walk(child,depth+1);
  };
  walk(object,0);
}
function localBounds(o,kind) {
  const {w,h}=dimensions(o);
  if(kind==='picture'||o.kind==='picture') {
    const b=pictureExportBounds({...o,w,h,rot:0}),e=normalizePictureVisual(o);
    // Combined filters are sequential: reserve their combined bleed as well.
    const extra=(e.glow?.size||0)*2+(e.shadow?.blur||0)*2+Math.max(Math.abs(e.shadow?.dx||0),Math.abs(e.shadow?.dy||0));
    return union(b,{x:-extra,y:-extra,w:w+extra*2,h:h+extra*2});
  }
  if(o.kind==='group'||isSmartArt(o)) {
    const smart=isSmartArt(o),items=smart?smartArtParts(o):(o.groupItems??[]),bw=smart?w:n(o.groupSize?.w,w)||w,bh=smart?h:n(o.groupSize?.h,h)||h;
    let b={x:0,y:0,w:bw,h:bh};
    for(const child of items)if(!child.hidden) {
      const cb=rotated(localBounds(child,child.kind==='picture'?'picture':'shape'),child);
      b=union(b,{...cb,x:cb.x+n(child.x),y:cb.y+n(child.y)});
    }
    return transformBounds(b,(x,y)=>[(o.flip?bw-x:x)*w/bw,(o.flipV?bh-y:y)*h/bh]);
  }
  let b={x:0,y:0,w,h};
  const path=normalizedPathBounds(o.path);
  if(path) b=union(b,transformBounds({x:path.minX*w,y:path.minY*h,w:(path.maxX-path.minX)*w,h:(path.maxY-path.minY)*h},(x,y)=>[o.flip?w-x:x,o.flipV?h-y:y]));
  const sw=o.stroke?Math.max(0,n(o.strokeWidth,1)):0,ends=shapeLineEnds(o);
  const arrows=ends.headEnd.type!=='none'||ends.tailEnd.type!=='none';
  const sd=o.shadow?typeof o.shadow==='object'?o.shadow:{}:null;
  const bleed=(sd?Math.max(Math.abs(n(sd.dx,2.5)),Math.abs(n(sd.dy,2.5)))+n(sd.blur,2.5)*3:0)+(o.glow?Math.max(0,n(o.glow.size,5))*3:0)+Math.max(0,n(o.soft))*3;
  const pad=bleed+(arrows?sw*8:(path||isShapeLine(o))?sw*(o.lineJoin==='miter'?2:.5):0);
  return {x:b.x-pad,y:b.y-pad,w:b.w+pad*2,h:b.h+pad*2};
}

/** Bounds in the object's local coordinates, including its rotation and effects. */
export function objectImageBounds(object, kind='shape') {
  checkedTree(object);
  return rotated(localBounds(object,kind),object);
}

function cssValue(value) {
  if(/[\\\u0000-\u001f\u007f]|\/\*|@import|expression\s*\(|image-set\s*\(/i.test(value))fail('SVG에 저장할 수 없는 스타일이 있습니다.');
  const rest=value.replace(/url\(\s*(?:"([^"]*)"|'([^']*)'|([^\s)'" ]+))\s*\)/gi,(_,a,b,c)=>{
    if(!safeUrl(a??b??c,'reference'))fail('SVG의 외부 자원 참조는 저장할 수 없습니다.');return '';
  });
  if(/url\s*\(/i.test(rest))fail('SVG 자원 참조가 올바르지 않습니다.');
  return value;
}
function imageData(source, depth=0) {
  if(depth>8||typeof source!=='string'||source.length>MAX_XML)fail('그림 데이터가 너무 크거나 복잡합니다.');
  if(!safeUrl(source,'image')||!/^data:image\//i.test(source))fail('연결된 그림은 먼저 파일 데이터로 불러온 뒤 저장하세요.');
  const m=/^data:image\/([a-z0-9+.-]+)((?:;[^,]*)?),(.*)$/is.exec(source);
  if(!m)fail('그림 데이터 형식이 올바르지 않습니다.');
  if(m[1].toLowerCase()==='svg+xml') {
    let text;
    try { text=/;base64(?:;|$)/i.test(m[2])?new TextDecoder('utf-8',{fatal:true}).decode(Uint8Array.from(atob(m[3]),c=>c.charCodeAt(0))):decodeURIComponent(m[3]); }
    catch { fail('SVG 그림 데이터를 읽지 못했습니다.'); }
    return 'data:image/svg+xml;charset=utf-8,'+encodeURIComponent(safeObjectSvg(text,depth+1));
  }
  if(!['png','jpeg','gif','webp','avif','bmp','x-icon'].includes(m[1].toLowerCase())||!/^;base64$/i.test(m[2])||! /^(?:[A-Za-z0-9+/]{4})*(?:[A-Za-z0-9+/]{2}==|[A-Za-z0-9+/]{3}=)?$/.test(m[3]))fail('지원하지 않는 그림 데이터 형식입니다.');
  return source;
}

function svgAttributes(tag) {
  const result=new Map(),re=/([A-Za-z_][\w:.-]*)="([^"]*)"/g;let m;
  while((m=re.exec(tag)))result.set(m[1],decodeEntities(m[2]));return result;
}
/** Expand sanitized SVG pictures to editable nested vectors, retaining crop/filter
 * wrappers. Every local paint/clip/use id is namespaced for duplicate icons.
 * SVGs without a usable viewBox/px size remain embedded vector images.
 */
function inlineSvgPictures(svg) {
  let serial=0;
  const used=new Set([...svg.matchAll(/\sid="([^"]*)"/g)].map(m=>m[1]));
  const expand=(text,depth=0)=>text.replace(/<image\b[^>]*\/>/g,tag=>{
    if(depth>=8)return tag;
    const outer=svgAttributes(tag),src=outer.get('href')??outer.get('xlink:href');
    if(!src?.startsWith('data:image/svg+xml;charset=utf-8,'))return tag;
    const inner=decodeURIComponent(src.slice(src.indexOf(',')+1)),root=/^\s*<svg\b[^>]*>/.exec(inner);
    if(!root)return tag;
    const attrs=svgAttributes(root[0]);let viewBox=attrs.get('viewBox');
    if(!viewBox) {
      const length=(s,fallback)=>s===undefined?fallback:/^\d+(?:\.\d+)?(?:px)?$/.test(s)?Number(s.replace(/px$/,'')):NaN;
      const w=length(attrs.get('width'),300),h=length(attrs.get('height'),150);
      if(!(w>0&&h>0))return tag;viewBox=`0 0 ${w} ${h}`;
    }
    const nums=viewBox.trim().split(/[\s,]+/).map(Number);if(nums.length!==4||nums.some(v=>!Number.isFinite(v))||nums[2]<=0||nums[3]<=0)return tag;
    const content=root[0].endsWith('/>')?'':inner.slice(root[0].length).replace(/<\/svg>\s*$/,'');
    // SVG image sources retain their own aspect policy unless the outer image
    // explicitly stretches them. This differs from raster image letterboxing.
    const aspect=outer.get('preserveAspectRatio')==='none'?'none':attrs.get('preserveAspectRatio')??outer.get('preserveAspectRatio')??'xMidYMid meet';
    for(const k of ['x','y','width','height','viewBox','preserveAspectRatio'])attrs.delete(k);
    for(const [k,v] of outer)if(!['href','xlink:href','xmlns'].includes(k))attrs.set(k,v);
    attrs.set('viewBox',viewBox);attrs.set('preserveAspectRatio',aspect);
    attrs.set('style',((attrs.get('style')??'').replace(/(?:^|;)\s*overflow\s*:[^;]*/g,'')+';overflow:hidden').replace(/^;/,''));
    let prefix;do{prefix=`wixel-inline-${++serial}-`;}while([...used].some(v=>v.startsWith(prefix)));
    const ids=new Map(),ref=id=>{if(!ids.has(id))ids.set(id,prefix+ids.size);return ids.get(id);};
    const serialized=`<svg${[...attrs].map(([k,v])=>` ${k}="${esc(v)}"`).join('')}>${expand(content,depth+1)}</svg>`;
    return serialized.replace(/<[^>]+>/g,(element,offset)=>element.replace(/([A-Za-z_][\w:.-]*)="([^"]*)"/g,(all,k,v)=>{
      // Attributes originating on the outer image still refer to the outer SVG.
      if(offset===0&&outer.has(k)&&!['href','xlink:href','xmlns'].includes(k))return all;
      let value=decodeEntities(v);
      if(k==='id')value=ref(value);
      else if((k==='href'||k==='xlink:href')&&value.startsWith('#'))value='#'+ref(value.slice(1));
      else value=value.replace(/url\(\s*(?:"#([^"]*)"|'#([^']*)'|#([^\s)'" ]+))\s*\)/g,(_,a,b,c)=>`url(#${ref(a??b??c)})`);
      return `${k}="${esc(value)}"`;
    }));
  });
  return expand(svg);
}

/** Strict SVG serialization; no HTML, scripts, DTD, CSS imports or external loads.
 * Unlike the tolerant OOXML parser, this preserves mixed text/tspan ordering.
 * The existing safeUrl policy is reused with the stricter self-contained rule.
 */
export function safeObjectSvg(svg, depth=0, deferredImages=null) {
  if(typeof svg!=='string'||svg.length>MAX_XML||depth>8)fail('SVG 파일이 너무 크거나 복잡합니다.');
  const source=svg.replace(/^\s*<\?xml\s[^?]*\?>/i,''),stack=[],out=[];
  let pos=0,root=false,closed=false,count=0;
  while(pos<source.length) {
    if(source.startsWith('<!--',pos)) {const end=source.indexOf('-->',pos+4);if(end<0)fail('SVG 주석이 올바르지 않습니다.');pos=end+3;continue;}
    if(source.startsWith('<![CDATA[',pos)) {const end=source.indexOf(']]>',pos+9);if(end<0||!stack.length)fail('SVG 텍스트가 올바르지 않습니다.');out.push(esc(source.slice(pos+9,end)));pos=end+3;continue;}
    if(source[pos]!=='<') {let end=source.indexOf('<',pos);if(end<0)end=source.length;const text=source.slice(pos,end);if(!stack.length&&text.trim())fail('SVG 문서 구조가 올바르지 않습니다.');out.push(esc(decodeEntities(text)));pos=end;continue;}
    const match=/^<(\/?)(([A-Za-z][\w.-]*:)?[A-Za-z][\w.-]*)((?:\s+[A-Za-z_][\w:.-]*\s*=\s*(?:"[^"]*"|'[^']*'))*)\s*(\/?)>/.exec(source.slice(pos));
    if(!match)fail('SVG 문법 또는 선언을 저장할 수 없습니다.');
    pos+=match[0].length;
    const close=!!match[1],name=match[2],empty=!!match[5];
    if(!SVG_TAGS.has(name))fail(`SVG의 ${name} 요소는 그림 저장에서 지원하지 않습니다.`);
    if(close) {if(match[4].trim()||empty||stack.pop()!==name)fail('SVG 요소의 여닫는 순서가 올바르지 않습니다.');out.push(`</${name}>`);if(!stack.length)closed=true;continue;}
    if(++count>100000||stack.length>64||closed)fail('SVG 문서 구조가 너무 복잡합니다.');
    if(!root) {if(name!=='svg')fail('SVG 루트 요소가 필요합니다.');root=true;}
    const attrs=[],seen=new Set(),re=/([A-Za-z_][\w:.-]*)\s*=\s*(?:"([^"]*)"|'([^']*)')/g;
    let am;
    while((am=re.exec(match[4]))) {
      const key=am[1],low=key.toLowerCase();let value=decodeEntities(am[2]??am[3]);
      if(seen.has(key))fail('SVG에 중복 속성이 있습니다.');seen.add(key);
      if(low.startsWith('on')||['src','srcdoc','xml:base','base','target'].includes(low)||key.includes(':')&&!['xmlns:xlink','xlink:href','xml:space'].includes(key))fail('SVG에 실행 또는 외부 연결 속성이 있습니다.');
      if(key==='xmlns'&&value!==SVG_NS||key==='xmlns:xlink'&&value!=='http://www.w3.org/1999/xlink')fail('SVG 네임스페이스가 올바르지 않습니다.');
      if(low==='href'||low==='xlink:href') {
        if(name==='image') {
          if(deferredImages&&!/^data:/i.test(value)) {
            if(!safeUrl(value,'image')||! /^(https?:|blob:)/i.test(value))fail('SVG의 그림 주소가 올바르지 않습니다.');
            if(!deferredImages.has(value))deferredImages.set(value,'data:image/png;base64,'+btoa('wixel-export-image-'+deferredImages.size));
            value=deferredImages.get(value);
          } else value=imageData(value,depth);
        }
        else if(name!=='use'||!safeUrl(value,'reference'))fail('SVG의 외부 참조는 저장할 수 없습니다.');
      }
      if(low==='style') {
        const declarations=[];
        for(const p of value.split(';'))if(p.trim()) {const at=p.indexOf(':'),prop=p.slice(0,at).trim().toLowerCase();if(at<0||!CSS_PROPS.has(prop))fail('지원하지 않는 SVG 스타일입니다.');declarations.push(`${prop}:${cssValue(p.slice(at+1).trim())}`);}
        value=declarations.join(';');
      } else if(PAINT.has(low))cssValue(value);
      if(low.startsWith('data-')||low.startsWith('aria-')||['role','tabindex'].includes(low))continue;
      attrs.push(` ${key}="${esc(value)}"`);
    }
    if(name==='svg'&&!stack.length&&!seen.has('xmlns'))attrs.push(` xmlns="${SVG_NS}"`);
    out.push(`<${name}${attrs.join('')}${empty?'/':''}>`);
    if(!empty)stack.push(name);else if(!stack.length)closed=true;
  }
  if(!root||stack.length)fail('SVG 문서가 완전하지 않습니다.');
  return out.join('');
}

/** Inline explicitly authorized image sources, then validate the complete SVG.
 * Returns a string; range and object callers can share the same preparation.
 */
export async function prepareImageSvg(svg,{resolveImage}={}) {
  const deferred=new Map();let prepared=safeObjectSvg(svg,0,deferred);
  if(deferred.size&&!resolveImage)fail('연결된 그림은 먼저 파일 데이터로 불러온 뒤 저장하세요.');
  // Validate the complete document before invoking any resource resolver.
  for(const [url,placeholder] of deferred) {
    const src=imageData(await resolveImage(url));
    prepared=prepared.split(placeholder).join(esc(src));
  }
  return safeObjectSvg(inlineSvgPictures(safeObjectSvg(prepared)));
}

function renderShape(o) {
  if(isSmartArt(o))return smartArtSvg(o);
  if(o.kind==='group')return groupSvg(o,renderShape);
  return shapeSvg(o);
}
function renderObject(o,kind) {
  if(kind==='picture'||o.kind==='picture')return pictureSvg(o,'export-picture');
  if(o.kind==='group'||isSmartArt(o))return renderShape(o);
  // The shared group renderer includes rich text, padding, wrapping and rotation.
  const child={...o,x:0,y:0,rot:0,hidden:false};
  if(isShapeLine(o)){child.text='';child.paras=undefined;}
  return groupSvg({id:'export-text',kind:'group',w:o.w,h:o.h,groupItems:[child]},renderShape);
}

/** Caller may explicitly supply resolveImage(url) -> embedded data URL. No fetch by default. */
export async function objectImageSvg(object,{kind='shape',title='',resolveImage}={}) {
  if(!['shape','picture'].includes(kind))fail('지원하지 않는 그림 저장 대상입니다.');
  checkedTree(object);
  const copy=structuredClone(object),warnings=[],sources=new Map();let id=0;
  const prepare=async(o,picture=false)=>{
    o.id='export'+(++id);delete o.hyperlink;delete o.macro;
    if(picture||o.kind==='picture') {
      if(o.media)fail('동영상은 그림으로 저장할 수 없습니다.');
      if(o.linked&&!o.linkedPng)fail('연결된 그림의 현재 모습을 먼저 준비하세요.');
      let src=o.linkedPng||o.src;
      if(typeof src!=='string')fail('그림 데이터가 없습니다.');
      if(!src.startsWith('data:')) {
        if(!safeUrl(src,'image')||!resolveImage)fail('연결된 그림은 먼저 파일 데이터로 불러온 뒤 저장하세요.');
        if(!sources.has(src))sources.set(src,Promise.resolve(resolveImage(src)));
        src=await sources.get(src);
      }
      o.src=imageData(src);
      if(/^data:image\/(gif|webp)/i.test(src))warnings.push('움직이는 그림을 PNG 또는 JPEG로 저장하면 한 프레임만 저장됩니다.');
    }
    for(const child of o.groupItems??[])if(!child.hidden)await prepare(child,child.kind==='picture');
  };
  await prepare(copy,kind==='picture');
  const bounds=objectImageBounds(copy,kind),width=Math.max(1,bounds.w),height=Math.max(1,bounds.h),{w,h}=dimensions(copy);
  if(!Number.isFinite(width)||!Number.isFinite(height)||width>65536||height>65536)fail('SVG 그림의 한 변은 65,536픽셀 이하로 설정하세요.');
  const body=renderObject(copy,kind);
  const svg=await prepareImageSvg(`<svg xmlns="${SVG_NS}" width="${width}" height="${height}" viewBox="0 0 ${width} ${height}">${title?`<title>${esc(title)}</title>`:''}<g transform="translate(${-bounds.x},${-bounds.y}) rotate(${n(copy.rot)},${w/2},${h/2})">${body}</g></svg>`);
  return {svg,width,height,bounds,warnings:[...new Set(warnings)]};
}

/** Exact requested raster size; never silently reduces resolution. */
export function objectRasterSize(width,height,scale=2) {
  if(!Number.isFinite(width)||!Number.isFinite(height)||width<=0||height<=0||![1,2,3].includes(scale))fail('그림 크기 또는 배율이 올바르지 않습니다.');
  const w=Math.ceil(width*scale),h=Math.ceil(height*scale);
  if(w>OBJECT_IMAGE_MAX_SIDE||h>OBJECT_IMAGE_MAX_SIDE||w*h>OBJECT_IMAGE_MAX_PIXELS)fail('저장 그림은 한 변 4,096픽셀, 전체 1,600만 픽셀 이하로 설정하세요. 배율을 낮추거나 범위를 줄이세요.');
  return {width:w,height:h};
}

/** Shared by object and cell-range exports. Browser APIs are used only for raster. */
export async function svgImageBlob(svg,{width,height,format='png',scale=2,background='transparent',resolveImage}={}) {
  if(!['png','jpeg','svg'].includes(format)||!['transparent','white'].includes(background))fail('지원하지 않는 그림 저장 형식입니다.');
  if(!Number.isFinite(width)||!Number.isFinite(height)||width<=0||height<=0||width>65536||height>65536)fail('SVG 그림의 크기는 0보다 크고 한 변 65,536픽셀 이하여야 합니다.');
  let safe=await prepareImageSvg(svg,{resolveImage});
  if(background==='white'||format==='jpeg') {
    const root=/^\s*<svg\b[^>]*>/.exec(safe),vb=root&&/\sviewBox="([^"]+)"/.exec(root[0]);
    const box=vb?vb[1].trim().split(/[\s,]+/).map(Number):[0,0,width,height];
    if(box.length!==4||box.some(v=>!Number.isFinite(v))||box[2]<=0||box[3]<=0)fail('SVG의 표시 범위가 올바르지 않습니다.');
    const backgroundRect=`<rect x="${box[0]}" y="${box[1]}" width="${box[2]}" height="${box[3]}" fill="#ffffff"/>`;
    safe=safe.replace(root[0],root[0].endsWith('/>')?root[0].slice(0,-2)+'>'+backgroundRect+'</svg>':root[0]+backgroundRect);
  }
  if(format==='svg')return new Blob([safe],{type:'image/svg+xml;charset=utf-8'});
  const size=objectRasterSize(width,height,scale);
  if(typeof document==='undefined'||typeof Image==='undefined')fail('PNG/JPEG 저장은 브라우저에서 실행하세요.');
  if(document.fonts?.ready)await document.fonts.ready;
  const url=URL.createObjectURL(new Blob([safe],{type:'image/svg+xml;charset=utf-8'})),image=new Image();
  try {
    await new Promise((resolve,reject)=>{
      const timer=setTimeout(()=>{image.onload=image.onerror=null;image.src='';reject(new Error('그림 변환 시간이 초과되었습니다.'));},15000);
      image.onload=()=>{clearTimeout(timer);image.onload=image.onerror=null;resolve();};
      image.onerror=()=>{clearTimeout(timer);image.onload=image.onerror=null;reject(new Error('그림을 변환하지 못했습니다. SVG 또는 원본 그림 형식을 확인하세요.'));};image.src=url;
    });
    const canvas=document.createElement('canvas');canvas.width=size.width;canvas.height=size.height;
    const ctx=canvas.getContext('2d');if(!ctx)fail('그림 저장을 위한 캔버스를 만들지 못했습니다.');
    if(format==='jpeg'||background==='white'){ctx.fillStyle='#ffffff';ctx.fillRect(0,0,size.width,size.height);}
    ctx.drawImage(image,0,0,size.width,size.height);
    return await new Promise((resolve,reject)=>{try{canvas.toBlob(blob=>blob?resolve(blob):reject(new Error('그림 파일을 만들지 못했습니다.')),format==='jpeg'?'image/jpeg':'image/png',.92);}catch{reject(new Error('이 그림의 픽셀을 저장할 수 없습니다. 외부 그림은 먼저 파일로 삽입하세요.'));}});
  } finally {URL.revokeObjectURL(url);}
}

export async function objectImageBlob(object,options={}) {
  const result=await objectImageSvg(object,options);
  return svgImageBlob(result.svg,{...options,width:result.width,height:result.height});
}
