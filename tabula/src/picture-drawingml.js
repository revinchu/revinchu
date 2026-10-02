// 그림의 표준 DrawingML 효과와 원본 복원용 작은 확장. 이미지 바이트는 별도 관계로 보관한다.
import { normalizePictureVisual, needsPictureBake } from './picture-filter.js';
import { child, descendants, esc } from './xml.js';
import { fromBase64 } from './vba.js';
const EMU = 9525;
export const PICTURE_EDIT_URI = '{7B33E95F-8515-43EC-90B0-574958504943}';
const NS = 'https://wixel.app/picture/edit/1';
const MEDIA = ['src','png','emf','originalSrc','originalPng','originalEmf'];
const TINTS = {blue:'#266BD9',green:'#2E994D',orange:'#F26B1A'};
const finite = (n,d=0) => Number.isFinite(Number(n)) ? Number(n) : d;
const clamp = (n,a,b) => Math.min(b,Math.max(a,n));
const pc = n => Math.round(n*100000);
const hex = c => String(c).slice(1).toUpperCase();
const hash = bytes => { let n=2166136261; for(const b of bytes)n=Math.imul(n^b,16777619);return bytes.length+':'+(n>>>0).toString(16); };
export function pictureMediaSource(p) {
  if (needsPictureBake(p) && !p.effectPng) throw new Error('그림 보정 효과를 저장용 이미지로 준비하지 못했습니다. 그림을 다시 저장하세요.');
  return p.effectPng ?? p.emf ?? (p.png && /^data:image\/svg\+xml/.test(p.src??'') ? p.png : p.src);
}
function colorsXml(p) {
  if(p.effectPng)return '';
  const e=normalizePictureVisual(p),c=e.correction,v=e.color;
  let out=c.brightness||c.contrast?`<a:lum bright="${pc(c.brightness)}" contrast="${pc(c.contrast)}"/>`:'';
  if(v.recolor==='grayscale')out+='<a:grayscl/>';
  else {const duo=v.duotone??(TINTS[v.recolor]?['#000000',TINTS[v.recolor]]:null);if(duo)out+=`<a:duotone>${duo.map(c=>`<a:srgbClr val="${hex(c)}"/>`).join('')}</a:duotone>`;}
  return out;
}
export function pictureBlipXml(p) {
  const e=normalizePictureVisual(p);
  return (e.opacity!==1?`<a:alphaModFix amt="${pc(e.opacity)}"/>`:'')+colorsXml(p);
}
export function pictureGeometryXml(p) {
  const e=normalizePictureVisual(p);
  return e.radius?`<a:prstGeom prst="roundRect"><a:avLst><a:gd name="adj" fmla="val ${Math.round(e.radius/Math.min(p.w,p.h)*100000)}"/></a:avLst></a:prstGeom>`:'<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>';
}
export function pictureBorderXml(p) {
  if(!/^#[0-9a-f]{6}$/i.test(p.border??''))return '';
  const dash={solid:'solid',dash:'dash',dot:'sysDot',dashDot:'dashDot'}[normalizePictureVisual(p).borderDash];
  return `<a:ln w="${Math.round(Math.max(0,finite(p.borderW,2))*EMU)}"><a:solidFill><a:srgbClr val="${hex(p.border)}"/></a:solidFill><a:prstDash val="${dash}"/></a:ln>`;
}
export function pictureEffectXml(p) {
  const e=normalizePictureVisual(p),parts=[];
  // CT_EffectList 순서: blur,fillOverlay,glow,innerShadow,outerShadow,presetShadow,reflection,softEdge.
  if(e.glow?.size)parts.push(`<a:glow rad="${Math.round(e.glow.size*EMU)}"><a:srgbClr val="${hex(e.glow.color)}"><a:alpha val="${pc(e.glow.opacity)}"/></a:srgbClr></a:glow>`);
  if(e.shadow){const s=e.shadow;parts.push(`<a:outerShdw blurRad="${Math.round(s.blur*EMU)}" dist="${Math.round(Math.hypot(s.dx,s.dy)*EMU)}" dir="${Math.round((Math.atan2(s.dy,s.dx)*180/Math.PI+360)%360*60000)}" algn="ctr" rotWithShape="1"><a:srgbClr val="${hex(s.color)}"><a:alpha val="${pc(s.opacity)}"/></a:srgbClr></a:outerShdw>`);}
  if(e.reflection?.size){const r=e.reflection;parts.push(`<a:reflection blurRad="0" stA="${pc(r.opacity)}" stPos="0" endA="0" endPos="${pc(r.size)}" dist="${Math.round(r.gap*EMU)}" dir="5400000" fadeDir="5400000" sx="100000" sy="-100000" algn="bl" rotWithShape="1"/>`);}
  if(e.softEdge)parts.push(`<a:softEdge rad="${Math.round(e.softEdge*EMU)}"/>`);
  return parts.length?`<a:effectLst>${parts.join('')}</a:effectLst>`:'';
}
export function readPictureAdjustments(p,blip,sp,colorOf,opacityOf) {
  const lum=child(blip,'lum');if(lum)p.correction={brightness:clamp(finite(lum.attrs.bright)/100000,-1,1),contrast:clamp(finite(lum.attrs.contrast)/100000,-1,1),sharpness:0};
  if(child(blip,'grayscl'))p.color={recolor:'grayscale'};
  const duo=child(blip,'duotone');if(duo){const colors=duo.children.map(c=>colorOf({children:[c]})).filter(Boolean);if(colors.length===2){const preset=Object.keys(TINTS).find(k=>colors[0].toUpperCase()==='#000000'&&colors[1].toUpperCase()===TINTS[k]);p.color=preset?{recolor:preset}:{duotone:colors};}}
  const dash=child(child(sp,'ln'),'prstDash')?.attrs.val;if(dash)p.borderDash=({solid:'solid',dash:'dash',sysDash:'dash',dot:'dot',sysDot:'dot',dashDot:'dashDot',sysDashDot:'dashDot'})[dash]??'solid';
  const fx=child(sp,'effectLst'),glow=child(fx,'glow'),soft=child(fx,'softEdge'),reflection=child(fx,'reflection');
  if(glow)p.glow={color:colorOf(glow)??'#5b9bd5',size:Math.max(0,finite(glow.attrs.rad)/EMU),opacity:opacityOf(glow)};
  if(soft)p.softEdge=Math.max(0,finite(soft.attrs.rad)/EMU);
  if(reflection)p.reflection={opacity:clamp(finite(reflection.attrs.stA,100000)/100000,0,1),size:clamp(finite(reflection.attrs.endPos,100000)/100000,0,1),gap:Math.max(0,finite(reflection.attrs.dist)/EMU)};
}
/** Return a single a:ext (caller merges its cNvPr extLst). */
export function pictureMetadataXml(p,embedSource) {
  const baked=!!p.effectPng,original=MEDIA.some(k=>k.startsWith('original')&&p[k]);
  if(!baked&&!original)return '';
  const source=pictureMediaSource(p),m=/^data:[^;,]+;base64,(.*)$/s.exec(source??'');if(!m)return '';
  const metadata={version:1,baked,hash:hash(fromBase64(m[1])),native:colorsXml(p)},refs=[];
  // Excel tracks r:embed attributes when renumbering/retaining media. An rId hidden
  // inside JSON is not a relationship and its original image is discarded on save.
  for(const key of MEDIA)if((baked||key.startsWith('original'))&&p[key]){const rel=embedSource(p[key]);if(rel)refs.push(`<wx:source key="${key}" r:embed="${rel}"/>`);}
  for(const key of ['originalWidth','originalHeight'])if(Number.isFinite(p[key]))metadata[key]=p[key];
  if(baked){const e=normalizePictureVisual(p);metadata.correction=e.correction;metadata.color=e.color;metadata.artistic=e.artistic;}
  return `<a:ext uri="${PICTURE_EDIT_URI}"><wx:picture xmlns:wx="${NS}" json="${esc(JSON.stringify(metadata))}">${refs.join('')}</wx:picture></a:ext>`;
}
export function restorePictureMetadata(p,nv,bytes,readSource) {
  const el=descendants(nv,'picture').find(x=>x.attrs['xmlns:wx']===NS);if(!el?.attrs.json||el.attrs.json.length>20000)return;
  try {
    const m=JSON.parse(el.attrs.json);if(m.version!==1||m.hash!==hash(bytes)||m.native!==colorsXml(p))return;
    const restored={};for(const ref of el.children??[]){const k=ref.attrs.key,rel=ref.attrs['r:embed'];if(ref.name==='source'&&MEDIA.includes(k)&&typeof rel==='string'){const src=readSource(rel);if(src)restored[k]=src;}}
    if(m.baked&&!restored.src)return;
    Object.assign(p,restored);
    for(const k of ['originalWidth','originalHeight'])if(Number.isFinite(m[k])&&m[k]>0&&m[k]<=100000)p[k]=m[k];
    if(m.baked){const e=normalizePictureVisual(m);p.correction=e.correction;p.color=e.color;p.artistic=e.artistic;delete p.effectPng;}
  } catch { /* Unknown/corrupt optional editing metadata never replaces native picture data. */ }
}
