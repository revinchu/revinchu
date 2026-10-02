import { child, kids, esc } from './xml.js';
import { normalizeChartAreaFormat, chartAreaColor } from './chart-area-format.js';

// ISO/IEC 29500 DrawingML CT_ShapeProperties, CT_GradientFillProperties, CT_Shape3D.
// Package/media relationships belong to the caller. This module neither fetches nor writes files.
const EMU = 9525;
const n = (value, fallback = 0) => Number.isFinite(Number(value)) && value !== undefined && value !== null ? Number(value) : fallback;
const emu = v => Math.round(v * EMU);
const rgb = (color, opacity = 1) => `<a:srgbClr val="${color.slice(1).toUpperCase()}">${opacity < 1 ? `<a:alpha val="${Math.round(opacity * 100000)}"/>` : ''}</a:srgbClr>`;
function gradientXml(g, alpha = 1) {
  return `<a:gradFill rotWithShape="1"><a:gsLst>${g.stops.map(([offset,color,opacity])=>`<a:gs pos="${Math.round(offset*100000)}">${rgb(color,opacity*alpha)}</a:gs>`).join('')}</a:gsLst>${g.type==='radial'?'<a:path path="circle"><a:fillToRect l="50000" t="50000" r="50000" b="50000"/></a:path>':`<a:lin ang="${Math.round(g.ang*60000)}" scaled="1"/>`}</a:gradFill>`;
}
/** imageRel(src) -> relationship id (or {id}); image bytes remain the caller's responsibility. */
export function chartAreaFormatXml(value, { tag = 'c:spPr', imageRel, kind = 'chart' } = {}) {
  if (!value) return '';
  const f=normalizeChartAreaFormat(value,kind), out=[];
  if(f.fillMode==='none')out.push('<a:noFill/>');
  else if(f.fillMode==='solid')out.push(`<a:solidFill>${rgb(f.fill,f.fillOpacity)}</a:solidFill>`);
  else if(f.fillMode==='gradient')out.push(gradientXml(f.grad,f.fillOpacity));
  else if(f.fillMode==='pattern')out.push(`<a:pattFill prst="${f.pattern.preset}"><a:fgClr>${rgb(f.pattern.fg,f.fillOpacity)}</a:fgClr><a:bgClr>${rgb(f.pattern.bg,f.fillOpacity)}</a:bgClr></a:pattFill>`);
  else if(f.fillMode==='picture'&&f.picture.src){
    const result=imageRel?.(f.picture.src), rid=typeof result==='object'?result?.id:result;
    if(rid)out.push(`<a:blipFill rotWithShape="1"><a:blip r:embed="${esc(rid)}">${f.picture.opacity*f.fillOpacity<1?`<a:alphaModFix amt="${Math.round(f.picture.opacity*f.fillOpacity*100000)}"/>`:''}</a:blip>${f.picture.mode==='tile'?`<a:tile sx="${Math.round(f.picture.scale*100000)}" sy="${Math.round(f.picture.scale*100000)}" tx="0" ty="0" algn="tl" flip="none"/>`:'<a:stretch><a:fillRect/></a:stretch>'}</a:blipFill>`);
  }
  if(f.lineMode!=='auto')out.push(`<a:ln w="${emu(f.strokeWidth)}" cap="${f.lineCap}" cmpd="${f.compound}" algn="ctr">${f.lineMode==='none'?'<a:noFill/>':f.lineMode==='gradient'?gradientXml(f.strokeGrad,f.strokeOpacity):`<a:solidFill>${rgb(f.stroke,f.strokeOpacity)}</a:solidFill>`}<a:prstDash val="${f.dash||'solid'}"/>${f.lineJoin==='miter'?'<a:miter lim="800000"/>':`<a:${f.lineJoin}/>`}</a:ln>`);
  const effects=[];
  // CT_EffectList sequence: blur, fillOverlay, glow, innerShdw, outerShdw, prstShdw, reflection, softEdge.
  if(f.glow?.size)effects.push(`<a:glow rad="${emu(f.glow.size)}">${rgb(f.glow.color,f.glow.opacity)}</a:glow>`);
  if(f.shadow){const s=f.shadow,dist=Math.hypot(s.dx,s.dy),angle=(Math.atan2(s.dy,s.dx)*180/Math.PI+360)%360;effects.push(`<a:outerShdw blurRad="${emu(s.blur)}" dist="${emu(dist)}" dir="${Math.round(angle*60000)}" sx="${Math.round(s.scale*100000)}" sy="${Math.round(s.scale*100000)}" algn="ctr" rotWithShape="0">${rgb(s.color,s.opacity)}</a:outerShdw>`);}
  if(f.soft)effects.push(`<a:softEdge rad="${emu(f.soft)}"/>`);
  if(effects.length)out.push(`<a:effectLst>${effects.join('')}</a:effectLst>`);
  const d=f.threeD;
  if(d.depth||d.contourWidth||d.bevelTop.type!=='none'||d.bevelBottom.type!=='none'){
    out.push(`<a:scene3d><a:camera prst="orthographicFront"/><a:lightRig rig="${d.lightRig}" dir="tl"><a:rot lat="0" lon="0" rev="${Math.round(d.lightAngle*60000)}"/></a:lightRig></a:scene3d>`);
    const bevel=(b,name)=>b.type==='none'?'':`<a:${name} w="${emu(b.w)}" h="${emu(b.h)}" prst="${b.type}"/>`;
    out.push(`<a:sp3d extrusionH="${emu(d.depth)}" contourW="${emu(d.contourWidth)}" prstMaterial="${d.material}">${bevel(d.bevelTop,'bevelT')}${bevel(d.bevelBottom,'bevelB')}<a:extrusionClr>${rgb(f.fill)}</a:extrusionClr><a:contourClr>${rgb(d.contourColor)}</a:contourClr></a:sp3d>`);
  }
  tag=['c:spPr','cx:spPr','a:spPr'].includes(tag)?tag:'c:spPr';
  return `<${tag}>${out.join('')}</${tag}>`;
}

/** readColor(container) resolves theme/system colors; imageSource(rid, external) resolves package media. */
export function readChartAreaFormat(spPr, { readColor, imageSource, kind = 'chart' } = {}) {
  if(!spPr)return undefined;
  const colorNode=container=>container?.children?.find(c=>['srgbClr','schemeClr','sysClr','scrgbClr','prstClr'].includes(c.name));
  const color=(container,fallback)=>{
    const node=colorNode(container);if(!node)return fallback;
    const custom=readColor?.(container);
    if(custom)return chartAreaColor(typeof custom==='object'?custom.color:custom,fallback);
    if(node.name==='srgbClr')return chartAreaColor('#'+node.attrs.val,fallback);
    if(node.name==='sysClr')return chartAreaColor('#'+node.attrs.lastClr,fallback);
    return fallback;
  };
  const alpha=container=>{const node=colorNode(container),base=child(node,'alpha');let value=base?n(base.attrs.val,100000)/100000:1;for(const t of node?.children??[]){if(t.name==='alphaMod')value*=n(t.attrs.val,100000)/100000;else if(t.name==='alphaOff')value+=n(t.attrs.val)/100000;}return Math.max(0,Math.min(1,value));};
  const gradient=el=>({type:child(el,'path')?'radial':'linear',ang:n(child(el,'lin')?.attrs.ang,5400000)/60000,stops:kids(child(el,'gsLst'),'gs').map(gs=>[n(gs.attrs.pos)/100000,color(gs,'#4472c4'),alpha(gs)])});
  const f={};let recognized=false;
  const solid=child(spPr,'solidFill'),grad=child(spPr,'gradFill'),pattern=child(spPr,'pattFill'),image=child(spPr,'blipFill');
  if(child(spPr,'noFill')){f.fillMode='none';recognized=true;}
  else if(solid){f.fillMode='solid';f.fill=color(solid,'#ffffff');f.fillOpacity=alpha(solid);recognized=true;}
  else if(grad){f.fillMode='gradient';f.grad=gradient(grad);recognized=true;}
  else if(pattern){f.fillMode='pattern';f.pattern={preset:pattern.attrs.prst,fg:color(child(pattern,'fgClr'),'#4472c4'),bg:color(child(pattern,'bgClr'),'#ffffff')};f.fillOpacity=alpha(child(pattern,'fgClr'));recognized=true;}
  else if(image){const blip=child(image,'blip'),rid=blip?.attrs['r:embed']??blip?.attrs['r:link'],tile=child(image,'tile'),result=imageSource?.(rid,!!blip?.attrs['r:link']);f.fillMode='picture';f.picture={...(typeof result==='object'?result:{src:result??''}),mode:tile?'tile':'stretch',scale:n(tile?.attrs.sx,100000)/100000,opacity:n(child(blip,'alphaModFix')?.attrs.amt,100000)/100000};recognized=true;}
  const line=child(spPr,'ln');
  if(line){recognized=true;const solid=child(line,'solidFill'),grad=child(line,'gradFill');f.lineMode=child(line,'noFill')?'none':grad?'gradient':solid?'solid':'auto';f.stroke=color(solid,'#bfbfbf');f.strokeOpacity=alpha(solid);f.strokeWidth=n(line.attrs.w,9525)/EMU;f.compound=line.attrs.cmpd;f.lineCap=line.attrs.cap;f.lineJoin=child(line,'miter')?'miter':child(line,'bevel')?'bevel':'round';f.dash=child(line,'prstDash')?.attrs.val;if(grad)f.strokeGrad=gradient(grad);}
  const effects=child(spPr,'effectLst'),shadow=child(effects,'outerShdw'),glow=child(effects,'glow'),soft=child(effects,'softEdge');
  if(effects)recognized=true;
  if(shadow){const a=n(shadow.attrs.dir)/60000*Math.PI/180,dist=n(shadow.attrs.dist)/EMU;f.shadow={dx:Math.cos(a)*dist,dy:Math.sin(a)*dist,blur:n(shadow.attrs.blurRad)/EMU,color:color(shadow,'#000000'),opacity:alpha(shadow),scale:n(shadow.attrs.sx,100000)/100000};}
  if(glow)f.glow={size:n(glow.attrs.rad)/EMU,color:color(glow,'#ffc000'),opacity:alpha(glow)};
  if(soft)f.soft=n(soft.attrs.rad)/EMU;
  const d=child(spPr,'sp3d'),scene=child(spPr,'scene3d'),light=child(scene,'lightRig');
  if(d){recognized=true;const bevel=name=>{const b=child(d,name);return b?{type:b.attrs.prst??'circle',w:n(b.attrs.w,76200)/EMU,h:n(b.attrs.h,76200)/EMU}:{type:'none'};};f.threeD={bevelTop:bevel('bevelT'),bevelBottom:bevel('bevelB'),depth:n(d.attrs.extrusionH)/EMU,contourWidth:n(d.attrs.contourW)/EMU,contourColor:color(child(d,'contourClr'),'#4472c4'),material:d.attrs.prstMaterial,lightRig:light?.attrs.rig,lightAngle:n(child(light,'rot')?.attrs.rev)/60000};}
  return recognized?normalizeChartAreaFormat(f,kind):undefined;
}
