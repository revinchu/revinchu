// Standard DrawingML groups + WIXEL editable diagram metadata. No native dgm claim.
import { smartArtParts, normalizeSmartArt, isSmartArt } from './smartart.js';
import { child, descendants, esc, parseXml } from './xml.js';
import { fromBase64 } from './vba.js';
import { pictureMediaSource, pictureBlipXml, pictureGeometryXml, pictureBorderXml, pictureEffectXml, pictureMetadataXml } from './picture-drawingml.js';
const EMU = 9525;
export const WIXEL_GROUP_URI = '{EE8A84AB-E019-456C-BBCB-374EA11B2A53}';
export function isDrawingGroup(shape) { return isSmartArt(shape) || shape?.kind === 'group' && Array.isArray(shape.groupItems); }
/** Click links are data only. Navigation and protocol checks belong to the UI. */
export function readDrawingHyperlink(nonVisual, rels) {
  const h = child(nonVisual, 'hlinkClick');
  if (!h) return null;
  const key = Object.keys(h.attrs).find(k => k === 'r:id' || k.endsWith(':id'));
  const rel = key && rels[h.attrs[key]];
  if (!rel || rel.type !== 'hyperlink') return null;
  const target = rel.rawTarget ?? rel.target;
  if (typeof target !== 'string' || !target) return null;
  const out = { target };
  if (rel.targetMode === 'External' || rel.external) out.targetMode = 'External';
  else if (rel.targetMode === 'Internal') out.targetMode = 'Internal';
  for (const k of ['tooltip', 'action']) if (h.attrs[k] !== undefined) out[k] = h.attrs[k];
  return out;
}
export function drawingHyperlinkXml(hyperlink, relationship) {
  if (typeof hyperlink?.target !== 'string' || !hyperlink.target) return '';
  const mode = ['External', 'Internal'].includes(hyperlink.targetMode) ? hyperlink.targetMode : undefined;
  const id = relationship(hyperlink.target, mode);
  return `<a:hlinkClick r:id="${esc(id)}"${hyperlink.tooltip !== undefined ? ` tooltip="${esc(hyperlink.tooltip)}"` : ''}${hyperlink.action !== undefined ? ` action="${esc(hyperlink.action)}"` : ''}/>`;
}
function drawingNonVisual(element) {
  const name = { sp: 'nvSpPr', cxnSp: 'nvCxnSpPr', pic: 'nvPicPr', grpSp: 'nvGrpSpPr' }[element.name];
  return child(child(element, name), 'cNvPr');
}
function hyperlinkSignature(link) {
  return link ? [link.target, link.targetMode === 'External', link.tooltip ?? '', link.action ?? ''] : null;
}
function metadata(shape) {
  return isSmartArt(shape) ? { kind: 'smartart', smartArt: normalizeSmartArt(shape.smartArt) } : { kind: 'group', groupItems: shape.groupItems, groupSize: shape.groupSize ?? { w: shape.w, h: shape.h } };
}
function checksum(data) { let hash = 2166136261; for (let i = 0; i < data.length; i++) hash = Math.imul(hash ^ (typeof data === 'string' ? data.charCodeAt(i) : data[i]), 16777619); return `${data.length}:${(hash >>> 0).toString(16)}`; }
// Excel may omit DrawingML defaults when saving. Ignore only equivalent defaults.
function canonical(node, imageFor) {
  if (!node) return null;
  const effectDefaults = node.name === 'reflection' ? {blurRad:'0',stPos:'0',endA:'0',fadeDir:'5400000',sx:'100000',rotWithShape:'1'}
    : node.name === 'outerShdw' ? {rotWithShape:'1'} : {};
  const attrs = Object.entries(node.attrs ?? {}).filter(([k, v]) => !k.startsWith('xmlns') && effectDefaults[k] !== v && !(node.name === 'path' && (k === 'stroke' && ['1','true'].includes(v) || k === 'fill' && v === 'norm')));
  if (node.name === 'blip') for (const entry of attrs) if (entry[0].endsWith(':embed') || entry[0] === 'embed') { entry[0] = 'imageChecksum'; entry[1] = imageFor(entry[1]); }
  return [node.name, attrs.sort((a,b) => a[0].localeCompare(b[0])), node.text ?? '', (node.children ?? []).filter(n => n.name !== 'extLst').map(n => canonical(n, imageFor))];
}
function drawingHidden(element) {
  const nvName = { sp: 'nvSpPr', cxnSp: 'nvCxnSpPr', pic: 'nvPicPr', grpSp: 'nvGrpSpPr' }[element.name];
  return ['1', 'true'].includes(child(child(element, nvName), 'cNvPr')?.attrs.hidden);
}
function nativeSignature(element, imageFor, linkFor = () => null) {
  const parts = element.children.filter(n => ['sp', 'cxnSp', 'pic', 'grpSp'].includes(n.name)).map(n => {
    const part = n.name === 'grpSp'
      ? ['group', canonical(child(n, 'grpSpPr'), imageFor), nativeSignature(n, imageFor, linkFor)]
      : [n.name, canonical(child(n, 'spPr'), imageFor), canonical(child(n, 'txBody'), imageFor), canonical(child(n, 'blipFill'), imageFor)];
    // Keep previous visible-object signatures; omitted/0/false all mean visible.
    if (drawingHidden(n)) part.push(['hidden', true]);
    const link = hyperlinkSignature(linkFor(drawingNonVisual(n)));
    if (link) part.push(['hyperlink', link]);
    return part;
  });
  return checksum(JSON.stringify(parts));
}
export function drawingGroupXml(shape, id, { shapeXml, xfrm, nextId, embedImage, embedSource = source => embedImage({src:source}), hyperlinkXml = () => '' }, depth = 0) {
  if (depth > 16) throw new Error('도형 그룹의 중첩 수준이 너무 깊습니다.');
  const items = isSmartArt(shape) ? smartArtParts(shape) : shape.groupItems;
  if (items.length > 1000) throw new Error('도형 그룹은 1,000개 이하로 저장하세요.');
  const base = isSmartArt(shape) ? { w: shape.w, h: shape.h } : shape.groupSize ?? shape;
  const images = new Map();
  const hyperlinks = {};
  const linkXml = object => {
    const xml = hyperlinkXml(object);
    if (xml) {
      const node = parseXml(xml), id = node.attrs['r:id'];
      hyperlinks[id] = { type: 'hyperlink', rawTarget: object.hyperlink.target, targetMode: object.hyperlink.targetMode };
    }
    return xml;
  };
  const registerImage = picture => { const rel = embedImage(picture), data = /^data:[^;,]+;base64,(.*)$/s.exec(pictureMediaSource(picture) ?? ''); if (rel && data) images.set(rel, checksum(fromBase64(data[1]))); return rel; };
  const transform = `<a:xfrm${shape.rot ? ` rot="${Math.round(shape.rot * 60000)}"` : ''}${shape.flip ? ' flipH="1"' : ''}${shape.flipV ? ' flipV="1"' : ''}><a:off x="${Math.round(shape.x * EMU)}" y="${Math.round(shape.y * EMU)}"/><a:ext cx="${Math.round(shape.w * EMU)}" cy="${Math.round(shape.h * EMU)}"/><a:chOff x="0" y="0"/><a:chExt cx="${Math.round(base.w * EMU)}" cy="${Math.round(base.h * EMU)}"/></a:xfrm>`;
  const content = items.map(p => {
    const childId = nextId();
    if (isDrawingGroup(p)) return drawingGroupXml(p, childId, { shapeXml, xfrm, nextId, embedImage: registerImage, embedSource, hyperlinkXml: linkXml }, depth + 1);
    if (p.kind !== 'picture') return shapeXml(p, childId, xfrm, linkXml);
    const rel = registerImage(p);
    if (!rel) throw new Error('그룹 안의 그림을 XLSX에 포함할 수 없습니다. PNG/JPEG 그림을 사용하세요.');
    const metadata = pictureMetadataXml(p, embedSource);
    return `<xdr:pic><xdr:nvPicPr><xdr:cNvPr id="${childId}" name="${esc(p.name ?? '그림')}"${p.hidden ? ' hidden="1"' : ''}${p.alt !== undefined ? ` descr="${esc(p.alt)}"` : ''}>${linkXml(p)}${metadata ? `<a:extLst>${metadata}</a:extLst>` : ''}</xdr:cNvPr><xdr:cNvPicPr><a:picLocks noChangeAspect="${p.lockAspect === false ? 0 : 1}"/></xdr:cNvPicPr></xdr:nvPicPr><xdr:blipFill><a:blip r:embed="${rel}">${pictureBlipXml(p)}</a:blip>${p.crop ? `<a:srcRect${['l','t','r','b'].map(k => ` ${k}="${Math.round((p.crop[k] ?? 0) * 100000)}"`).join('')}/>` : ''}<a:stretch><a:fillRect/></a:stretch></xdr:blipFill><xdr:spPr>${xfrm(p)}${pictureGeometryXml(p)}${pictureBorderXml(p)}${pictureEffectXml(p)}</xdr:spPr></xdr:pic>`;
  }).join('');
  const signature = nativeSignature(parseXml(`<group>${content}</group>`), rel => images.get(rel) ?? 'missing', nv => readDrawingHyperlink(nv, hyperlinks));
  const props = `<a:extLst><a:ext uri="${WIXEL_GROUP_URI}"><wx:group xmlns:wx="https://wixel.app/drawing/group/1" signature="${signature}" json="${esc(JSON.stringify(metadata(shape), (key, value) => key === 'effectPng' ? undefined : value))}"/></a:ext></a:extLst>`;
  return `<xdr:grpSp><xdr:nvGrpSpPr><xdr:cNvPr id="${id}" name="${esc(shape.name || (isSmartArt(shape) ? 'SmartArt' : '그룹'))}"${shape.hidden ? ' hidden="1"' : ''}>${linkXml(shape)}${props}</xdr:cNvPr><xdr:cNvGrpSpPr/></xdr:nvGrpSpPr><xdr:grpSpPr>${transform}</xdr:grpSpPr>${content}</xdr:grpSp>`;
}
export function readDrawingGroup(element, box, id, { files = {}, rels = {} } = {}) {
  const nv = child(child(element, 'nvGrpSpPr'), 'cNvPr'), data = descendants(nv, 'group').find(x => x.attrs['xmlns:wx'] === 'https://wixel.app/drawing/group/1');
  if (!data?.attrs.json || data.attrs.json.length > 20000000) return null;
  try {
    if (!data.attrs.signature || data.attrs.signature !== nativeSignature(element, rel => { const bytes = files[rels[rel]?.target]; return bytes ? checksum(bytes) : 'missing'; }, nv => readDrawingHyperlink(nv, rels))) return null;
    const m = JSON.parse(data.attrs.json);
    if (m.kind === 'smartart') m.smartArt = normalizeSmartArt(m.smartArt);
    else if (m.kind !== 'group' || !Array.isArray(m.groupItems) || m.groupItems.length > 1000 || !(m.groupSize?.w > 0 && m.groupSize?.h > 0)) return null;
    const xf = child(child(element, 'grpSpPr'), 'xfrm');
    const hyperlink = readDrawingHyperlink(nv, rels);
    return { id, name: nv?.attrs.name || 'SmartArt', ...m, ...box, ...(hyperlink ? { hyperlink } : {}), ...(drawingHidden(element) ? { hidden: true } : {}), ...(Number(xf?.attrs.rot) ? { rot: Number(xf.attrs.rot) / 60000 } : {}), ...(xf?.attrs.flipH === '1' ? { flip: true } : {}), ...(xf?.attrs.flipV === '1' ? { flipV: true } : {}) };
  } catch { return null; }
}
