// SVG 3차원 차트 기하. 값 축은 앞면 기준이며 깊이 면은 같은 값의 투영입니다.
const bound = (v, lo, hi, fallback) => Number.isFinite(v) ? Math.max(lo, Math.min(hi, v)) : fallback;
const num = (v) => Number(v.toFixed(2));
const pt = (x, y) => `${num(x)},${num(y)}`;
const safe = (v) => String(v).replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&apos;' })[c]);
export function chartView3D(chart) {
  const v = chart.view3D ?? {};
  return { rotX: bound(v.rotX, -90, 90, chart.type === 'pie' ? 35 : 20), rotY: bound(v.rotY, 0, 360, chart.type === 'pie' ? 0 : 30), depthPercent: bound(v.depthPercent, 20, 2000, 100), perspective: bound(v.perspective, 0, 240, 30), rAngAx: v.rAngAx !== false };
}
export function chartDepth(chart, width, height) {
  const v = chartView3D(chart);
  const depth = Math.min(width, height) * Math.min(0.25, 0.1 * v.depthPercent / 100) / (v.rAngAx ? 1 : 1 + v.perspective / 240);
  return { dx: num(Math.sin(v.rotY * Math.PI / 180) * depth), dy: num(-Math.sin(v.rotX * Math.PI / 180) * depth) };
}
export function extrudedPolygon(points, dx, dy, fill, attrs = '') {
  if (points.length < 2) return '';
  const f = safe(fill);
  let out = '<g data-3d="solid">';
  for (let i = 0; i < points.length; i++) {
    const a = points[i], b = points[(i + 1) % points.length];
    // Only outward faces facing the viewer are visible (screen clockwise front).
    if ((b[0] - a[0]) * dy - (b[1] - a[1]) * dx >= 0) continue;
    const shape = `${pt(...a)} ${pt(...b)} ${pt(b[0] + dx, b[1] + dy)} ${pt(a[0] + dx, a[1] + dy)}`;
    out += `<polygon points="${shape}" fill="${f}"${attrs}/><polygon points="${shape}" fill="#000" fill-opacity="0.24" pointer-events="none"/>`;
  }
  out += `<polygon points="${points.map((p) => pt(...p)).join(' ')}" fill="${f}"${attrs}/></g>`;
  return out;
}
export function chartWalls3D(area, offset, fill = '#f4f6f8') {
  const { x, y, w, h } = area, { dx, dy } = offset;
  return `<g data-3d="walls" stroke="#cbd5e1" stroke-width="0.6"><rect x="${num(x + dx)}" y="${num(y + dy)}" width="${num(w)}" height="${num(h)}" fill="${safe(fill)}"/><polygon points="${pt(x, y + h)} ${pt(x + w, y + h)} ${pt(x + w + dx, y + h + dy)} ${pt(x + dx, y + h + dy)}" fill="#e2e8f0"/><line x1="${num(x)}" y1="${num(y)}" x2="${num(x + dx)}" y2="${num(y + dy)}"/></g>`;
}
export function pieProjection3D(chart, radius) {
  const v = chartView3D(chart);
  return { squash: Math.max(0.2, Math.abs(Math.sin(v.rotX * Math.PI / 180))), depth: Math.max(2, radius * Math.min(0.5, 0.2 * v.depthPercent / 100) / (v.rAngAx ? 1 : 1 + v.perspective / 240)), rotation: v.rotY * Math.PI / 180 };
}
export function pieSolid3D(slices, cx, cy, radius, projection) {
  const { squash, depth } = projection, ry = radius * squash;
  const sides = [], tops = [];
  for (const s of slices) {
    const x = cx + s.ox, y = cy + s.oy;
    const xy = (a, z = 0) => pt(x + Math.cos(a) * radius, y + Math.sin(a) * ry + z);
    const color = safe(s.color), attrs = s.attrs ?? '';
    const lo = Math.floor(s.a / Math.PI), hi = Math.ceil(s.b / Math.PI);
    for (let k = lo; k < hi; k++) {
      const a = Math.max(s.a, k * Math.PI), b = Math.min(s.b, (k + 1) * Math.PI);
      if (b - a < 1e-8 || Math.sin((a + b) / 2) <= 0) continue;
      const path = `M${xy(a)}A${num(radius)},${num(ry)} 0 0 1 ${xy(b)}L${xy(b, depth)}A${num(radius)},${num(ry)} 0 0 0 ${xy(a, depth)}Z`;
      sides.push(`<path d="${path}" fill="${color}"${attrs}/><path d="${path}" fill="#000" fill-opacity="0.25" pointer-events="none"/>`);
    }
    if (s.ox || s.oy) for (const a of [s.a, s.b]) {
      const path = `M${pt(x, y)}L${xy(a)}L${xy(a, depth)}L${pt(x, y + depth)}Z`;
      sides.push(`<path d="${path}" fill="${color}"${attrs}/><path d="${path}" fill="#000" fill-opacity="0.18" pointer-events="none"/>`);
    }
    if (s.b - s.a >= Math.PI * 2 - 1e-7) tops.push(`<ellipse cx="${num(x)}" cy="${num(y)}" rx="${num(radius)}" ry="${num(ry)}" fill="${color}"${attrs}/>`);
    else tops.push(`<path d="M${pt(x, y)}L${xy(s.a)}A${num(radius)},${num(ry)} 0 ${s.b - s.a > Math.PI ? 1 : 0} 1 ${xy(s.b)}Z" fill="${color}" stroke="#fff" stroke-width="0.8"${attrs}/>`);
  }
  return `<g data-3d="pie">${sides.join('')}${tops.join('')}</g>`;
}

/** 막대 값축 끝점은 그대로 두고 단면을 원/삼각뿔로 투영한다. */
export function barSolid3D({ x, y, w, h, from, to, horizontal = false, dx, dy, shape = 'box', fill, attrs = '' }) {
  if (!['cylinder', 'cone', 'pyramid'].includes(shape)) return extrudedPolygon([[x,y],[x+w,y],[x+w,y+h],[x,y+h]],dx,dy,fill,attrs);
  const f=safe(fill), circular=shape!=='pyramid', count=circular?32:4;
  const section=(a,value,r=1)=> {
    const c=Math.cos(a),s=Math.sin(a);
    if(horizontal)return [value+dx/2+s*dx*r/2,y+h/2+c*h*r/2+s*dy*r/2];
    return [x+w/2+c*w*r/2+s*dx*r/2,value+dy/2+s*dy*r/2];
  };
  const angle=i=>circular?i*Math.PI*2/count:Math.PI/4+i*Math.PI/2;
  // 사각 단면은 타원의 내접이 아니라 원래 폭과 깊이를 유지한다.
  const point=(i,value,r=1)=> {
    if(circular)return section(angle(i),value,r);
    const q=[[0,0],[1,0],[1,1],[0,1]][(i+count)%count],a=(q[0]-.5)*r,b=(q[1]-.5)*r;
    return horizontal?[value+dx/2+b*dx,y+h/2+a*h+b*dy]:[x+w/2+a*w+b*dx,value+dy/2+b*dy];
  };
  const cap=(value,r)=>Array.from({length:count},(_,i)=>point(i,value,r));
  const endRadius=shape==='cylinder'?1:0,base=cap(from,1),end=cap(to,endRadius);
  let out=`<g data-3d="${shape}"><polygon points="${base.map(p=>pt(...p)).join(' ')}" fill="${f}"${attrs}/>`;
  const faces=Array.from({length:count},(_,i)=>({i,depth:Math.sin(angle(i)+Math.PI/count)})).sort((a,b)=>b.depth-a.depth);
  for(const {i} of faces){const next=(i+1)%count,points=[base[i],base[next],end[next],end[i]].map(p=>pt(...p)).join(' '),shade=.06+.28*(1+Math.cos(angle(i)))/2;out+=`<polygon points="${points}" fill="${f}"${attrs}/><polygon points="${points}" fill="#000" fill-opacity="${num(shade)}" pointer-events="none"/>`;}
  if(endRadius)out+=`<polygon points="${end.map(p=>pt(...p)).join(' ')}" fill="${f}"${attrs}/><polygon points="${end.map(p=>pt(...p)).join(' ')}" fill="#fff" fill-opacity="0.18" pointer-events="none"/>`;
  return out+'</g>';
}
