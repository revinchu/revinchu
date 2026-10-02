// Planar arrangement of flattened filled paths. Original models are never edited.
// Curves are approximated only for the new result (default tolerance 0.2 sheet px).
import { GEOM } from './shapes.js';
import { validShapePath } from './shape-path.js';
import { parseSvgPath, transformCommands, flattenCommands, GEOMETRY_LIMIT } from './geometry-path.js';

const cross=(a,b)=>a[0]*b[1]-a[1]*b[0];
const sub=(a,b)=>[a[0]-b[0],a[1]-b[1]];
const area=ring=>ring.reduce((s,p,i)=>s+cross(sub(p,ring[0]),sub(ring[(i+1)%ring.length],ring[0])),0)/2;
const tooComplex=()=>new Error('도형 조합이 너무 복잡합니다. 도형이나 점 수를 줄인 뒤 다시 시도하세요.');
export function ringsContain(rings,p,evenodd=false){
  let winding=0;for(const ring of rings)for(let i=0;i<ring.length;i++){
    const a=ring[i],b=ring[(i+1)%ring.length],side=cross(sub(b,a),sub(p,a));
    if(a[1]<=p[1]){if(b[1]>p[1]&&side>0)winding++;}else if(b[1]<=p[1]&&side<0)winding--;
  }return evenodd?Math.abs(winding)%2===1:winding!==0;
}
function shapeRegion(shape,tolerance){
  if(!shape||shape.kind==='group'||shape.kind==='picture'||shape.kind==='smartArt'||shape.customGeometry)throw new Error('그룹·그림 또는 해석할 수 없는 도형은 먼저 도형으로 변환하거나 그룹을 해제하세요.');
  const {x=0,y=0,w,h}=shape;if(![x,y,w,h,shape.rot??0].every(Number.isFinite)||w<=0||h<=0)throw new Error('도형 크기와 좌표가 올바르지 않습니다.');
  const angle=(shape.rot??0)*Math.PI/180,co=Math.cos(angle),sn=Math.sin(angle),fx=shape.flip?-1:1,fy=shape.flipV?-1:1;
  const m=[co*fx,sn*fx,-sn*fy,co*fy,x+w/2-co*fx*w/2+sn*fy*h/2,y+h/2-sn*fx*w/2-co*fy*h/2];
  let paths;
  if(shape.path){if(!validShapePath(shape.path))throw new Error('도형 경로를 해석할 수 없습니다.');
    paths=shape.path.paths.map(p=>({commands:transformCommands(p.commands,[w,0,0,h,0,0]),fill:p.fill!==false,evenodd:!!p.evenodd}));
  }else {const geom=GEOM[shape.kind];if(!geom)throw new Error('이 도형 종류는 조합을 지원하지 않습니다.');
    paths=geom(w,h).map(p=>({commands:parseSvgPath(p.d),fill:!p.line,evenodd:!!p.evenodd}));}
  const parts=paths.filter(p=>p.fill).map(p=>({rings:flattenCommands(transformCommands(p.commands,m),tolerance),evenodd:p.evenodd}));
  if(!parts.length)throw new Error('선과 열린 자유곡선은 면적 도형으로 조합할 수 없습니다. 닫힌 도형을 선택하세요.');
  return parts;
}

/** Filled path regions -> oriented, nonoverlapping rings grouped by membership.
 * A region contains parts; each part uses SVG nonzero/evenodd fill semantics. */
export function arrangeRegions(regions,operation='union'){
  if(!regions.length||regions.length>16)throw new Error('도형은 1개 이상 16개 이하로 처리할 수 있습니다.');
  if(!['union','combine','fragment','intersect','subtract'].includes(operation))throw new Error('지원하지 않는 도형 조합입니다.');
  const segments=[];let magnitude=1,minX=Infinity,minY=Infinity,maxX=-Infinity,maxY=-Infinity;
  for(const parts of regions)for(const part of parts)for(const ring of part.rings){
    for(const p of ring){if(!p.every(n=>Number.isFinite(n)&&Math.abs(n)<=1e8))throw new Error('도형 좌표가 올바르지 않습니다.');magnitude=Math.max(magnitude,Math.abs(p[0]),Math.abs(p[1]));minX=Math.min(minX,p[0]);minY=Math.min(minY,p[1]);maxX=Math.max(maxX,p[0]);maxY=Math.max(maxY,p[1]);}
    for(let i=0;i<ring.length;i++){const a=ring[i],b=ring[(i+1)%ring.length];if(Math.hypot(b[0]-a[0],b[1]-a[1])>1e-10)segments.push({a,b,t:[0,1]});}
  }
  if(segments.length>GEOMETRY_LIMIT)throw tooComplex();
  const eps=Math.max(1e-8,(maxX-minX)*1e-10,(maxY-minY)*1e-10,magnitude*Number.EPSILON*16),all=(1<<regions.length)-1;
  let splitEvents=0;
  const insert=(s,t)=>{if(t>1e-10&&t<1-1e-10){if(++splitEvents>24000)throw tooComplex();s.t.push(t);}};
  for(let i=0;i<segments.length;i++)for(let j=i+1;j<segments.length;j++){
    const a=segments[i],b=segments[j];
    if(Math.max(a.a[0],a.b[0])+eps<Math.min(b.a[0],b.b[0])||Math.max(b.a[0],b.b[0])+eps<Math.min(a.a[0],a.b[0])||Math.max(a.a[1],a.b[1])+eps<Math.min(b.a[1],b.b[1])||Math.max(b.a[1],b.b[1])+eps<Math.min(a.a[1],a.b[1]))continue;
    const r=sub(a.b,a.a),s=sub(b.b,b.a),q=sub(b.a,a.a),d=cross(r,s),len=Math.hypot(...r),len2=Math.hypot(...s);
    if(Math.abs(d)>eps*Math.max(len,len2)){
      const t=cross(q,s)/d,u=cross(q,r)/d;if(t>=-1e-10&&t<=1+1e-10&&u>=-1e-10&&u<=1+1e-10){insert(a,t);insert(b,u);}
    }else if(Math.abs(cross(q,r))<=eps*len){
      const project=(p,s)=>{const r=sub(s.b,s.a),q=sub(p,s.a);return (q[0]*r[0]+q[1]*r[1])/(r[0]*r[0]+r[1]*r[1]);};
      insert(a,project(b.a,a));insert(a,project(b.b,a));insert(b,project(a.a,b));insert(b,project(a.b,b));
    }
  }
  const vertices=new Map(),atoms=new Map();
  const vertex=p=>{const key=Math.round(p[0]/eps)+','+Math.round(p[1]/eps);if(!vertices.has(key))vertices.set(key,{key,p});return vertices.get(key);};
  for(const s of segments){s.t.sort((a,b)=>a-b);for(let i=1;i<s.t.length;i++){
    const lo=s.t[i-1],hi=s.t[i];if(hi-lo<1e-10)continue;
    const at=t=>[s.a[0]+(s.b[0]-s.a[0])*t,s.a[1]+(s.b[1]-s.a[1])*t],a=vertex(at(lo)),b=vertex(at(hi));if(a===b)continue;
    const key=a.key<b.key?a.key+';'+b.key:b.key+';'+a.key;if(!atoms.has(key))atoms.set(key,{a,b});if(atoms.size>12000)throw tooComplex();
  }}
  if(atoms.size*segments.length*2>20000000)throw tooComplex();
  const membership=p=>{let mask=0;for(let i=0;i<regions.length;i++)if(regions[i].some(part=>ringsContain(part.rings,p,part.evenodd)))mask|=1<<i;return mask;};
  const selected=m=>operation==='union'?m!==0:operation==='intersect'?m===all:operation==='subtract'?(m&1)!==0&&(m&~1)===0:popcount(m)%2===1;
  const groups=new Map();
  const edge=(key,a,b)=>{if(!groups.has(key))groups.set(key,[]);groups.get(key).push({a,b});};
  for(const {a,b} of atoms.values()){
    const dx=b.p[0]-a.p[0],dy=b.p[1]-a.p[1],length=Math.hypot(dx,dy),offset=Math.min(length*1e-5,eps*8),x=(a.p[0]+b.p[0])/2,y=(a.p[1]+b.p[1])/2;
    const left=membership([x-dy/length*offset,y+dx/length*offset]),right=membership([x+dy/length*offset,y-dx/length*offset]);
    if(operation==='fragment'){if(left!==right){if(left)edge(left,a,b);if(right)edge(right,b,a);}}
    else if(selected(left)!==selected(right))selected(left)?edge(1,a,b):edge(1,b,a);
  }
  if(groups.size>256)throw tooComplex();
  return [...groups].map(([mask,edges])=>({mask,rings:stitch(edges,eps)}));
}
function popcount(n){let count=0;while(n){n&=n-1;count++;}return count;}
function stitch(edges,eps){
  const outgoing=new Map();for(const e of edges){if(!outgoing.has(e.a.key))outgoing.set(e.a.key,[]);outgoing.get(e.a.key).push(e);}
  const rings=[];
  for(const first of edges){if(first.used)continue;let e=first;const ring=[];
    while(!e.used){e.used=true;ring.push(e.a.p);if(e.b===first.a)break;
      const options=(outgoing.get(e.b.key)??[]).filter(v=>!v.used);if(!options.length)throw new Error('도형 경계가 너무 가깝거나 교차가 복잡해 안전하게 조합하지 못했습니다.');
      const reverse=Math.atan2(e.a.p[1]-e.b.p[1],e.a.p[0]-e.b.p[0]);
      const turn=v=>(reverse-Math.atan2(v.b.p[1]-v.a.p[1],v.b.p[0]-v.a.p[0])+Math.PI*4)%(Math.PI*2);
      options.sort((a,b)=>turn(a)-turn(b));e=options[0];if(ring.length>12000)throw tooComplex();
    }
    if(e.b!==first.a)throw new Error('닫힌 도형 경계를 만들 수 없습니다. 원본은 변경하지 않았습니다.');
    const clean=ring.filter((p,i)=>{const a=ring[(i+ring.length-1)%ring.length],b=ring[(i+1)%ring.length];return Math.abs(cross(sub(p,a),sub(b,p)))>eps*(Math.hypot(...sub(p,a))+Math.hypot(...sub(b,p)));});
    if(clean.length>=3&&Math.abs(area(clean))>eps*eps)rings.push(clean);
  }return rings;
}
export function geometryShape(rings,template={},id=''){
  let x=Infinity,y=Infinity,right=-Infinity,bottom=-Infinity;for(const r of rings)for(const p of r){x=Math.min(x,p[0]);y=Math.min(y,p[1]);right=Math.max(right,p[0]);bottom=Math.max(bottom,p[1]);}
  if(!Number.isFinite(x)||right-x<1e-8||bottom-y<1e-8)throw new Error('면적이 없는 도형은 만들 수 없습니다.');
  const w=Math.max(1,right-x),h=Math.max(1,bottom-y),commands=[];for(const ring of rings){commands.push(['M',(ring[0][0]-x)/w,(ring[0][1]-y)/h]);for(const p of ring.slice(1))commands.push(['L',(p[0]-x)/w,(p[1]-y)/h]);commands.push(['Z']);}
  if(commands.length>10000)throw tooComplex();
  const result={...structuredClone(template),id:id||'shg'+Date.now().toString(36)+Math.random().toString(36).slice(2,8),kind:'freeform',x,y,w,h,path:{paths:[{commands,fill:true,stroke:true}]}};
  for(const k of ['rot','flip','flipV','customGeometry','groupItems','groupSize','icon','png','src','anchor'])delete result[k];return result;
}
export function geometryComponents(rings){
  const outers=rings.filter(r=>area(r)>0).map(r=>[r]);
  for(const hole of rings.filter(r=>area(r)<0)){
    const targets=outers.filter(rs=>ringsContain([rs[0]],hole[0])).sort((a,b)=>area(a[0])-area(b[0]));
    if(!targets.length)throw new Error('도형 안쪽 빈 영역을 보존할 수 없어 조합하지 않았습니다.');targets[0].push(hole);
  }return outers;
}
export function mergeShapes(shapes,operation,options={}){
  if(!Array.isArray(shapes)||shapes.length<2||shapes.length>16)throw new Error('조합할 도형을 2개 이상 16개 이하로 선택하세요.');
  const tolerance=options.tolerance??.2;if(!Number.isFinite(tolerance)||tolerance<.01||tolerance>1)throw new Error('곡선 정밀도는 0.01~1 픽셀이어야 합니다.');
  const regions=shapes.map(sh=>shapeRegion(sh,tolerance)),arrangement=arrangeRegions(regions,operation),template=shapes[0];
  const pieces=operation==='fragment'?arrangement.flatMap(g=>geometryComponents(g.rings)):arrangement.map(g=>g.rings);
  if(pieces.length>256)throw tooComplex();
  return pieces.filter(r=>r.length).map((rings,i)=>geometryShape(rings,template,options.idPrefix?options.idPrefix+i:''));
}
