// SVG path grammar -> literal DrawingML-compatible commands. No DOM/eval.
// https://www.w3.org/TR/SVG2/paths.html#PathData
const NUMBER = /^[+-]?(?:\d+\.?\d*|\.\d+)(?:e[+-]?\d+)?/i;
export const GEOMETRY_LIMIT = 2000;
const error = () => new Error('SVG 경로가 올바르지 않거나 지원 범위를 벗어났습니다.');
export function parseSvgPath(source) {
  if (typeof source !== 'string' || source.length > 500000) throw error();
  let at=0, op='', x=0,y=0,sx=0,sy=0,prev='',ctrl=null; const out=[];
  const skip=()=>{while(/[\s,]/.test(source[at]??'') && at<source.length)at++;};
  const number=(flag=false)=>{skip();const m=NUMBER.exec(source.slice(at));if(!m)throw error();
    if(flag){if(source[at]!=='0'&&source[at]!=='1')throw error();return Number(source[at++]);}
    at+=m[0].length;const n=Number(m[0]);if(!Number.isFinite(n)||Math.abs(n)>1e8)throw error();return n;};
  const add=c=>{if(out.length>=10000)throw error();out.push(c);};
  while(at<source.length){skip();if(at>=source.length)break;
    if(/[a-z]/i.test(source[at]))op=source[at++];else if(!op)throw error();
    const k=op.toUpperCase(),rel=op!==k,ox=rel?x:0,oy=rel?y:0;
    if(!'MLHVCSQTAZ'.includes(k)||(!out.length&&k!=='M'))throw error();
    if(k==='Z'){add(['Z']);x=sx;y=sy;prev=k;ctrl=null;op='';continue;}
    const point=()=>[number()+ox,number()+oy];let c;
    if(k==='M'||k==='L'){const p=point();c=[k,...p];x=p[0];y=p[1];if(k==='M'){sx=x;sy=y;op=rel?'l':'L';}}
    else if(k==='H'){x=number()+ox;c=['L',x,y];}
    else if(k==='V'){y=number()+oy;c=['L',x,y];}
    else if(k==='C'){const a=point(),b=point(),p=point();c=['C',...a,...b,...p];ctrl=b;x=p[0];y=p[1];}
    else if(k==='S'){const a=(prev==='C'||prev==='S')&&ctrl?[2*x-ctrl[0],2*y-ctrl[1]]:[x,y],b=point(),p=point();c=['C',...a,...b,...p];ctrl=b;x=p[0];y=p[1];}
    else if(k==='Q'||k==='T'){const a=k==='Q'?point():(prev==='Q'||prev==='T')&&ctrl?[2*x-ctrl[0],2*y-ctrl[1]]:[x,y],p=point();c=['Q',...a,...p];ctrl=a;x=p[0];y=p[1];}
    else if(k==='A'){const rx=number(),ry=number(),angle=number(),large=number(true),sweep=number(true),p=point();
      for(const a of arcCubics(x,y,rx,ry,angle,large,sweep,p[0],p[1]))add(a);x=p[0];y=p[1];}
    if(c)add(c);if(!'CSQT'.includes(k))ctrl=null;prev=k;
  }
  if(!out.length)throw error();return out;
}
function arcCubics(x,y,rx,ry,angle,large,sweep,nx,ny){
  if(x===nx&&y===ny)return [];rx=Math.abs(rx);ry=Math.abs(ry);if(!rx||!ry)return [['L',nx,ny]];
  const phi=angle*Math.PI/180,co=Math.cos(phi),sn=Math.sin(phi),dx=(x-nx)/2,dy=(y-ny)/2,px=co*dx+sn*dy,py=-sn*dx+co*dy;
  const scale=Math.sqrt(px*px/(rx*rx)+py*py/(ry*ry));if(scale>1){rx*=scale;ry*=scale;}
  const s=(large===sweep?-1:1)*Math.sqrt(Math.max(0,(rx*rx*ry*ry-rx*rx*py*py-ry*ry*px*px)/(rx*rx*py*py+ry*ry*px*px)));
  const cxp=s*rx*py/ry,cyp=-s*ry*px/rx,cx=co*cxp-sn*cyp+(x+nx)/2,cy=sn*cxp+co*cyp+(y+ny)/2;
  let a=Math.atan2((py-cyp)/ry,(px-cxp)/rx),delta=Math.atan2((-py-cyp)/ry,(-px-cxp)/rx)-a;
  if(sweep&&delta<0)delta+=Math.PI*2;if(!sweep&&delta>0)delta-=Math.PI*2;
  const n=Math.ceil(Math.abs(delta)/(Math.PI/4)),step=delta/n,out=[];
  const p=t=>[cx+rx*co*Math.cos(t)-ry*sn*Math.sin(t),cy+rx*sn*Math.cos(t)+ry*co*Math.sin(t)];
  const d=t=>[-rx*co*Math.sin(t)-ry*sn*Math.cos(t),-rx*sn*Math.sin(t)+ry*co*Math.cos(t)];
  for(let i=0;i<n;i++){const b=a+step,f=4/3*Math.tan(step/4),pa=p(a),pb=p(b),da=d(a),db=d(b);out.push(['C',pa[0]+f*da[0],pa[1]+f*da[1],pb[0]-f*db[0],pb[1]-f*db[1],...pb]);a=b;}
  if(out.length){out.at(-1)[5]=nx;out.at(-1)[6]=ny;}return out;
}
export function transformCommands(commands,m){return commands.map(c=>{const o=[c[0]];for(let i=1;i<c.length;i+=2)o.push(m[0]*c[i]+m[2]*c[i+1]+m[4],m[1]*c[i]+m[3]*c[i+1]+m[5]);return o;});}
export function multiplyMatrix(a,b){return [a[0]*b[0]+a[2]*b[1],a[1]*b[0]+a[3]*b[1],a[0]*b[2]+a[2]*b[3],a[1]*b[2]+a[3]*b[3],a[0]*b[4]+a[2]*b[5]+a[4],a[1]*b[4]+a[3]*b[5]+a[5]];}
export function parseSvgTransform(text=''){
  let m=[1,0,0,1,0,0],end=0;const re=/([a-zA-Z]+)\s*\(([^)]*)\)/g;let hit;
  while((hit=re.exec(text))){if(!/^[\s,]*$/.test(text.slice(end,hit.index)))throw error();end=re.lastIndex;
    const a=hit[2].trim().split(/[\s,]+/).map(Number);if(!a.every(Number.isFinite))throw error();let n;
    if(hit[1]==='matrix'&&a.length===6)n=a;
    else if(hit[1]==='translate'&&(a.length===1||a.length===2))n=[1,0,0,1,a[0],a[1]??0];
    else if(hit[1]==='scale'&&(a.length===1||a.length===2))n=[a[0],0,0,a[1]??a[0],0,0];
    else if(hit[1]==='rotate'&&(a.length===1||a.length===3)){const co=Math.cos(a[0]*Math.PI/180),sn=Math.sin(a[0]*Math.PI/180),x=a[1]??0,y=a[2]??0;n=[co,sn,-sn,co,x-co*x+sn*y,y-sn*x-co*y];}
    else if((hit[1]==='skewX'||hit[1]==='skewY')&&a.length===1){const t=Math.tan(a[0]*Math.PI/180);n=hit[1]==='skewX'?[1,0,t,1,0,0]:[1,t,0,1,0,0];}
    else throw error();m=multiplyMatrix(m,n);
  }
  if(!/^[\s,]*$/.test(text.slice(end))||!m.every(n=>Number.isFinite(n)&&Math.abs(n)<1e8))throw error();return m;
}
const mix=(a,b)=>[(a[0]+b[0])/2,(a[1]+b[1])/2];
export function flattenCommands(commands,tolerance=.2,limit=GEOMETRY_LIMIT){
  const rings=[];let ring=null,point=null,count=0;
  const add=p=>{if(++count>limit)throw new Error('도형 경로가 너무 복잡합니다. 점 수를 줄인 뒤 다시 시도하세요.');ring.push(p);};
  const distance=(p,a,b)=>{const x=b[0]-a[0],y=b[1]-a[1],d=x*x+y*y,t=d?Math.max(0,Math.min(1,((p[0]-a[0])*x+(p[1]-a[1])*y)/d)):0;return Math.hypot(p[0]-a[0]-t*x,p[1]-a[1]-t*y);};
  const curve=(a,b,c,d,depth)=>{if(Math.max(distance(b,a,d),distance(c,a,d))<=tolerance){add(d);return;}
    if(depth>=20)throw new Error('곡선의 정밀도를 유지할 수 없어 변환하지 않았습니다.');
    const ab=mix(a,b),bc=mix(b,c),cd=mix(c,d),abc=mix(ab,bc),bcd=mix(bc,cd),mid=mix(abc,bcd);curve(a,ab,abc,mid,depth+1);curve(mid,bcd,cd,d,depth+1);};
  for(const c of commands){
    if(c[0]==='M'){ring=[];rings.push(ring);point=c.slice(1);add(point);}
    else if(!ring)throw error();
    else if(c[0]==='L'){point=c.slice(1);add(point);}
    else if(c[0]==='C'){curve(point,c.slice(1,3),c.slice(3,5),c.slice(5),0);point=c.slice(5);}
    else if(c[0]==='Q'){const q=c.slice(1,3),d=c.slice(3);curve(point,[point[0]+2/3*(q[0]-point[0]),point[1]+2/3*(q[1]-point[1])],[d[0]+2/3*(q[0]-d[0]),d[1]+2/3*(q[1]-d[1])],d,0);point=d;}
    else if(c[0]==='Z'){ring.closed=true;point=ring[0];}
    else throw error();
  }return rings;
}
