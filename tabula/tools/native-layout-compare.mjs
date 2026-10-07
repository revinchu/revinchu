// 비공개 Native Excel JSON과 WIXEL 모델 표본 비교. 실제 화면·조건부 서식 평가 합격이 아닙니다.
import { readFile, writeFile } from 'node:fs/promises';
import { resolve, dirname } from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';
const TOOL=fileURLToPath(import.meta.url),SOURCE=resolve(dirname(TOOL),'../src');
export const POINT_TO_CSS_PX=96/72;
const SIDES={left:'bl',top:'bt',right:'br',bottom:'bb',diagonalDown:'dd',diagonalUp:'du'};
const BORDER_STYLE={thin:[1,2],hair:[1,1],medium:[1,-4138],thick:[1,4],dotted:[-4118,2],dashed:[-4115,2],dashDot:[4,2],dashDotDot:[5,2],double:[-4119,2],mediumDashed:[-4115,-4138],mediumDashDot:[4,-4138],mediumDashDotDot:[5,-4138],slantDashDot:[13,-4138]};
const HORIZONTAL={1:'general','-4131':'left','-4108':'center','-4152':'right',7:'centerContinuous','-4130':'justify','-4117':'distributed',5:'fill'};
const VERTICAL={'-4160':'top','-4108':'middle','-4107':'bottom','-4130':'justify','-4117':'distributed'};
const DASH={1:'solid',2:'dot',3:'dot',4:'dash',5:'dashDot',6:'dashDotDot',7:'lgDash',8:'lgDashDot'};
const COMPOUND={1:'sng',2:'dbl',3:'thinThick',4:'thickThin',5:'tri'};
const color=value=>/^#[0-9a-f]{6}$/i.test(value??'')?value.toLowerCase():null;
export function oleColor(value) {
  if(!Number.isInteger(value)||value<0||value>0xffffff)return null;
  return '#'+[value&255,(value>>>8)&255,(value>>>16)&255].map(n=>n.toString(16).padStart(2,'0')).join('');
}
const formatCode=value=>/^(?:general|g\/표준)$/i.test(String(value??'').trim())?'General':String(value??'General');
function colIndex(name) {let c=0;for(const ch of name)c=c*26+ch.charCodeAt(0)-64;return c-1;}
function columnName(c) {let out='';for(c++;c>0;c=Math.floor((c-1)/26))out=String.fromCharCode(65+(c-1)%26)+out;return out;}
function cellAddress(r,c) {return columnName(c)+(r+1);}
function parseAddress(address) {const m=/^([A-Z]{1,3})([1-9][0-9]{0,6})$/.exec(address);if(!m||colIndex(m[1])>=16384||+m[2]>1048576)throw new Error('올바른 단일 A1 주소가 필요합니다.');return {r:+m[2]-1,c:colIndex(m[1])};}
function mergeAddress(merge) {return cellAddress(merge.r1,merge.c1)+':'+cellAddress(merge.r2,merge.c2);}
function nativeBorder(border) {
  if(!border)return null;
  if(border.style===-4142)return {active:false};
  return {active:true,style:border.style,weight:border.weight,color:oleColor(border.color)};
}
function modelBorder(style,key) {
  if(!style[key])return {active:false};
  const [line,weight]=BORDER_STYLE[style[key+'s']??'thin']??[null,null];
  return {active:true,style:line,weight,color:color(style[key+'c']??'#000000')};
}
export function normalizeNativeStyle(style) {
  if(!style)return null;
  const font=style.font??{},fill=style.fill??{},alignment=style.alignment??{},borders={};
  for(const side of Object.keys(SIDES))borders[side]=nativeBorder(style.borders?.[side]);
  return {
    font:{name:typeof font.name==='string'?font.name.toLowerCase():null,size:font.size??null,bold:font.bold??null,italic:font.italic??null,underline:font.underline==null?null:font.underline!==-4142,strike:font.strike??null,color:oleColor(font.color)},
    fill:{kind:fill.pattern===-4142?'none':fill.pattern===1?'solid':null,color:fill.pattern===-4142?null:oleColor(fill.color)},borders,
    numberFormat:style.numberFormat==null?null:formatCode(style.numberFormat),
    alignment:{horizontal:HORIZONTAL[alignment.horizontal]??null,vertical:VERTICAL[alignment.vertical]??null,wrap:alignment.wrap??null,shrink:alignment.shrink??null,indent:alignment.indent??null,orientation:alignment.orientation===-4128?0:alignment.orientation===-4166?255:alignment.orientation===-4170?-90:alignment.orientation===-4171?90:alignment.orientation??null},
  };
}
function normalizeModelStyle(style,book,fmtCode,fileCode) {
  const borders={};for(const [side,key]of Object.entries(SIDES))borders[side]=modelBorder(style,key);
  return {
    font:{name:String(style.font??book.defaultFont?.name??'맑은 고딕').toLowerCase(),size:style.size??book.defaultFont?.size??11,bold:!!style.bold,italic:!!style.italic,underline:!!style.underline,strike:!!style.strike,color:color(style.color??'#000000')},
    fill:{kind:style.pattern||style.gradient?null:style.fill?'solid':'none',color:style.fill?color(style.fill):null},borders,
    numberFormat:formatCode(fileCode(fmtCode(style)??'General')),
    alignment:{horizontal:style.align??'general',vertical:style.valign??'bottom',wrap:!!style.wrap,shrink:!!style.shrink,indent:style.indent??0,orientation:style.rotate??0},
  };
}
function boundedChildren(children,depth,budget){const out=[];for(const child of children){if(budget.remaining<=0)break;out.push(objectSample(child,child.src?'images':'shapes',depth+1,budget));}return out;}
function objectSample(object,category,depth=0,budget={remaining:512}) {
  budget.remaining--;
  const children=Array.isArray(object.children)?depth>=8?[]:boundedChildren(object.children,depth,budget):null;
  const area=category==='charts'?object.chartAreaFormat:null;
  const lineColor=category==='images'?object.border:category==='charts'?(area?.lineMode==='solid'?area.stroke:object.border):object.stroke;
  const lineKnown=category!=='charts'||area?.lineMode && area.lineMode!=='auto'||object.border!=null;
  const fillValue=category==='shapes'?object.fill:category==='charts'?area?.fill:undefined;
  return {category,depth,name:object.name??null,z:object.z??null,geometry:{left:object.x,top:object.y,width:object.w,height:object.h},rotation:object.rot??0,visible:!object.hidden,placement:{twoCell:1,oneCell:2,absolute:3}[object.placement??'oneCell']??null,
    line:{active:lineKnown?!!lineColor:null,color:color(lineColor),width:category==='images'?object.borderW??1:category==='charts'?area?.strokeWidth??1:object.strokeWidth??1,dash:object.dash??area?.dash??'solid',compound:object.compound??area?.compound??'sng'},
    fill:category==='images'?null:{kind:object.grad||object.pattern||area?.fillMode && area.fillMode!=='solid' && area.fillMode!=='none'?null:fillValue?'solid':'none',color:color(fillValue)},
    ...(children?{children,childrenLimited:children.length<object.children.length}:{}),
  };
}
/** 이미 가져온 모델 하나에서 Native 주소의 서식/geometry만 추출합니다. 값·수식·미디어 바이트는 복사하지 않습니다. */
export async function sampleNativeTargets(book,targets,{source=SOURCE,maxCells=64,maxRows=32,maxColumns=64,maxObjects=512}={}) {
  const [{Axis},{fmtCode,fileCode}]=await Promise.all([import(pathToFileURL(resolve(source,'axis.js')).href),import(pathToFileURL(resolve(source,'format.js')).href)]);
  const result={schemaVersion:1,scope:'bounded model direct styles and logical geometry; no formula evaluation, conditional-format rendering, image pixel or Safari test',units:{geometry:'CSS px at96/in, independent of browser/device zoom',fontSize:'points',nativeGeometryMultiplier:POINT_TO_CSS_PX,columnWidthCharacters:'not directly compared to pixels'},sheets:[]};
  for(const target of targets.sheets??[]) {
    const si=target.index??book.sheets.findIndex(s=>s.name===target.name),sheet=book.sheets[si];
    if(!sheet || target.name && target.name!==sheet.name)throw new Error('Native targets의 시트 이름/index가 모델과 다릅니다.');
    const cols=new Axis(sheet.defColW??64,sheet.colWidths,[sheet.hiddenCols],16384),rows=new Axis(sheet.defRowH??20,sheet.rowHeights,[sheet.hiddenRows,sheet.filter?.hidden,...(sheet.tables??[]).map(t=>t.filter?.hidden)],1048576);
    const item={index:si,name:sheet.name,conditionalDefinitions:(sheet.cond??[]).length,cells:[],rows:[],columns:[],drawings:[]};
    for(const address of [...new Set(target.cells??[])].slice(0,maxCells)) {
      const {r,c}=parseAddress(address),style=book.styleAt(si,r,c),merge=(sheet.merges??[]).find(m=>r>=m.r1&&r<=m.r2&&c>=m.c1&&c<=m.c2);
      item.cells.push({address,direct:normalizeModelStyle(style,book,fmtCode,fileCode),rawStyle:style,geometry:{left:cols.pos(c),top:rows.pos(r),width:cols.size(c),height:rows.size(r)},merged:!!merge,...(merge?{mergeAddress:mergeAddress(merge)}:{})});
    }
    for(const row of [...new Set(target.rows??[])].slice(0,maxRows)) {if(!Number.isInteger(row)||row<1||row>1048576)throw new Error('Native row는 1기반 행 번호입니다.');const r=row-1;item.rows.push({row,height:rows.size(r),top:rows.pos(r),hidden:rows.isHidden(r)});}
    for(const column of [...new Set(target.columns??[])].slice(0,maxColumns)) {if(!/^[A-Z]{1,3}$/.test(column)||colIndex(column)>=16384)throw new Error('Native column은 A~XFD입니다.');const c=colIndex(column);item.columns.push({column,width:cols.size(c),left:cols.pos(c),hidden:cols.isHidden(c)});}
    const objects=[];for(const category of ['charts','images','shapes','slicers'])for(const object of sheet[category]??[])objects.push({object,category});
    objects.sort((a,b)=>(a.object.z??Infinity)-(b.object.z??Infinity));const budget={remaining:maxObjects};item.drawingCount=objects.length;for(const entry of objects){if(budget.remaining<=0)break;item.drawings.push(objectSample(entry.object,entry.category,0,budget));}item.drawingsLimited=item.drawings.length<objects.length;
    result.sheets.push(item);
  }
  return result;
}
export function compareNativeSamples(native,model,{geometryTolerance=.51,drawingTolerance=.51,maxDifferences=40}={}) {
  if(!Number.isFinite(geometryTolerance)||geometryTolerance<0||!Number.isFinite(drawingTolerance)||drawingTolerance<0||!Number.isSafeInteger(maxDifferences)||maxDifferences<0)throw new Error('비교 허용 오차/표본 수가 올바르지 않습니다.');
  const result={schemaVersion:1,equal:false,scope:'Native read-only bounded semantic styles/logical geometry vs model; raw number-format code differences may be display-equivalent; no actual screen/iPad verdict',units:{pointToCssPixel:POINT_TO_CSS_PX,fontSize:'points',columnWidth:'character units excluded; native Width points compared',zoomDprApplied:false},tolerances:{geometryCssPx:geometryTolerance,drawingCssPx:drawingTolerance},counts:{checks:0,mismatches:0,unavailable:0,displayStyleChanged:0,underlineVariantsNotCompared:0},categories:{},differences:[],unavailableExamples:[]};
  const issue=(category,path,a,b)=>{result.counts.mismatches++;result.categories[category]=(result.categories[category]??0)+1;if(result.differences.length<maxDifferences)result.differences.push({category,path,native:a,model:b,...(typeof a==='number'&&typeof b==='number'?{delta:b-a}:{})});};
  const missing=(path,why)=>{result.counts.unavailable++;if(result.unavailableExamples.length<maxDifferences)result.unavailableExamples.push({path,why});};
  const check=(category,path,a,b,tolerance=0)=>{if(a==null||b==null||typeof a==='number'&&!Number.isFinite(a)||typeof b==='number'&&!Number.isFinite(b)){missing(path,'value or supported semantic mapping unavailable');return;}result.counts.checks++;if(typeof a==='number'&&typeof b==='number'?Math.abs(a-b)>tolerance:a!==b)issue(category,path,a,b);};
  const geometry=(path,a,b,tolerance,keys=['left','top','width','height'])=>{for(const key of keys)check('geometry',path+'.'+key,a?.[key]==null?null:a[key]*POINT_TO_CSS_PX,b?.[key],tolerance);};
  const style=(path,a,b)=>{
    if(!a||!b){missing(path,'style unavailable');return;}
    for(const key of ['name','size','bold','italic','underline','strike','color'])check('font',path+'.font.'+key,a.font[key],b.font[key],key==='size'?.01:0);
    check('fill',path+'.fill.kind',a.fill.kind,b.fill.kind);if(a.fill.kind!=='none'&&b.fill.kind!=='none')check('fill',path+'.fill.color',a.fill.color,b.fill.color);
    for(const side of Object.keys(SIDES)){const x=a.borders[side],y=b.borders[side];check('border',path+'.'+side+'.active',x?.active,y?.active);if(x?.active&&y?.active)for(const key of ['style','weight','color'])check('border',path+'.'+side+'.'+key,x[key],y[key]);}
    check('number-format',path+'.numberFormat',a.numberFormat,b.numberFormat);
    for(const key of Object.keys(a.alignment))check('alignment',path+'.alignment.'+key,a.alignment[key],b.alignment[key]);
  };
  for(const ns of native.sheets??[]) {
    const ms=model.sheets?.find(s=>s.index===ns.index&&s.name===ns.name),prefix='sheet['+ns.index+']';
    if(!ms){issue('presence',prefix,'present','missing');continue;}
    for(const nc of ns.cells??[]) {
      const mc=ms.cells.find(c=>c.address===nc.address),path=prefix+'.'+nc.address;if(!mc){issue('presence',path,'present','missing');continue;}
      if(nc.status!=='complete')missing(path,'Native sample partial; individual fields checked only if present');
      const direct=normalizeNativeStyle(nc.direct),display=normalizeNativeStyle(nc.display);style(path,direct,mc.direct);
      if(JSON.stringify(direct)!==JSON.stringify(display))result.counts.displayStyleChanged++;
      if(nc.direct?.font?.underline!=null&&![2,-4142].includes(nc.direct.font.underline))result.counts.underlineVariantsNotCompared++;
      geometry(path,nc.geometry,mc.geometry,geometryTolerance);check('merge',path+'.merged',nc.merged,mc.merged);
      if(nc.merged&&mc.merged)check('merge',path+'.mergeAddress',nc.mergeAddress,mc.mergeAddress);
    }
    for(const nr of ns.rows??[]){const mr=ms.rows.find(r=>r.row===nr.row),path=prefix+'.row'+nr.row;geometry(path,nr,mr,geometryTolerance,['top','height']);check('hidden',path+'.hidden',nr.hidden,mr?.hidden);}
    for(const nc of ns.columns??[]){const mc=ms.columns.find(c=>c.column===nc.column),path=prefix+'.col'+nc.column;geometry(path,nc,mc,geometryTolerance,['left','width']);check('hidden',path+'.hidden',nc.hidden,mc?.hidden);}
    const compareObjects=(nativeObjects,modelObjects,path)=>{
      const unused=new Set(modelObjects);
      for(const no of nativeObjects) {
        let mo=modelObjects.find(o=>unused.has(o)&&o.name&&o.name===no.Name),matching='name';
        if(!mo){mo=modelObjects.find(o=>unused.has(o)&&o.z===no.ZOrderPosition);matching='z';}
        if(!mo){issue('drawing-presence',path+'.'+no.Name,'present','missing matching name/z');continue;}unused.delete(mo);
        const here=path+'.'+no.Name;geometry(here,{left:no.Left,top:no.Top,width:no.Width,height:no.Height},mo.geometry,drawingTolerance);
        if(matching==='name')check('drawing-name',here+'.name',no.Name,mo.name);
        check('drawing-visible',here+'.visible',no.Visible==null||![-1,0].includes(no.Visible)?null:no.Visible===-1,mo.visible);check('drawing-placement',here+'.placement',no.Placement,mo.placement);
        const angle=value=>((value%360)+360)%360;check('drawing-rotation',here+'.rotation',no.Rotation==null?null:angle(no.Rotation),angle(mo.rotation),.51);
        if(no.line){const active=no.line.visible==null||![-1,0].includes(no.line.visible)?null:no.line.visible===-1;check('drawing-line',here+'.line.active',active,mo.line.active);if(active&&mo.line.active){check('drawing-line',here+'.line.color',oleColor(no.line.color?.rgb),mo.line.color);check('drawing-line',here+'.line.width',no.line.weight==null?null:no.line.weight*POINT_TO_CSS_PX,mo.line.width,.51);check('drawing-line',here+'.line.dash',DASH[no.line.dash]??null,mo.line.dash);check('drawing-line',here+'.line.compound',COMPOUND[no.line.style]??null,mo.line.compound);}}
        if(no.fill&&mo.fill){const kind=no.fill.visible===0?'none':no.fill.visible===-1&&no.fill.type===1?'solid':null;check('drawing-fill',here+'.fill.kind',kind,mo.fill.kind);if(kind==='solid'&&mo.fill.kind==='solid')check('drawing-fill',here+'.fill.color',oleColor(no.fill.color?.rgb),mo.fill.color);}
        if(no.children){if(no.groupLimited||mo.childrenLimited)missing(here+'.children','bounded group nodes/depth limited');else compareObjects(no.children,mo.children??[],here+'.children');}
      }
      if(!ns.shapesLimited&&!ms.drawingsLimited)for(const extra of unused)issue('drawing-presence',path+'.'+(extra.name??extra.z),'missing native','extra model object');
    };
    compareObjects(ns.drawings??[],ms.drawings??[],prefix+'.drawings');
    if(ns.shapesLimited||ms.drawingsLimited)missing(prefix+'.drawings','bounded drawing count limited');
  }
  if(native.opened!==true||native.readOnly!==true||native.ownershipVerified!==true||native.originalUnchanged!==true||native.status!=='complete')missing('native.scope','native result not complete/read-only unchanged');
  result.equal=result.counts.mismatches===0&&result.counts.unavailable===0;return result;
}
async function main(args) {
  const options={};for(let i=0;i<args.length;i+=2){const key=args[i];if(!['--native','--model','--out'].includes(key)||!args[i+1])throw new Error('사용법: node tools/native-layout-compare.mjs --native NATIVE.json --model model-samples.json --out NEW.json');options[key.slice(2)]=resolve(args[i+1]);}
  if(!options.native||!options.model||!options.out)throw new Error('--native --model --out이 필요합니다.');
  const json=async path=>JSON.parse((await readFile(path,'utf8')).replace(/^\ufeff/,'')),report=compareNativeSamples(await json(options.native),await json(options.model));
  await writeFile(options.out,JSON.stringify(report,null,2),{flag:'wx'});console.log(JSON.stringify({equal:report.equal,...report.counts,categories:report.categories}));if(!report.equal)process.exitCode=1;
}
if(process.argv[1]&&resolve(process.argv[1])===TOOL)main(process.argv.slice(2)).catch(error=>{console.error(error.message);process.exitCode=2;});
