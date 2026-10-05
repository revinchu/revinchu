import { el, openDialog } from './ui.js';
import { svgImageBlob, objectRasterSize } from './object-image-export.js';

/** 미리보기와 저장은 같은 독립 SVG를 사용한다. 파일 쓰기는 호출자가 담당한다. */
export function openImageExportDialog({title='그림으로 저장',name='그림',description='',svgAllowed=true,initialFormat='png',gridlines,makeSvg,valid,save}) {
  let alive=true,revision=0,prepared=null,previewUrl=null,automaticScale=true;
  const field=(caption,node)=>el('label',{class:'image-export-field'},el('span',{},caption),node);
  const format=el('select',{name:'imageFormat','aria-label':'파일 형식'},[['png','PNG (.png)'],['jpeg','JPEG (.jpg)'],...(svgAllowed?[['svg','SVG (.svg)']]:[])].map(([v,t])=>el('option',{value:v},t)));
  format.value=svgAllowed&&initialFormat==='svg'?'svg':'png';
  const filename=el('input',{name:'imageFilename',type:'text',value:name,'aria-label':'파일 이름',maxLength:180});
  const scale=el('select',{name:'imageScale','aria-label':'해상도'},[[1,'1배 · 기본'],[2,'2배 · 선명하게'],[3,'3배 · 고해상도']].map(([v,t])=>el('option',{value:v},t)));scale.value='2';
  const background=el('select',{name:'imageBackground','aria-label':'배경'},el('option',{value:'transparent'},'투명'),el('option',{value:'white'},'흰색'));background.value=gridlines===undefined?'transparent':'white';
  const grid=gridlines===undefined?null:el('input',{name:'imageGridlines',type:'checkbox',checked:gridlines});
  const status=el('div',{class:'image-export-status',role:'status','aria-live':'polite'},'미리보기를 준비하고 있습니다.');
  const problem=el('div',{class:'image-export-error',role:'alert',hidden:true});
  const preview=el('img',{alt:'저장할 그림 미리보기'}),previewBox=el('div',{class:'image-export-preview'},preview);
  const options=el('div',{class:'image-export-options'},field('파일 이름(N)',filename),field('파일 형식(F)',format),field('해상도(R)',scale),field('배경(B)',background),grid?el('label',{class:'image-export-check'},grid,'눈금선 포함(G)'):null);
  const body=el('div',{class:'image-export-body'},el('p',{class:'image-export-description'},description),previewBox,status,options,problem);
  const current=()=>({format:format.value,scale:Number(scale.value),background:format.value==='jpeg'?'white':background.value,gridlines:grid?.checked});
  const showError=message=>{problem.textContent=message;problem.hidden=!message;};
  const fits=(width,height,scale)=>{try{objectRasterSize(width,height,scale);return true;}catch{return false;}};
  const canRaster=()=>prepared&&fits(prepared.width,prepared.height,Number(scale.value));
  const update=()=>{
    const o=current();scale.disabled=o.format==='svg';background.disabled=o.format==='jpeg';previewBox.classList.toggle('white',o.background==='white');
    if(prepared){const k=o.format==='svg'?1:o.scale;status.textContent=`${Math.ceil(prepared.width*k).toLocaleString()} × ${Math.ceil(prepared.height*k).toLocaleString()} px`+(o.format==='svg'?' · 확대해도 선명한 SVG':'')+(prepared.warnings?.length?' · '+prepared.warnings.join(' '):'');showError(o.format!=='svg'&&!canRaster()?'한 변 4,096px, 총 1,600만 픽셀 이내로 저장할 수 있습니다. 해상도를 낮추거나 선택 영역을 줄이세요.':'');}
  };
  const refresh=()=>{
    const request=++revision;prepared=null;showError('');status.textContent='미리보기를 준비하고 있습니다.';
    return Promise.resolve().then(()=>{if(!alive||!valid())throw new Error('문서 또는 선택한 대상이 변경되었습니다. 창을 다시 여세요.');return makeSvg(current());}).then(async result=>{
      if(!alive||request!==revision)return;
      if(!valid())throw new Error('문서 또는 선택한 대상이 변경되었습니다. 창을 다시 여세요.');
      prepared=result;
      if(automaticScale){const preferred=[2,1].find(n=>fits(result.width,result.height,n));scale.value=String(preferred??1);}
      const fontWarnings=[];
      const blob=await svgImageBlob(result.svg,{width:result.width,height:result.height,format:'svg',scale:1,background:'transparent',onWarning:message=>fontWarnings.push(message)});
      if(fontWarnings.length)result.warnings=[...(result.warnings||[]),...fontWarnings];
      if(!alive||request!==revision)return;
      if(previewUrl)URL.revokeObjectURL(previewUrl);previewUrl=URL.createObjectURL(blob);preview.src=previewUrl;update();
    }).catch(error=>{if(alive&&request===revision){prepared=null;status.textContent='미리보기를 만들지 못했습니다.';showError(error.message);}});
  };
  format.addEventListener('change',update);background.addEventListener('change',update);scale.addEventListener('change',()=>{automaticScale=false;update();});grid?.addEventListener('change',refresh);
  const dialog=openDialog({title,width:610,body,initialFocus:filename,onClose:()=>{alive=false;revision++;if(previewUrl)URL.revokeObjectURL(previewUrl);},buttons:[
    {label:'저장(S)',accessKey:'s',primary:true,action:()=>{
      if(!filename.value.trim()){showError('파일 이름을 입력하세요.');filename.focus();return false;}
      if(!valid()){showError('문서 또는 선택한 대상이 변경되었습니다. 창을 다시 여세요.');return false;}
      if(!prepared){showError('미리보기를 준비한 뒤 다시 저장하세요.');return false;}
      if(format.value!=='svg'&&!canRaster()){update();return false;}
      const data=prepared,o=current(),ext=o.format==='jpeg'?'jpg':o.format;
      const base=filename.value.trim().replace(/\.(png|jpe?g|svg)$/i,'').replace(/[\\/:*?"<>|]/g,'_')||'그림';
      return save(`${base}.${ext}`,()=>svgImageBlob(data.svg,{width:data.width,height:data.height,...o}));
    }},{label:'취소',accessKey:'none'}]});
  dialog.root.classList.add('image-export-dialog');refresh();return dialog;
}
