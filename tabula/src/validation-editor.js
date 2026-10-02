// 데이터 유효성 대화상자의 초안. 확인 전에는 통합 문서에 쓰지 않는다.
import { el } from './ui.js';
import { VALIDATION_TYPES, VALIDATION_OPS } from './validation.js';
import { parseInput, formatCode } from './format.js';

let sequence=0;
export function createValidationEditor({rule={},date1904=false,canApplySame=false,onPick}={}) {
  const uid=`dv-editor-${++sequence}`,fields={},labels={},pickers={};
  const label=(name,text,node)=>{node.id=`${uid}-${name}`;node.name=name;fields[name]=node;const title=el('span',{},text);labels[name]=title;return el('label',{class:'dv-field',for:node.id},title,node);};
  const check=(name,text,value)=>{const node=el('input',{type:'checkbox',checked:value});const wrap=label(name,text,node);wrap.classList.add('dv-check');wrap.prepend(node);return wrap;};
  const select=(name,text,options)=>label(name,text,el('select',{},options.map(([value,title])=>el('option',{value},title))));
  const text=(name,title,multiline=false,maxLength)=>label(name,title,el(multiline?'textarea':'input',{...(multiline?{rows:5}:{type:'text'}),maxLength,spellcheck:false}));
  const set=(title,...children)=>el('fieldset',{class:'dv-fieldset'},el('legend',{},title),...children);
  const type=select('type','제한 대상(A):',VALIDATION_TYPES.map(t=>[t.id,t.label]));
  const blank=check('allowBlank','공백 무시(B)',true),dd=check('showDropdown','드롭다운 표시(I)',true);
  const op=select('op','제한 방법(D):',VALIDATION_OPS.map(t=>[t.id,t.label]));
  const f1=text('f1','최소값(M):'),f2=text('f2','최대값(M):');
  for(const [name,row] of [['f1',f1],['f2',f2]]){
    const input=fields[name];const button=el('button',{type:'button',class:'dv-ref-pick','data-access-key':'none',onclick:()=>onPick?.(name,input)},'▴');
    pickers[name]=button;const box=el('span',{class:'dv-reference'},input,button);row.append(box);
  }
  const same=check('applySame','변경 내용을 설정이 같은 모든 셀에 적용(P)',false);fields.applySame.disabled=!canApplySame;
  const settings=el('div',{class:'dv-settings'},set('유효성 조건',el('div',{class:'dv-condition-grid'},type,el('div',{class:'dv-options'},blank,dd),op,f1,f2)),same);
  const showPrompt=check('showPrompt','셀을 선택하면 설명 메시지 표시(S)',true);
  const promptTitle=text('promptTitle','제목(T):',false,32),prompt=text('prompt','설명 메시지(I):',true,255);
  const promptPage=el('div',{},showPrompt,set('셀을 선택하면 나타낼 설명 메시지',promptTitle,prompt));
  const showError=check('showError','유효하지 않은 데이터를 입력하면 오류 메시지 표시(S)',true);
  const errorStyle=select('errorStyle','스타일(Y):',[['stop','중지'],['warning','경고'],['info','정보']]);
  const icon=el('span',{class:'dv-icon stop','aria-hidden':true},'✕');
  const errorTitle=text('errorTitle','제목(T):',false,32),error=text('error','오류 메시지(E):',true,225);
  const errorPage=el('div',{},showError,set('유효하지 않은 데이터를 입력하면 나타낼 오류 메시지',el('div',{class:'dv-error-grid'},el('div',{class:'dv-error-style'},errorStyle,icon),el('div',{},errorTitle,error))));
  const ime=select('imeMode','모드(M):',[
    ['noControl','현재 상태 유지'],['off','영문'],['on','입력기 켜기'],['disabled','사용 안 함'],
    ['halfHangul','한글'],['fullHangul','한글(전각)'],['halfAlpha','영숫자(반각)'],['fullAlpha','영숫자(전각)'],
    ['hiragana','히라가나'],['fullKatakana','가타카나(전각)'],['halfKatakana','가타카나(반각)'],
  ]);
  const imePage=el('div',{},set('입력기',ime),el('p',{class:'dv-ime-note'},'입력기 모드는 Excel 파일에 저장됩니다. 브라우저에서는 운영체제의 IME 모드를 직접 전환할 수 없습니다.'));
  const pages=[['설정',settings],['설명 메시지',promptPage],['오류 메시지',errorPage],['IME 모드',imePage]];
  const tabs=el('div',{class:'dlg-tabs',role:'tablist','aria-label':'데이터 유효성 범주'}),panel=el('div',{class:'dv-page',role:'tabpanel',id:`${uid}-panel`});
  const errorBox=el('div',{class:'dv-settings-error',role:'alert',hidden:true});
  let selected=0;
  const showTab=i=>{selected=i;[...tabs.children].forEach((b,j)=>{b.classList.toggle('on',i===j);b.setAttribute('aria-selected',String(i===j));b.tabIndex=i===j?0:-1;});panel.setAttribute('aria-label',pages[i][0]);panel.setAttribute('aria-labelledby',`${uid}-tab-${i}`);panel.replaceChildren(pages[i][1]);};
  pages.forEach(([name],i)=>tabs.append(el('button',{type:'button',class:'dlg-tab',role:'tab',id:`${uid}-tab-${i}`,'aria-controls':panel.id,'data-access-key':'none',onclick:()=>showTab(i)},name)));
  tabs.addEventListener('keydown',e=>{if(!['ArrowLeft','ArrowRight','Home','End'].includes(e.key)||e.isComposing)return;e.preventDefault();e.stopPropagation();const n=e.key==='Home'?0:e.key==='End'?3:(selected+(e.key==='ArrowRight'?1:-1)+4)%4;showTab(n);tabs.children[n].focus();});
  const body=el('div',{class:'dv-editor'},tabs,panel,errorBox);
  body.addEventListener('keydown',e=>{if(e.ctrlKey&&e.key==='Tab'){e.preventDefault();e.stopPropagation();const n=(selected+(e.shiftKey?-1:1)+4)%4;showTab(n);tabs.children[n].focus();}});
  const refresh=()=>{
    const t=fields.type.value,ranged=['whole','decimal','date','time','textLength'].includes(t),two=ranged&&['between','notBetween'].includes(fields.op.value);
    op.hidden=!ranged&&t!=='any';fields.op.disabled=!ranged;fields.allowBlank.disabled=t==='any';dd.hidden=t!=='list';
    f1.hidden=t==='any';f2.hidden=!two;
    labels.f1.textContent=t==='list'?'원본(S):':t==='custom'?'수식(F):':t==='date'?(two?'시작 날짜(S):':'날짜(S):'):t==='time'?(two?'시작 시간(S):':'시간(S):'):two?'최소값(M):':['lessThan','lessThanOrEqual'].includes(fields.op.value)?'최대값(M):':['greaterThan','greaterThanOrEqual'].includes(fields.op.value)?'최소값(M):':'값(V):';
    labels.f2.textContent=t==='date'?'끝 날짜(E):':t==='time'?'끝 시간(E):':'최대값(M):';
    pickers.f1.setAttribute('aria-label',t==='list'?'원본 범위 선택':'첫 번째 조건 범위 선택');pickers.f2.setAttribute('aria-label','두 번째 조건 범위 선택');
    fields.f1.placeholder=t==='list'?'사과,배,포도 또는 =$A$1:$A$5':t==='custom'?'=A1>0':t==='date'?'2026-01-01':t==='time'?'09:00':'';fields.f2.placeholder=fields.f1.placeholder;
    for(const name of ['promptTitle','prompt'])fields[name].disabled=!fields.showPrompt.checked;
    for(const name of ['errorStyle','errorTitle','error'])fields[name].disabled=!fields.showError.checked;
    icon.className=`dv-icon ${fields.errorStyle.value}`;icon.textContent={stop:'✕',warning:'!',info:'i'}[fields.errorStyle.value];icon.classList.toggle('disabled',!fields.showError.checked);
  };
  const displayFormula=(f,t)=>{const x=String(f??'');if(t==='list'&&x.startsWith('"')&&x.endsWith('"'))return x.slice(1,-1).replaceAll('""','"');if(['date','time'].includes(t)&&/^-?\d+(\.\d+)?$/.test(x))return formatCode(Number(x),t==='date'?'yyyy-mm-dd':'hh:mm:ss',date1904).text;if(t==='list'&&!x.startsWith('=')&&/^(?:'[^']+'!|[^!]+!)?\$?[A-Z]{1,3}\$?\d+(?::\$?[A-Z]{1,3}\$?\d+)?$/i.test(x))return '='+x;return x;};
  const load=source=>{
    const defaults={type:'any',op:'between',allowBlank:true,showDropdown:true,showPrompt:true,showError:true,errorStyle:'stop',imeMode:'noControl'};
    for(const [name,node] of Object.entries(fields)){if(name==='applySame')continue;const value=source[name]??defaults[name]??'';if(node.type==='checkbox')node.checked=!!value;else node.value=['f1','f2'].includes(name)?displayFormula(value,source.type):value;}
    refresh();errorBox.hidden=true;
  };
  const read=()=>{
    const rule={};for(const [name,node]of Object.entries(fields))if(name!=='applySame')rule[name]=node.type==='checkbox'?node.checked:node.value;
    if(!['whole','decimal','date','time','textLength'].includes(rule.type))delete rule.op;
    for(const name of ['f1','f2']){
      let value=String(rule[name]??'').trim();
      if(rule.type==='any'||name==='f2'&&!['between','notBetween'].includes(rule.op))value='';
      if(value&&['date','time'].includes(rule.type)&&!value.startsWith('=')){const n=parseInput(value,date1904).value;if(typeof n==='number')value=String(n);}
      if(value)rule[name]=rule.type==='list'&&name==='f1'&&!value.startsWith('=')?'"'+value.replaceAll('"','""')+'"':value;else delete rule[name];
    }
    for(const name of ['promptTitle','prompt','errorTitle','error'])if(!rule[name])delete rule[name];
    return rule;
  };
  body.addEventListener('change',refresh);load(rule);showTab(0);
  return {body,fields,read,refresh,showTab,reset:()=>{load({});showTab(0);},setError:message=>{errorBox.textContent=message??'';errorBox.hidden=!message;}};
}
