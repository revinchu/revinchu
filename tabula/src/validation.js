// 데이터 유효성 검사 (DOM 없음)
// 규칙: { r1,c1,r2,c2, type, op, f1, f2, allowBlank, showDropdown, errorStyle, errorTitle, error, promptTitle, prompt }
import { parse, evalAny, evaluateArray, shiftFormula, isError, Range, RefValue } from './formula.js';
import { validationEvaluation } from './validation-eval.js';
import { blockValue } from './block.js';
import { parseInput, formatValue } from './format.js';

export const VALIDATION_TYPES = [
  { id: 'any', label: '모든 값' },
  { id: 'whole', label: '정수' },
  { id: 'decimal', label: '소수점' },
  { id: 'list', label: '목록' },
  { id: 'date', label: '날짜' },
  { id: 'time', label: '시간' },
  { id: 'textLength', label: '텍스트 길이' },
  { id: 'custom', label: '사용자 지정' },
];

export const VALIDATION_OPS = [
  { id: 'between', label: '해당 범위' },
  { id: 'notBetween', label: '제외 범위' },
  { id: 'equal', label: '=' },
  { id: 'notEqual', label: '<>' },
  { id: 'greaterThan', label: '>' },
  { id: 'lessThan', label: '<' },
  { id: 'greaterThanOrEqual', label: '>=' },
  { id: 'lessThanOrEqual', label: '<=' },
];

export function validationAt(sheet, r, c) {
  const list = sheet.validations ?? [];
  for (let i = list.length - 1; i >= 0; i--) {
    const v = list[i];
    if (r >= v.r1 && r <= v.r2 && c >= v.c1 && c <= v.c2) return v;
  }
  return null;
}

export const VALIDATION_IME_MODES = ['noControl','off','on','disabled','hiragana','fullKatakana','halfKatakana','fullAlpha','halfAlpha','fullHangul','halfHangul'];
export const VALIDATION_LIST_LIMIT = 100000;

/** 모델의 리터럴과 수식을 구분한다. XLSX의 이름/함수 목록은 읽을 때 =를 보충한다. */
function formulaSource(rule, text) {
  const t=String(text??'').trim();
  if(!t)return null;
  if(t.startsWith('='))return t;
  if(rule.type==='custom')return '='+t;
  if(rule.type==='list') {
    if(t.startsWith('"'))return null;
    try { const ast=parse(t);return ['ref','range','sref','spill'].includes(ast.type)?'='+t:null; } catch { return null; }
  }
  const parsed=parseInput(t);
  if(typeof parsed.value==='number'||typeof parsed.value==='boolean')return null;
  try {parse(t);return '='+t;} catch {return null;}
}
function operand(wb,si,rule,text,r,c,session) {
  if(text===undefined||text===null||text==='')return null;
  const f=formulaSource(rule,text);
  if(!f)return parseInput(String(text).trim(),wb.date1904).value;
  const ev=session??validationEvaluation(wb,si,r,c);
  return ev.attempt(()=>{const ast=parse(shiftFormula(f,r-rule.r1,c-rule.c1).slice(1)),value=evalAny(ast,ev.context);return value instanceof RefValue&&value.single?ev.context.cell(value.sheet,value.r1,value.c1):value instanceof Range?value.rows[0]?.[0]??null:value;});
}
function compare(op,v,a,b) {
  switch(op) {
    case 'notBetween':return v<a||v>b;
    case 'equal':return v===a;
    case 'notEqual':return v!==a;
    case 'greaterThan':return v>a;
    case 'lessThan':return v<a;
    case 'greaterThanOrEqual':return v>=a;
    case 'lessThanOrEqual':return v<=a;
    default:return v>=a&&v<=b;
  }
}
const splitList=s=>s.split(',').map(x=>x.trim()).filter(x=>x!=='');

/** 목록은 원시 값과 표시 문구를 함께 보존한다. 참조 범위/이름/INDIRECT/배열은 1차원만 허용. */
export function validationList(wb,si,rule,r=rule.r1,c=rule.c1) {
  const source=String(rule.f1??'').trim(),f=formulaSource(rule,source);
  if(!f) {
    const literal=source.startsWith('"')&&source.endsWith('"')?source.slice(1,-1).replace(/""/g,'"'):source;
    const items=splitList(literal);return {items,itemValues:[...items],values:[...items],literal:true,limited:false,error:null,hasBlank:false,named:false};
  }
  const ev=validationEvaluation(wb,si,r,c),text=shiftFormula(f,r-rule.r1,c-rule.c1).slice(1);
  let node,ref;
  try {node=parse(text);ref=evalAny(node,ev.context);}catch{return {items:[],itemValues:[],values:[],literal:false,limited:false,error:'목록 원본 수식이나 참조를 확인하세요.'};}
  let rows=null,styleAt=null;
  if(ref instanceof RefValue) {
    if(ref.r1!==ref.r2&&ref.c1!==ref.c2)return {items:[],itemValues:[],values:[],literal:false,limited:false,error:'목록 원본은 한 행 또는 한 열이어야 합니다.'};
    const index=wb.resolveSheet(ref.sheet,si);
    if(index<0)return {items:[],itemValues:[],values:[],literal:false,limited:false,error:'목록 원본 시트를 찾을 수 없습니다.'};
    let r2=ref.r2,c2=ref.c2;
    // 전체 행/열 참조만 실제 사용 범위로 줄인다. 유한 원본은 빈 칸도 그대로 판정한다.
    if(r2-ref.r1>=1048575)r2=Math.max(ref.r1,wb.usedRange(index).rows-1);
    if(c2-ref.c1>=16383)c2=Math.max(ref.c1,wb.usedRange(index).cols-1);
    if((r2-ref.r1+1)*(c2-ref.c1+1)>VALIDATION_LIST_LIMIT)return {items:[],itemValues:[],values:[],literal:false,limited:true,error:'목록 원본은 100,000셀 이하로 지정하세요.'};
    rows=ev.attempt(()=>ev.context.range(ref.sheet,ref.r1,ref.c1,r2,c2));
    styleAt=(row,col)=>wb.styleAt(index,ref.r1+row,ref.c1+col);
  } else {
    const array=ev.attempt(()=>evaluateArray(node,ev.context));
    if(array instanceof Range)rows=array.rows;
    else return {items:[],itemValues:[],values:[],literal:false,limited:false,error:'목록 원본은 한 행 또는 한 열의 참조나 배열이어야 합니다.'};
  }
  if(isError(rows)||!Array.isArray(rows))return {items:[],itemValues:[],values:[],literal:false,limited:false,error:'목록 원본을 계산할 수 없습니다.'};
  if(rows.length>1&&rows.some(row=>row.length>1))return {items:[],itemValues:[],values:[],literal:false,limited:false,error:'목록 원본은 한 행 또는 한 열이어야 합니다.'};
  if(rows.length*(rows[0]?.length??0)>VALIDATION_LIST_LIMIT)return {items:[],itemValues:[],values:[],literal:false,limited:true,error:'목록 원본은 100,000셀 이하로 지정하세요.'};
  const items=[],itemValues=[],values=[],seen=new Set();let hasBlank=false;
  for(let row=0;row<rows.length;row++)for(let col=0;col<rows[row].length;col++) {
    const value=rows[row][col];
    if(value===null||value===undefined||value===''){hasBlank=true;continue;}
    if(isError(value))continue;
    values.push(value);
    const label=formatValue(value,styleAt?.(row,col),wb.date1904).text;
    if(label!==''&&!seen.has(label)){seen.add(label);items.push(label);itemValues.push(value);}
  }
  return {items,itemValues,values,literal:false,limited:false,error:null,hasBlank,named:node.type==='name'};
}
export function listItems(wb,si,rule,r=rule.r1,c=rule.c1) {return validationList(wb,si,rule,r,c).items;}

/** 후보 입력을 격리 평가한다. 원문/수식/파일 저장값·dirty/Undo를 변경하지 않는다. */
export function checkValidation(wb,si,rule,r,c,text,knownList) {
  if(!rule||rule.type==='any')return true;
  const blank=text===''||text===null||text===undefined;
  if(blank&&rule.allowBlank!==false)return true;
  const session=validationEvaluation(wb,si,r,c,blank?'':text),value=session.value();
  if(isError(value)&&rule.type!=='custom')return false;
  switch(rule.type) {
    case 'list': {
      const list=knownList??validationList(wb,si,rule,r,c);if(list.error)return false;
      if(rule.allowBlank!==false&&list.named&&list.hasBlank)return true;
      return list.values.some(it=>typeof it===typeof value&&(typeof it==='string'?it.toLowerCase()===value.toLowerCase():it===value))
        ||list.items.some(it=>it===String(value)||it.toLowerCase()===String(value).toLowerCase());
    }
    case 'whole':case 'decimal':case 'date':case 'time':case 'textLength': {
      const number=rule.type==='textLength'?String(value??'').length:value===null?0:value;
      if(typeof number!=='number'||!Number.isFinite(number)||(rule.type==='whole'&&!Number.isInteger(number)))return false;
      let a=operand(wb,si,rule,rule.f1,r,c,session),b=operand(wb,si,rule,rule.f2,r,c,session);
      const two=['between','notBetween'].includes(rule.op??'between');
      if(!String(rule.f1??'').trim()||two&&!String(rule.f2??'').trim())return false;
      if(a===null||two&&b===null){if(rule.allowBlank!==false)return true;a??=0;b??=0;}
      if(typeof a!=='number'||!Number.isFinite(a)||two&&(typeof b!=='number'||!Number.isFinite(b)))return false;
      return compare(rule.op??'between',number,a,b);
    }
    case 'custom': {
      const result=operand(wb,si,rule,rule.f1,r,c,session);
      return result===true||typeof result==='number'&&Number.isFinite(result)&&result!==0;
    }
    default:return true;
  }
}

/** 잘라진 적용 범위의 왼쪽 위가 달라져도 같은 상대 참조를 유지한다. */
export function relocateValidation(rule,range) {
  const result={...rule,...range};
  for(const key of ['f1','f2']) {const f=formulaSource(rule,rule[key]);if(f){const moved=shiftFormula(f,result.r1-rule.r1,result.c1-rule.c1);result[key]=String(rule[key]).startsWith('=')?moved:moved.slice(1);}}
  return result;
}

/** 좌표 자체는 제외하고 상대 참조는 signed offset으로 비교(#REF!로 변환하지 않음). */
export function validationSignature(rule) {
  const formulaKey=value=>{
    const f=formulaSource(rule,value);if(!f)return String(value??'').trim();
    try {
      const visit=value=>{
        if(Array.isArray(value))return value.map(visit);
        if(!value||typeof value!=='object')return value;
        const copy={};for(const key of Object.keys(value).sort())if(key!=='s'&&key!=='e')copy[key]=visit(value[key]);
        if(value.type==='ref') {const ref=copy.ref;for(const key of ['r1','r2'])if(!ref['a'+key]&&!ref.cols)ref[key]-=rule.r1;for(const key of ['c1','c2'])if(!ref['a'+key]&&!ref.rows)ref[key]-=rule.c1;if(ref.sheet)ref.sheet=ref.sheet.toLowerCase();}
        return copy;
      };
      return visit(parse(f.slice(1)));
    }catch{return f;}
  };
  return JSON.stringify({type:rule.type??'any',op:rule.op??'between',f1:formulaKey(rule.f1),f2:formulaKey(rule.f2),allowBlank:rule.allowBlank!==false,showDropdown:rule.showDropdown!==false,showError:rule.showError!==false,showPrompt:rule.showPrompt!==false,errorStyle:rule.errorStyle??'stop',errorTitle:rule.errorTitle??'',error:rule.error??'',promptTitle:rule.promptTitle??'',prompt:rule.prompt??'',imeMode:rule.imeMode??'noControl'});
}

/** 편집기의 설정 검증. 현재 셀 값의 유효 여부와는 별개다. */
export function validateValidationRule(wb,si,rule) {
  if(!VALIDATION_TYPES.some(t=>t.id===rule.type))return '제한 대상을 선택하세요.';
  if(!VALIDATION_IME_MODES.includes(rule.imeMode??'noControl'))return '올바른 IME 모드를 선택하세요.';
  for(const key of ['promptTitle','errorTitle'])if(String(rule[key]??'').length>32)return '설명과 오류 메시지의 제목은 32자 이하로 입력하세요.';
  for(const key of ['prompt','error'])if(String(rule[key]??'').length>255)return '설명과 오류 메시지는 255자 이하로 입력하세요.';
  if(rule.type==='any')return null;
  if(!String(rule.f1??'').trim())return rule.type==='list'?'목록 원본을 입력하세요.':'조건 또는 수식을 입력하세요.';
  if(rule.type==='list') {
    const f=formulaSource(rule,rule.f1);if(!f&&String(rule.f1).replace(/^"|"$/g,'').length>255)return '직접 입력하는 목록은 255자 이하로 입력하거나 셀 범위를 사용하세요.';
    return validationList(wb,si,rule).error;
  }
  if(rule.type==='custom') {
    try {parse(String(rule.f1).replace(/^=/,''));}catch{return '사용자 지정 수식의 형식을 확인하세요.';}
    const value=operand(wb,si,rule,rule.f1,rule.r1,rule.c1);
    return isError(value)||value!==null&&typeof value!=='boolean'&&typeof value!=='number'?'사용자 지정 수식을 계산할 수 없습니다. 참조와 함수 이름을 확인하세요.':null;
  }
  if(!VALIDATION_OPS.some(op=>op.id===(rule.op??'between')))return '제한 방법을 선택하세요.';
  const two=['between','notBetween'].includes(rule.op??'between');
  if(two&&!String(rule.f2??'').trim())return '최댓값을 입력하세요.';
  const a=operand(wb,si,rule,rule.f1,rule.r1,rule.c1)??0,b=two?(operand(wb,si,rule,rule.f2,rule.r1,rule.c1)??0):null;
  if(typeof a!=='number'||!Number.isFinite(a)||two&&(typeof b!=='number'||!Number.isFinite(b)))return '조건을 숫자·날짜·시간 또는 계산 가능한 수식으로 입력하세요.';
  if(two&&a>b)return '최댓값은 최솟값보다 크거나 같아야 합니다.';
  if(rule.type==='time'&&(a<0||a>=1||two&&(b<0||b>=1)))return '시간 조건은 0:00:00부터 23:59:59까지 입력하세요.';
  if(rule.type==='textLength'&&(a<0||!Number.isInteger(a)||two&&(b<0||!Number.isInteger(b))))return '텍스트 길이는 0 이상의 정수로 입력하세요.';
  return null;
}

/** 규칙 설명 (오류 메시지가 없을 때) */
export function describeRule(rule) {
  const t = VALIDATION_TYPES.find((x) => x.id === rule.type)?.label ?? rule.type;
  if (rule.type === 'list') return `목록에 있는 값만 입력할 수 있습니다.`;
  if (rule.type === 'custom') return '이 셀에 입력할 수 있는 값이 제한되어 있습니다.';
  const op = VALIDATION_OPS.find((x) => x.id === (rule.op ?? 'between'))?.label ?? '';
  const range = rule.op === 'between' || rule.op === 'notBetween' || !rule.op ? `${rule.f1 ?? ''} ~ ${rule.f2 ?? ''}` : `${op} ${rule.f1 ?? ''}`;
  return `${t} 값만 입력할 수 있습니다 (${rule.op === 'notBetween' ? '제외 범위 ' : ''}${range}).`;
}

/** 규칙 범위에서 rg 를 뺀 나머지 (최대 4조각) */
export function subtractRange(rule, rg) {
  if (rule.r2 < rg.r1 || rule.r1 > rg.r2 || rule.c2 < rg.c1 || rule.c1 > rg.c2) return [rule];
  const out = [];
  if (rule.r1 < rg.r1) out.push({ ...rule, r2: rg.r1 - 1 });
  if (rule.r2 > rg.r2) out.push({ ...rule, r1: rg.r2 + 1 });
  const r1 = Math.max(rule.r1, rg.r1);
  const r2 = Math.min(rule.r2, rg.r2);
  if (rule.c1 < rg.c1) out.push({ ...rule, r1, r2, c2: rg.c1 - 1 });
  if (rule.c2 > rg.c2) out.push({ ...rule, r1, r2, c1: rg.c2 + 1 });
  return out.map(part => relocateValidation(rule, part));
}

/** 범위 안에서 규칙에 맞지 않는 셀 (잘못된 데이터 표시) */
export function invalidCells(wb,si,limit=2000) {
  const sheet=wb.sheets[si],out=[],lists=new Map();
  if(!sheet||limit<=0)return out;
  const inspect=(cell,r,c)=>{
    if(out.length>=limit||!cell?.raw)return;
    const rule=validationAt(sheet,r,c);if(!rule||rule.type==='any')return;
    const text=!cell.formula&&typeof cell.v==='string'&&!cell.raw.startsWith("'")?"'"+cell.v:cell.raw;
    let knownList;
    if(rule.type==='list') {
      const f=formulaSource(rule,rule.f1);let key=!f?'literal:'+rule.f1:null;
      if(f)try {
        const shifted=shiftFormula(f,r-rule.r1,c-rule.c1),node=parse(shifted.slice(1));
        if(node.type==='ref')key=shifted;
        else if(node.type==='name') {
          const entry=wb.findName(node.v,si,node.sheet),ast=entry&&parse(String(entry.ref).replace(/^=/,'')),ref=ast?.type==='ref'?ast.ref:null;
          if(ref&&ref.ar1&&ref.ar2&&ref.ac1&&ref.ac2)key=shifted;
        }
      }catch{}
      if(key!==null) {knownList=lists.get(key);if(!knownList){knownList=validationList(wb,si,rule,r,c);if(lists.size<1000)lists.set(key,knownList);}}
    }
    if(!checkValidation(wb,si,rule,r,c,text,knownList))out.push({r,c});
  };
  sheet.cells.forEachRC(inspect);
  if(out.length>=limit)return out;
  for(const block of sheet.blocks??[])for(const rule of sheet.validations??[]) {
    if(rule.type==='any')continue;
    const r1=Math.max(block.r0,rule.r1),r2=Math.min(block.r0+block.n-1,rule.r2),c1=Math.max(block.c0,rule.c1),c2=Math.min(block.c0+block.cols.length-1,rule.c2);
    for(let c=c1;c<=c2;c++)for(let r=r1;r<=r2;r++) {
      if(sheet.cells.getRC(r,c)||validationAt(sheet,r,c)!==rule||blockValue(block,r,c)===null)continue;
      inspect(wb.getCell(si,r,c),r,c);if(out.length>=limit)return out;
    }
  }
  return out;
}
