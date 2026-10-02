// 후보 입력을 실제 셀/파일 캐시/Undo에 기록하지 않는 유효성 검사 전용 계산 문맥.
import { parse, evaluateArray, evalAny, Range, RefValue, ERR, isError } from './formula.js';
import { parseInput } from './format.js';

export function validationEvaluation(wb, si, r, c, input) {
  const memo = new Map(), busy = new Set(), names = new Set();
  let reads = 0;
  const scalar = value => value instanceof Range ? value.rows[0]?.[0] ?? null : value;
  const attempt = fn => { try { return fn(); } catch (error) { return isError(error) ? error : ERR.VALUE; } };
  const get = (s, rr, cc) => {
    if (!wb.sheets[s] || rr < 0 || cc < 0) return ERR.REF;
    const key = `${s}:${rr},${cc}`;
    if (memo.has(key)) return scalar(memo.get(key));
    if (busy.has(key) || busy.size >= 100) return ERR.CIRC;
    if (++reads > 1000000) return ERR.VALUE;
    const candidate = input !== undefined && s === si && rr === r && cc === c;
    const cell = candidate ? null : wb.getCell(s, rr, cc);
    const raw = candidate ? String(input ?? '') : cell?.raw ?? '';
    if (candidate && !raw.startsWith('=')) { const parsed=parseInput(raw,wb.date1904);return parsed.errorLiteral ? Object.values(ERR).find(e=>e.code===parsed.errorLiteral)??ERR.VALUE : parsed.value; }
    if (!candidate && !cell?.formula) {
      // 이미 존재하는 분산 셀을 읽더라도 실제 통합 문서에 새 분산 값을 쓰지 않는다.
      const owner = wb.spillOwner?.get(key), spill = owner && wb.spills?.get(owner);
      if (!cell?.raw && spill) {
        get(s, spill.r, spill.c);
        const array = memo.get(owner);
        return array instanceof Range ? array.rows[rr - spill.r]?.[cc - spill.c] ?? null : array;
      }
      return cell?.v ?? null;
    }
    busy.add(key);
    const value = attempt(() => {
      const ctx = context(s, rr, cc);
      ctx.dr = candidate ? 0 : cell.dr ?? 0;
      ctx.dc = candidate ? 0 : cell.dc ?? 0;
      const ast = candidate ? parse(raw.slice(1)) : cell.ast;
      return ast ? evaluateArray(ast, ctx) : ERR.NAME;
    });
    busy.delete(key); memo.set(key, value);
    return scalar(value);
  };
  const context = (s, rr, cc) => {
    const ctx = wb.ctxFor(s, rr, cc);
    ctx.cell = (sheet, row, col) => get(wb.resolveSheet(sheet, s), row, col);
    ctx.range = (sheet, r1, c1, r2, c2) => {
      if ((r2-r1+1)*(c2-c1+1) > 1000000) throw ERR.VALUE;
      const rows = [], source = wb.resolveSheet(sheet, s);
      for (let row=r1; row<=r2; row++) {
        const values=[];
        for(let col=c1;col<=c2;col++) values.push(get(source,row,col));
        rows.push(values);
      }
      return rows;
    };
    ctx.formulaText = (sheet, row, col) => {
      const source = wb.resolveSheet(sheet, s);
      if(input !== undefined && source===si && row===r && col===c) return String(input).startsWith('=') ? String(input) : null;
      const cell=wb.getCell(source,row,col);return cell?.formula ? cell.raw : null;
    };
    ctx.name = (name, sheet) => {
      const entry = wb.findName(name,s,sheet);
      if(!entry) {
        const ref = !sheet && ctx.structRef?.(name,'');
        return ref && !ref.error ? new RefValue(ref.sheet,ref.r1,ref.c1,ref.r2,ref.c2) : undefined;
      }
      if(names.has(entry))return ERR.CIRC;
      names.add(entry);
      const scope=entry.sheet?wb.sheetIndexByName(entry.sheet):s;
      const value=attempt(()=>evalAny(parse(String(entry.ref??'').replace(/^=/,'')),context(scope,rr,cc)));
      names.delete(entry);
      return value instanceof RefValue&&!value.sheet ? new RefValue(wb.sheets[scope]?.name,value.r1,value.c1,value.r2,value.c2) : value;
    };
    ctx.spillRef = (sheet, row, col) => {
      const source=wb.resolveSheet(sheet,s);get(source,row,col);
      const array=memo.get(`${source}:${row},${col}`);
      return array instanceof Range ? new RefValue(wb.sheets[source].name,row,col,row+array.height-1,col+array.width-1) : null;
    };
    return ctx;
  };
  return { context:context(si,r,c), value:()=>get(si,r,c), attempt };
}
