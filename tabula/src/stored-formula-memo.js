// JSON chunks decode repeated formula text as separate string storage. Reuse
// only exactly equal immutable strings; cells, AST offsets and cached results
// remain independently owned. This per-sheet memo never grows with row count.
export function createStoredFormulaMemo({limit=16384,charLimit=2<<20}={}) {
  const values=new Map();let chars=0,hits=0;
  return {
    share(raw) {
      if(typeof raw!=='string'||raw.length<2||raw[0]!=='=')return raw;
      const saved=values.get(raw);
      if(saved!==undefined){hits++;return saved;}
      if(values.size<limit&&raw.length<=charLimit-chars){values.set(raw,raw);chars+=raw.length;}
      return raw;
    },
    clear(){values.clear();chars=0;},
    get stats(){return{entries:values.size,chars,hits};}
  };
}
