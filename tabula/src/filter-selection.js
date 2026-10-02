const foldFilterText = value => String(value ?? '').toLocaleLowerCase();

// Excel filter search is a substring search; * and ? are wildcards, ~ escapes * ? ~.
// A greedy glob matcher avoids exponential RegExp backtracking on adversarial search text.
function compileFilterQuery(query) {
  const tokens=[],chars=Array.from(query);let wildcard=false,literal='';
  for(let i=0;i<chars.length;i++){
    const c=chars[i];
    if(c==='~'&&['*','?','~'].includes(chars[i+1])){tokens.push({text:chars[++i]});literal+=chars[i];}
    else if(c==='*'||c==='?'){wildcard=true;if(c!=='*'||tokens.at(-1)?.kind!=='*')tokens.push({kind:c});}
    else{tokens.push({text:c});literal+=c;}
  }
  if(!wildcard)return value=>value.includes(literal);
  if(tokens[0]?.kind!=='*')tokens.unshift({kind:'*'});
  if(tokens.at(-1)?.kind!=='*')tokens.push({kind:'*'});
  return value=>{
    const text=Array.from(value);let i=0,j=0,star=-1,retry=0;
    while(i<text.length){
      if(tokens[j]?.kind==='*'){star=j++;retry=i;}
      else if(tokens[j]&&(tokens[j].kind==='?'||tokens[j].text===text[i])){i++;j++;}
      else if(star>=0){j=star+1;i=++retry;}
      else return false;
    }
    while(tokens[j]?.kind==='*')j++;
    return j===tokens.length;
  };
}
export function filterSearchPredicate(query) {
  const matches=compileFilterQuery(foldFilterText(query).trim());
  return (value,label=value)=>matches(foldFilterText(value))||matches(foldFilterText(label));
}

// 검색 중 선택은 임시 상태다. 검색을 지우면 검색 전의 체크 선택을 복원한다.
export function filterSelection(items, initial = null, label = (value) => value) {
  const values = [...new Set(items)];
  const known = new Set(values), base = new Set([...(initial ?? values)].filter(value=>known.has(value)));
  const text = new Map(values.map((value) => [value, [foldFilterText(value),foldFilterText(label(value))]]));
  let query = '', visible = values, visibleSet = known, selected = base;
  return {
    search(value) {
      const next = foldFilterText(value).trim();
      if (next === query) return;
      const previousVisible = new Set(visible);
      const matches=compileFilterQuery(next);
      visible = next ? values.filter((item) => text.get(item).some(matches)) : values;
      visibleSet = next ? new Set(visible) : known;
      selected = next ? new Set(visible.filter((item) => !query || !previousVisible.has(item) || selected.has(item))) : base;
      query = next;
    },
    visible: () => visible,
    checked: (value) => selected.has(value),
    toggle(value, checked) { if (!visibleSet.has(value)) return; if (checked) selected.add(value); else selected.delete(value); },
    selectVisible(checked) { for (const value of visible) { if (checked) selected.add(value); else selected.delete(value); } },
    result(add = false) { return [...(query && add ? new Set([...base, ...selected]) : selected)]; },
    canApply: () => visible.length > 0 && selected.size > 0,
    searching: () => !!query,
  };
}
