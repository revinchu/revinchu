// 기존 숫자 키 JSON 형식을 유지하며 블록 전체 문자열/배열 복제를 피한다.
// 각 부분은 동기적으로 소비되므로 저장 도중 편집된 값이 섞이지 않는다.
function* arrayParts(value, chunk = 2048) {
  if (value == null) { yield 'null'; return; }
  yield '{';
  for (let i = 0; i < value.length; i += chunk) {
    const end = Math.min(value.length, i + chunk), parts = [];
    for (let j = i; j < end; j++) parts.push(JSON.stringify(String(j)) + ':' + JSON.stringify(value[j]));
    yield (i ? ',' : '') + parts.join(',');
  }
  yield '}';
}
export function* blockJsonParts(block) {
  yield JSON.stringify({r0:block.r0,c0:block.c0,n:block.n,ver:block.ver ?? 0}).slice(0,-1);
  if (block.perm) { yield ',"perm":'; yield* arrayParts(block.perm); }
  yield ',"cols":[';
  for (let i=0;i<block.cols.length;i++) {
    const col=block.cols[i];
    yield (i ? ',' : '') + '{"num":'; yield* arrayParts(col.num);
    yield ',"str":'; yield* arrayParts(col.str);
    yield ',"dict":[';
    for (let j=0;j<col.dict.length;j+=512) yield (j ? ',' : '') + col.dict.slice(j,j+512).map(v=>JSON.stringify(v) ?? 'null').join(',');
    yield '],"fmt":' + JSON.stringify(col.fmt ?? null) + '}';
  }
  yield ']}';
}
