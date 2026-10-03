import { jsonReplacer } from './block.js';

/** 큰 캐시의 숫자 배열을 JSON 숫자 키 객체로 수백 MB 복사하지 않는다.
 * 열 하나씩 인코딩하며 빈칸(NaN), 숫자, 문자열, 오류를 그대로 보존한다. */
export function* pivotSnapshotJsonParts(snapshots) {
  const entries = Object.entries(snapshots);
  if (!entries.length) return;
  yield ',"pivotSnapshots":{';
  let first = true;
  for (const [id, data] of entries) {
    yield (first ? '' : ',') + JSON.stringify(id) + ':';
    first = false;
    if (data?.kind === 'pivot-cache') {
      yield JSON.stringify({ kind: data.kind, version: data.version, header: data.header, n: data.n }).slice(0, -1) + ',"columns":[';
      for (let i = 0; i < data.columns.length; i++) yield (i ? ',' : '') + JSON.stringify(data.columns[i], jsonReplacer);
      yield ']}';
    } else {
      yield '[';
      for (let i = 0; i < data.length; i += 512) yield (i ? ',' : '') + data.slice(i, i + 512).map(row => JSON.stringify(row)).join(',');
      yield ']';
    }
  }
  yield '}';
}
