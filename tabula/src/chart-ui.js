// 차트 대화상자의 데이터 안내와 유형별 검사. 렌더러와 DOM에 의존하지 않습니다.
export function chartDataGuide(chart, data) {
  const t = chart.type, n = data.categories?.length ?? 0, k = data.series?.length ?? 0;
  let description = '항목을 첫 열에, 계열 이름을 첫 행에 놓고 값의 크기를 비교합니다.', error = '';
  if (t === 'map') {
    description = '첫 열은 국가/지역 이름(한국어·영어) 또는 ISO 코드, 다음 열은 수치입니다. 같은 국가 값은 합산하며 도시·우편번호 지도는 지원하지 않습니다.';
  } else if (t === 'stock') {
    const order = [...(chart.volume ? ['거래량'] : []), ...(chart.ohlc ? ['시가'] : []), '고가', '저가', '종가'];
    description = `첫 열은 날짜, 다음 열은 ${order.join(' → ')} 순서로 배치하세요.`;
    if (k !== order.length) error = `이 주식형은 ${order.length}개 수치 계열이 필요합니다. 현재 ${k}개입니다.`;
  } else if (t === 'surface') {
    description = '첫 행·열은 두 축의 항목, 교차하는 셀은 높이 값입니다. 색은 높이 구간을 나타냅니다.';
    if (k < 2 || n < 2) error = '표면형에는 수치 계열 2개와 항목 2개 이상이 필요합니다.';
  } else if (t === 'sunburst' || t === 'treemap') {
    description = '지역 → 제품처럼 바깥 분류부터 세부 분류까지 여러 열에 놓고 마지막 열에 양수 값을 넣으세요.';
  } else if (t === 'pieOfPie' || t === 'barOfPie') {
    description = '첫 수치 계열의 일부 항목을 보조 원형·막대형으로 분리합니다. 차트 서식에서 분할 기준과 보조 크기를 조정하세요.';
    if (n < 3) error = '보조 차트로 나누려면 항목을 3개 이상 선택하세요.';
  } else if (t === 'scatter' || t === 'bubble') description = t === 'bubble' ? '열 순서: X값 → Y값 → 거품 크기. 추가 계열은 Y값·크기 열을 한 쌍으로 배치하세요.' : '첫 열은 X값, 다음 열은 계열별 Y값입니다. 축에는 실제 숫자 간격을 사용합니다.';
  else if (t === 'histogram' || t === 'pareto' || t === 'boxWhisker') description = '분포를 분석할 원자료를 수치 열로 선택하세요. 히스토그램의 구간 수·너비는 차트 서식에서 조정합니다.';
  else if (t === 'waterfall') description = '증가·감소 값을 순서대로 놓으세요. 차트 서식에서 시작·중간·마지막 항목을 합계로 지정할 수 있습니다.';
  else if (t === 'combo') description = '각 계열의 차트 종류와 기본 축(왼쪽)·보조 축(오른쪽)을 아래에서 선택하세요.';
  else if (t === 'pie' || t === 'doughnut') description = t === 'pie' ? '첫 수치 계열이 전체에서 차지하는 비율을 표시합니다.' : '각 수치 계열을 별도 고리로 표시합니다.';
  return { description, error };
}

export function chartPresetMatches(chart, preset) {
  const defaults = { threeD: false, grouping: 'clustered', explode: 0, ohlc: false, volume: false, radarStyle: 'standard', scatterStyle: 'marker', comboAxis: 'secondary', bubble3D: false, surfaceStyle: 'surface', marker: 'none' };
  return Object.entries(preset).every(([key, value]) => (chart[key] ?? defaults[key]) === value);
}
