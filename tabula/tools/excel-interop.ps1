param([Parameter(Mandatory=$true)][string]$FixtureRoot,[string]$OutputRoot)
# Windows + installed Excel. Runs only files in a generated synthetic fixture manifest.
# Uses a separate empty COM instance, disables macros/events/link updates, never repairs a file.
$ErrorActionPreference='Stop'
$FixtureRoot=[IO.Path]::GetFullPath($FixtureRoot)
if(!$OutputRoot){$OutputRoot=Join-Path $FixtureRoot 'excel-saved'}
$OutputRoot=[IO.Path]::GetFullPath($OutputRoot)
if($OutputRoot -eq $FixtureRoot){throw '출력 폴더는 합성 입력 폴더와 달라야 합니다.'}
$manifest=Get-Content -LiteralPath (Join-Path $FixtureRoot 'synthetic-fixtures.json') -Raw|ConvertFrom-Json
if($manifest.format -ne 'wixel-excel-interop' -or $manifest.version -ne 1){throw '합성 검증 목록이 아닙니다.'}
$null=New-Item -ItemType Directory -Path $OutputRoot -Force
$xl=$null;$book=$null;$owned=$false;$results=@();$fatal=$null
try {
 $xl=New-Object -ComObject Excel.Application
 if($xl.Workbooks.Count -ne 0){throw '별도 빈 Excel 인스턴스가 아닙니다.'};$owned=$true
 $xl.Visible=$false;$xl.DisplayAlerts=$false;$xl.EnableEvents=$false;$xl.AutomationSecurity=3;$xl.AskToUpdateLinks=$false
 foreach($case in $manifest.fixtures){
  if($case.file -notmatch '^[a-zA-Z0-9_-]+\.xlsx$'){throw '합성 파일 이름을 확인하세요.'}
  $record=[ordered]@{file=$case.file;opened=$false;version=$xl.Version;build=$xl.Build}
  try {
   $book=$xl.Workbooks.Open((Join-Path $FixtureRoot $case.file),0,$true)
   for($readyAttempt=0;$readyAttempt -lt 30;$readyAttempt++){
    try{if($xl.Ready -and $book.Worksheets.Count -gt 0){break}}catch{}
    Start-Sleep -Milliseconds 100
   }
   if($book.Worksheets.Count -lt 1){throw 'Excel 통합 문서 준비 상태를 확인하지 못했습니다.'}
   $record.opened=$true;$record.sheets=$book.Sheets.Count
   $sh=$book.Worksheets.Item(1);$chartTypes=@();$seriesCounts=@();$seriesTypes=@();$axisGroups=@()
   foreach($co in $sh.ChartObjects()){
    $chartTypes+=$co.Chart.ChartType;$sc=$co.Chart.SeriesCollection();$seriesCounts+=$sc.Count
    $types=@();$axes=@()
    for($i=1;$i -le $sc.Count;$i++){$se=$sc.Item($i);$types+=$se.ChartType;try{$axes+=$se.AxisGroup}catch{$axes+=$null}}
    $seriesTypes+=,@($types);$axisGroups+=,@($axes)
   }
   $record.chartTypes=$chartTypes;$record.seriesCounts=$seriesCounts;$record.seriesTypes=$seriesTypes;$record.axisGroups=$axisGroups
   foreach($field in @('seriesCounts','seriesTypes','axisGroups')){
    if($case.PSObject.Properties.Name -contains $field){
     $actual=ConvertTo-Json -InputObject @($record[$field]) -Depth 5 -Compress
     $expected=ConvertTo-Json -InputObject @($case.$field) -Depth 5 -Compress
     if($actual -ne $expected){throw ($field+' 불일치: '+$actual+' 예상 '+$expected)}
    }
   }
   if($case.PSObject.Properties.Name -contains 'chartTypes'){
    if(($chartTypes -join ',') -ne ($case.chartTypes -join ',')){throw ('차트 종류 불일치: '+($chartTypes -join ',')+' 예상 '+($case.chartTypes -join ','))}
   }
   if($case.shape){$shape=$sh.Shapes.Item(1);if([Math]::Abs($shape.Width-$case.shape.width) -gt .1 -or [Math]::Abs($shape.Height-$case.shape.height) -gt .1){throw '그림·차트 크기가 바뀌었습니다.'}}
   if($case.PSObject.Properties.Name -contains 'date1904'){if($book.Date1904 -ne $case.date1904){throw '날짜 체계 불일치'}}
   if($case.PSObject.Properties.Name -contains 'tables'){if($sh.ListObjects.Count -ne $case.tables){throw '표 개수 불일치'}}
   if($case.PSObject.Properties.Name -contains 'pivots'){
    $psh=if($case.pivotSheet){$book.Worksheets.Item($case.pivotSheet)}else{$sh}
    $pivots=$psh.PivotTables();$record.pivots=$pivots.Count
    if($pivots.Count -ne $case.pivots){throw '피벗 테이블 개수 불일치'}
    if($case.refreshPivots){for($pi=1;$pi -le $pivots.Count;$pi++){if(!$pivots.Item($pi).RefreshTable()){throw '피벗 새로 고침 실패'}};$record.pivotsRefreshed=$true}
    foreach($check in $case.pivotChecks){
     $pt=$pivots.Item($check.name);$pairs=@($check.pairs);$df=$check.dataField
     $range=switch($pairs.Count){
      0 {$pt.GetPivotData($df)}
      2 {$pt.GetPivotData($df,$pairs[0],$pairs[1])}
      4 {$pt.GetPivotData($df,$pairs[0],$pairs[1],$pairs[2],$pairs[3])}
      6 {$pt.GetPivotData($df,$pairs[0],$pairs[1],$pairs[2],$pairs[3],$pairs[4],$pairs[5])}
      8 {$pt.GetPivotData($df,$pairs[0],$pairs[1],$pairs[2],$pairs[3],$pairs[4],$pairs[5],$pairs[6],$pairs[7])}
      default {throw '피벗 검증은 최대 필드·항목 4쌍까지 지원합니다.'}
     }
     if($check.PSObject.Properties.Name -contains 'text'){if($range.Text -ne $check.text){throw ('피벗 표시 불일치: '+$df+' 실제='+$range.Text+' 예상='+$check.text)}}
     elseif($range.Value2 -ne $check.value){throw ('피벗 값 불일치: '+$df+' 실제='+$range.Value2+' 예상='+$check.value)}
    }
   }
   if($case.PSObject.Properties.Name -contains 'slicers'){if($book.SlicerCaches.Count -ne $case.slicers){throw '슬라이서 캐시 개수 불일치'}}
   foreach($cell in $case.cells){$actual=$sh.Range($cell.address).Value2;if($actual -ne $cell.value){throw ("셀 값 불일치: "+$cell.address+" 실제="+$actual+" 예상="+$cell.value)}}
   $book.SaveAs((Join-Path $OutputRoot $case.file),51);$record.saved=$true
  } catch {$record.error=$_.Exception.Message}
  finally {if($null -ne $book){try{$book.Close($false)}catch{$record.error='Excel 닫기 실패: '+$_.Exception.Message};try{[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($book)}catch{};$book=$null}}
  $results+=[pscustomobject]$record;[pscustomobject]$record|ConvertTo-Json -Depth 5 -Compress
 }
} catch {$fatal=$_.Exception.Message;Write-Output ('Excel 검증 중단: '+$fatal)} finally {if($null -ne $xl){if($owned){try{$xl.Quit()}catch{$fatal=$_.Exception.Message}};try{[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($xl)}catch{}}}
$results|ConvertTo-Json -Depth 6|Set-Content -Encoding utf8 (Join-Path $OutputRoot 'results.json')
if($fatal -or $results.Count -ne $manifest.fixtures.Count -or @($results|Where-Object{$_.error}).Count){exit 1}
exit 0
