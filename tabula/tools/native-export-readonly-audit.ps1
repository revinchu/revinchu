param(
 [Parameter(Mandatory=$true)][string]$Path,
 [Parameter(Mandatory=$true)][string]$Output,
 [string]$Baseline,
 [string]$Id='F20-export',
 [switch]$RepairDiagnostic,
 [string]$AdditionalAddressesCsv='',
 [int]$SampleSheetIndex=-1,
 [switch]$Contended,
 [string]$AdditionalSheetAddressesJson='',
 [switch]$PivotFieldItems
)
# Downloaded audit copies only. Never opens, repairs or saves the business original. Optional repair applies only to the downloaded copy, which is never saved.
$ErrorActionPreference='Stop'
. (Join-Path $PSScriptRoot 'native-audit-structure.ps1')
$inputPath=[IO.Path]::GetFullPath($Path)
$outputPath=[IO.Path]::GetFullPath($Output)
if(!$inputPath.StartsWith('D:\Codex\Temp\wixel-final-audit\',[StringComparison]::OrdinalIgnoreCase)){throw '검수 사본 폴더의 입력만 허용합니다.'}
if(!$outputPath.StartsWith('D:\',[StringComparison]::OrdinalIgnoreCase)){throw '출력은 D:에만 저장합니다.'}
if(!(Test-Path -LiteralPath $inputPath -PathType Leaf)){throw '검수 사본이 없습니다.'}
$null=New-Item -ItemType Directory -Path ([IO.Path]::GetDirectoryName($outputPath)) -Force
if($SampleSheetIndex -lt -1){throw '추가 표본 시트 번호는 -1(전체) 또는 0 이상이어야 합니다.'}
$additionalAddresses=@($AdditionalAddressesCsv.Split(',')|ForEach-Object {$_.Trim().ToUpperInvariant()}|Where-Object {$_}|Select-Object -Unique)
foreach($address in $additionalAddresses){if($address -notmatch '^[A-Z]{1,3}[1-9][0-9]{0,6}$'){throw '추가 표본은 단일 셀 주소만 허용합니다.'}}
$additionalSheetAddresses=@{}
if($AdditionalSheetAddressesJson){
 $additionalSheetAddresses=ConvertFrom-Json -InputObject $AdditionalSheetAddressesJson -AsHashtable
 if($additionalSheetAddresses -isnot [Collections.IDictionary]){throw '추가 시트 표본은 시트 번호별 셀 주소 목록이어야 합니다.'}
 foreach($sampleSheet in $additionalSheetAddresses.Keys){
  if([string]$sampleSheet -notmatch '^(0|[1-9][0-9]*)$'){throw '추가 시트 번호는 0 이상 정수여야 합니다.'}
  foreach($address in @($additionalSheetAddresses[$sampleSheet])){if([string]$address -notmatch '^[A-Za-z]{1,3}[1-9][0-9]{0,6}$'){throw '추가 시트 표본은 단일 셀 주소만 허용합니다.'}}
 }
}
$baselineData=if($Baseline){Get-Content -LiteralPath $Baseline -Raw -Encoding utf8|ConvertFrom-Json}else{$null}
$before=Get-Item -LiteralPath $inputPath
$record=[ordered]@{id=$Id;path=$inputPath;mode=$(if($RepairDiagnostic){'repair-diagnostic-only'}else{'normal'});repairLogs=@();bytes=$before.Length;opened=$false;readOnly=$null;repairMode=$null;repairModeEvidence='unsupported-property-not-used';corruptLoad=$(if($RepairDiagnostic){1}else{0});readinessVerified=$false;readyWorksheetCount=$null;fullNameMatches=$false;normalOpenCompleted=$false;calculation=$null;sourceUnchanged=$null;sheets=@();pivotCaches=@();slicers=@();errors=@();sampleDifferences=@();structuralDifferences=@();structuralAudit=$null;preservationMatches=$null;contended=[bool]$Contended;timingNote=$(if($Contended){'다른 경량 검증과 CPU 경합 가능. 성능 비교 근거에서 제외.'}else{'독점 슬롯 실행. 별도 전체 프로세스 피크 계측은 아님.'});additionalSampleSheetIndex=$SampleSheetIndex;additionalSampleAddresses=$additionalAddresses;additionalSheetAddresses=$additionalSheetAddresses}
$xl=$null;$book=$null;$seed=$null;$owned=$false;$started=Get-Date
function Release($object){if($null -ne $object -and [Runtime.InteropServices.Marshal]::IsComObject($object)){try{[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($object)}catch{}}}
function Attempt($work){try{& $work}catch{$null}}
function Busy($error){return $error.Exception.HResult -in @(-2147418111,-2147417846,-2146777998) -or $error.Exception.Message -match 'RPC_E_CALL_REJECTED|RPC_E_SERVERCALL_RETRYLATER|0x80010001|0x8001010A|호출을 거부'}
function ReadReady($work){for($retry=0;;$retry++){try{return & $work}catch{if($retry-ge 40 -or !(Busy $_)){throw};Start-Sleep -Milliseconds 250}}}
function WaitNativeWorkbook($workbook){
 # Open is called once. Only properties of that returned workbook are retried.
 for($retry=0;$retry-lt 100;$retry++){
  try{
   $count=$workbook.Worksheets.Count;$readOnly=$workbook.ReadOnly;$fullName=[string]$workbook.FullName
   $samePath=$fullName.Equals($inputPath,[StringComparison]::OrdinalIgnoreCase)
   if($null-ne$count -and [int]$count-ge1 -and $readOnly-eq$true -and $samePath){return @{count=[int]$count;readOnly=$true;fullNameMatches=$true}}
  }catch{if(!(Busy $_)){throw}}
  Start-Sleep -Milliseconds 500
 }
 throw '통합 문서가 50초 안에 올바른 경로의 읽기 전용 준비 상태를 반환하지 않았습니다.'
}

function SaveProgress{$record|ConvertTo-Json -Depth 20|Set-Content -LiteralPath $outputPath -Encoding utf8}
function IsEqual($a,$b){if($null -eq $a -and $null -eq $b){return $true};if($null -eq $a -or $null -eq $b){return $false};if($a -is [ValueType] -and $b -is [ValueType]){return $a -eq $b};return ($a.GetType().Name -eq $b.GetType().Name -and [string]$a -ceq [string]$b)}
try{
 $xl=New-Object -ComObject Excel.Application
 if($xl.Workbooks.Count -ne 0){throw '별도 빈 Excel 인스턴스가 아닙니다.'};$owned=$true
 $xl.Visible=$false;$xl.DisplayAlerts=$false;$xl.EnableEvents=$false;$xl.AutomationSecurity=3;$xl.AskToUpdateLinks=$false;$xl.ScreenUpdating=$false
 $seed=$xl.Workbooks.Add();$xl.Calculation=-4135;$xl.CalculateBeforeSave=$false
 $record.excel=@{version=$xl.Version;build=$xl.Build;hwnd=$xl.Hwnd}
 SaveProgress
 Write-Output ($Id+' 읽기 전용 수동계산 열기 시작')
 $openStarted=Get-Date
 # Workbook.RepairMode is not an Excel property. Record the supported Open
 # policy explicitly; diagnostic repair output can never certify normal load.
 # Explicit optional values also avoid PowerShell COM Type.Missing ambiguity.
 $book=$xl.Workbooks.Open($inputPath,0,$true,5,'','',$true,2,'',$false,$false,0,$false,$false,$record.corruptLoad)
 $ready=WaitNativeWorkbook $book
 $record.openMs=[math]::Round(((Get-Date)-$openStarted).TotalMilliseconds)
 $record.opened=$true;$record.readOnly=$ready.readOnly;$record.readinessVerified=$true;$record.readyWorksheetCount=$ready.count;$record.fullNameMatches=$ready.fullNameMatches
 $record.calculation=$xl.Calculation;$record.saved=$book.Saved;$record.fileFormat=$book.FileFormat
 ReadReady {$seed.Close($false)};Release $seed;$seed=$null
 $allCaches=$book.PivotCaches()
 for($cacheIndex=1;$cacheIndex -le $allCaches.Count;$cacheIndex++){
  $cache=$allCaches.Item($cacheIndex)
  try{$record.pivotCaches+=@{index=$cacheIndex;recordCount=Attempt {$cache.RecordCount};sourceType=Attempt {$cache.SourceType};olap=Attempt {$cache.OLAP};refreshOnLoad=Attempt {$cache.RefreshOnFileOpen}}}finally{Release $cache}
 };Release $allCaches
 for($i=1;$i -le $book.Worksheets.Count;$i++){
  $sh=$book.Worksheets.Item($i)
  try{
   $row=[ordered]@{index=$i-1;name=[string]$sh.Name;visible=[int]$sh.Visible;tables=$sh.ListObjects.Count;autoFilterMode=Attempt {$sh.AutoFilterMode};filterMode=Attempt {$sh.FilterMode};tableFilters=@();pivots=@();shapes=@();sample=@()}
   $tables=$sh.ListObjects
   for($j=1;$j -le $tables.Count;$j++){
    $table=$tables.Item($j);$af=$null
    try{
     $af=$table.AutoFilter;$on=@()
     for($k=1;$k -le $af.Filters.Count;$k++){
      $filter=$af.Filters.Item($k)
      try{if($filter.On){$on+=@{index=$k-1;operator=Attempt {$filter.Operator};criteria1=Attempt {$filter.Criteria1};criteria2=Attempt {$filter.Criteria2}}}}finally{Release $filter}
     }
     $row.tableFilters+=@{name=[string]$table.Name;range=$table.Range.Address();showAutoFilter=$table.ShowAutoFilter;filterMode=Attempt {$af.FilterMode};active=$on}
    }finally{Release $af;Release $table}
   };Release $tables
   $pivots=$sh.PivotTables()
   for($j=1;$j -le $pivots.Count;$j++){
    $pv=$pivots.Item($j);$cache=$null
    try{
     $cache=$pv.PivotCache()
     $pivotRecord=@{name=[string]$pv.Name;range=$pv.TableRange2.Address();recordCount=Attempt {$cache.RecordCount};refreshOnLoad=Attempt {$cache.RefreshOnFileOpen};olap=Attempt {$cache.OLAP}}
     if($PivotFieldItems){
      $pivotRecord.fieldItems=@();$fields=$pv.PivotFields()
      try{for($fieldIndex=1;$fieldIndex -le $fields.Count;$fieldIndex++){
       $field=$fields.Item($fieldIndex);$fieldItems=$null
       try{
        $fieldItems=$field.PivotItems();$items=@()
        if($fieldItems.Count -le 128){for($itemIndex=1;$itemIndex -le $fieldItems.Count;$itemIndex++){$fieldItem=$fieldItems.Item($itemIndex);try{$items+=@{name=[string]$fieldItem.Name;visible=Attempt {$fieldItem.Visible}}}finally{Release $fieldItem}}}
        $pivotRecord.fieldItems+=@{name=[string]$field.Name;orientation=Attempt {$field.Orientation};count=$fieldItems.Count;items=$items}
       }finally{Release $fieldItems;Release $field}
      }}finally{Release $fields}
     }
     $row.pivots+=,$pivotRecord
    }finally{Release $cache;Release $pv}
   };Release $pivots
   $shapes=$sh.Shapes
   for($j=1;$j -le $shapes.Count;$j++){$shape=$shapes.Item($j);$row.shapes+=@{type=[int]$shape.Type;groups=Attempt {$shape.GroupItems.Count}};Release $shape};Release $shapes
   $reference=@($baselineData.sheets|Where-Object {$_.index -eq ($i-1)})
   $sheetAdditionalAddresses=@(if($SampleSheetIndex -eq -1 -or $SampleSheetIndex -eq ($i-1)){$additionalAddresses}else{@()})+@($additionalSheetAddresses[[string]($i-1)]|Where-Object {$_})
   foreach($address in (@('A1','B1','A2','B2','C3','D5','H10','P20')+@($sheetAdditionalAddresses)|Select-Object -Unique)){
    $cell=$sh.Range($address)
    try{
     $sample=@{address=$address;value=Attempt {$cell.Value2};formula=Attempt {$cell.Formula};text=Attempt {$cell.Text};font=Attempt {@{name=$cell.Font.Name;size=$cell.Font.Size;bold=$cell.Font.Bold;italic=$cell.Font.Italic}};fill=Attempt {$cell.Interior.Color}}
     $row.sample+=,$sample
     $prior=@($reference.sample|Where-Object {$_.address -eq $address})
     if($prior.Count -eq 1){
      if(!(IsEqual $prior[0].value $sample.value)){$record.sampleDifferences+=@{sheet=$i-1;address=$address;field='value';native=$prior[0].value;exported=$sample.value}}
      if(!(IsEqual $prior[0].formula $sample.formula)){$record.sampleDifferences+=@{sheet=$i-1;address=$address;field='formula';native=$prior[0].formula;exported=$sample.formula}}
     }
    }finally{Release $cell}
   }
   $record.sheets+=,$row
   SaveProgress
   Write-Output ($Id+' sheet '+$i+'/'+$book.Worksheets.Count+' pivots='+$row.pivots.Count+' shapes='+$row.shapes.Count)
  }catch{$record.errors+=('sheet '+$i+': '+$_.Exception.Message)}finally{Release $sh}
 }
 $caches=$book.SlicerCaches
 for($i=1;$i -le $caches.Count;$i++){
  $cache=$caches.Item($i);$slicers=$cache.Slicers
  $itemsCount=Attempt {$cache.SlicerItems.Count};$selectedCount=Attempt {$cache.VisibleSlicerItems.Count};$sampleItems=@()
  if($itemsCount -le 128){for($k=1;$k -le $itemsCount;$k++){$item=$cache.SlicerItems.Item($k);try{$sampleItems+=@{index=$k-1;name=[string]$item.Name;selected=[bool]$item.Selected}}finally{Release $item}}}
  try{for($j=1;$j -le $slicers.Count;$j++){$sl=$slicers.Item($j);try{$record.slicers+=@{name=[string]$sl.Name;left=Attempt {$sl.Left};top=Attempt {$sl.Top};width=Attempt {$sl.Width};height=Attempt {$sl.Height};sheet=Attempt {[string]$sl.Shape.Parent.Name};connections=Attempt {$cache.PivotTables.Count};cache=[string]$cache.Name;items=$itemsCount;selectedCount=$selectedCount;sampleItems=$sampleItems;filterCleared=Attempt {$cache.FilterCleared}}}finally{Release $sl}}}finally{Release $slicers;Release $cache}
 };Release $caches
}catch{$record.errors+=$_.Exception.Message;$record.protectedViewCount=Attempt {$xl.ProtectedViewWindows.Count};$record.openWorkbookCount=Attempt {$xl.Workbooks.Count};$record.exception=@{hresult=$_.Exception.HResult;inner=Attempt {$_.Exception.InnerException.Message};innerHresult=Attempt {$_.Exception.InnerException.HResult}}}
finally{
 if($book){try{ReadReady {$book.Close($false)}}catch{$record.errors+='닫기: '+$_.Exception.Message};Release $book}
 if($seed){try{ReadReady {$seed.Close($false)}}catch{};Release $seed}
 if($xl){if($owned){try{ReadReady {$xl.Quit()}}catch{$record.errors+='종료: '+$_.Exception.Message}};Release $xl}
 [GC]::Collect();[GC]::WaitForPendingFinalizers()
 $after=Get-Item -LiteralPath $inputPath;$record.sourceUnchanged=($before.Length -eq $after.Length -and $before.LastWriteTimeUtc.Ticks -eq $after.LastWriteTimeUtc.Ticks)
 if($RepairDiagnostic){
  $logDirectory=Join-Path ([IO.Path]::GetDirectoryName($outputPath)) 'repair-logs'
  foreach($tempPath in @($env:TEMP,([IO.Path]::Combine($env:LOCALAPPDATA,'Temp')))|Select-Object -Unique){
   if(Test-Path -LiteralPath $tempPath){Get-ChildItem -LiteralPath $tempPath -File -Filter 'error*.xml' -ErrorAction SilentlyContinue|Where-Object {$_.LastWriteTime -ge $started -and $_.Length -lt 4194304}|ForEach-Object {$null=New-Item -ItemType Directory -Path $logDirectory -Force;$destination=Join-Path $logDirectory $_.Name;Copy-Item -LiteralPath $_.FullName -Destination $destination;$record.repairLogs+=,$destination}}
  }
 }
 if($record.opened -and $baselineData.opened){$record.structuralAudit=Compare-NativeAuditStructure $baselineData $record;$record.structuralDifferences=@($record.structuralAudit.differences);$record.preservationMatches=($record.structuralAudit.matches -and $record.sampleDifferences.Count -eq 0)}
 $record.normalOpenCompleted=(!$RepairDiagnostic -and $record.corruptLoad-eq0 -and $record.readinessVerified -and $record.readOnly -and $record.fullNameMatches -and $record.readyWorksheetCount-eq$record.sheets.Count -and $record.errors.Count-eq0 -and $record.sourceUnchanged)
 $record.totalMs=[math]::Round(((Get-Date)-$started).TotalMilliseconds)
 SaveProgress
 [pscustomobject]@{id=$Id;opened=$record.opened;readOnly=$record.readOnly;corruptLoad=$record.corruptLoad;normalOpenCompleted=$record.normalOpenCompleted;repairModeEvidence=$record.repairModeEvidence;calculation=$record.calculation;sheets=$record.sheets.Count;pivots=(@($record.sheets|ForEach-Object {$_.pivots})).Count;slicers=$record.slicers.Count;sampleDifferences=$record.sampleDifferences.Count;structuralDifferences=$record.structuralDifferences.Count;preservationMatches=$record.preservationMatches;sourceUnchanged=$record.sourceUnchanged;openMs=$record.openMs;totalMs=$record.totalMs;errors=$record.errors}|ConvertTo-Json -Depth 4 -Compress
}
