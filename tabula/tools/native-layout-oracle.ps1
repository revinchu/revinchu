param(
  [Parameter(Mandatory=$true)][string]$InputPath,
  [Parameter(Mandatory=$true)][string]$TargetsJson,
  [Parameter(Mandatory=$true)][string]$OutputPath,
  [ValidateRange(1,512)][int]$MaxCellsPerSheet=64,
  [ValidateRange(1,512)][int]$MaxRowsPerSheet=32,
  [ValidateRange(1,512)][int]$MaxColumnsPerSheet=64,
  [ValidateRange(1,2048)][int]$MaxShapesPerSheet=512,
  [ValidateRange(1,32)][int]$MaxChartSeries=8,
  [ValidateRange(1,10)][int]$StartupRetries=5
)
# 원본은 열기만 합니다. 기존 Excel 프로세스에 attach/설정/종료하지 않습니다.
# 샘플 COM 서식·논리 geometry 기준이며 화면 픽셀 또는 WIXEL 전체 동등성 합격을 뜻하지 않습니다.
$ErrorActionPreference='Stop'
$script:themeUnavailable=0
function Release-Com($value) {
  if($null-ne$value -and [Runtime.InteropServices.Marshal]::IsComObject($value)) {
    try{[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($value)}catch{}
  }
}
function Invoke-Read([scriptblock]$Action,[int]$Retries=5) {
  for($attempt=0;$attempt-lt$Retries;$attempt++) {
    try{return ,(& $Action)}catch{
      $exception=$_.Exception;$busy=$false
      while($exception) {if($exception.HResult-in@(-2147418111,-2147417846)){$busy=$true};$exception=$exception.InnerException}
      if(!$busy -or $attempt-eq$Retries-1){throw}
      Start-Sleep -Milliseconds 1000
    }
  }
}
function Optional-Read([scriptblock]$Action,[System.Collections.Generic.List[string]]$Errors,[string]$Label) {
  try{return ,(Invoke-Read $Action $StartupRetries)}catch{
    # ThemeColor는 명시 RGB 색에서 Excel 문서상 1004 오류가 가능합니다. null을 0색으로 바꾸지 않습니다.
    $exception=$_.Exception;$notThemed=$false
    while($exception){if($exception.HResult-eq-2146827284){$notThemed=$true};$exception=$exception.InnerException}
    if($Label.EndsWith('.theme') -and $notThemed){$script:themeUnavailable++;return $null}
    $Errors.Add($Label+': '+$_.Exception.Message);return $null
  }
}
function Required-Number([scriptblock]$Action,[string]$Label) {
  $value=Invoke-Read $Action $StartupRetries
  $parsed=0.0
  if($null-eq$value -or $value-is[bool] -or ![double]::TryParse([string]$value,[ref]$parsed) -or [double]::IsNaN($parsed) -or [double]::IsInfinity($parsed)){throw ($Label+'의 유효한 숫자를 확인하지 못했습니다.')}
  return $parsed
}
function Fingerprint([string]$Path) {
  $file=Get-Item -LiteralPath $Path
  return [ordered]@{size=$file.Length;mtimeUtc=$file.LastWriteTimeUtc.ToString('o');sha256=(Get-FileHash -LiteralPath $Path -Algorithm SHA256).Hash.ToLowerInvariant()}
}
function Native-OpenPreflight([string]$Path) {
  $result=[ordered]@{status='verified';xlmMacros=$false;refreshOnLoad=$false;metadataParts=0;scope='XLM/refresh-on-load safety attributes only; no sheet/cache-record import'}
  $result.refreshMetadata=if([IO.Path]::GetExtension($Path)-ieq'.xlsb'){'binary-not-read'}else{'xml-attributes-read'}
  Add-Type -AssemblyName System.IO.Compression
  Add-Type -AssemblyName System.IO.Compression.FileSystem
  $archive=[IO.Compression.ZipFile]::OpenRead($Path)
  try {
    foreach($part in $archive.Entries) {
      if($part.FullName-match'^xl/macrosheets/'){$result.xlmMacros=$true}
      if($part.FullName-notmatch'^xl/(connections\.xml|queryTables/[^/]+\.xml|pivotCache/pivotCacheDefinition[^/]+\.xml)$'){continue}
      $result.metadataParts++;$stream=$null;$reader=$null
      try {
        $stream=$part.Open();$reader=[IO.StreamReader]::new($stream,[Text.Encoding]::UTF8,$true,8192,$true)
        $chars=New-Object char[] 8192;$carry=''
        for(;;){$count=$reader.Read($chars,0,$chars.Length);if(!$count){break};$text=$carry+[string]::new($chars,0,$count)
          if($text-match'\brefreshOnLoad\s*=\s*["''](?:1|true)["'']'){$result.refreshOnLoad=$true;break}
          $carry=if($text.Length-gt256){$text.Substring($text.Length-256)}else{$text}
        }
      }finally{if($reader){$reader.Dispose()};if($stream){$stream.Dispose()}}
    }
  }finally{$archive.Dispose()}
  if($result.xlmMacros){$result.status='xlm-macros-not-covered-by-force-disable'}
  return $result
}
function Excel-Inventory {
  $items=@()
  foreach($process in @(Get-Process -Name EXCEL -ErrorAction SilentlyContinue)) {
    $path=$null;$start=$null
    try{$path=$process.Path}catch{}
    try{$start=$process.StartTime.ToUniversalTime().ToString('o')}catch{}
    $items+=,[ordered]@{pid=$process.Id;path=$path;startUtc=$start}
  }
  return $items
}
function Property($Object,[string]$Key) {if($null-eq$Object){return $null};$entry=$Object.PSObject.Properties[$Key];if($entry){return $entry.Value};return $null}
function Read-Color($Color,[System.Collections.Generic.List[string]]$Errors,[string]$Prefix) {
  if($null-eq$Color){return $null}
  return [ordered]@{
    rgb=Optional-Read {$Color.RGB} $Errors ($Prefix+'.rgb')
    theme=Optional-Read {$Color.ObjectThemeColor} $Errors ($Prefix+'.theme')
    tint=Optional-Read {$Color.TintAndShade} $Errors ($Prefix+'.tint')
  }
}
function Read-Font($Font,[System.Collections.Generic.List[string]]$Errors) {
  if($null-eq$Font){return $null}
  return [ordered]@{
    name=Optional-Read {$Font.Name} $Errors 'font.name'
    size=Optional-Read {$Font.Size} $Errors 'font.size'
    bold=Optional-Read {$Font.Bold} $Errors 'font.bold'
    italic=Optional-Read {$Font.Italic} $Errors 'font.italic'
    underline=Optional-Read {$Font.Underline} $Errors 'font.underline'
    strike=Optional-Read {$Font.Strikethrough} $Errors 'font.strike'
    color=Optional-Read {$Font.Color} $Errors 'font.color'
    theme=Optional-Read {$Font.ThemeColor} $Errors 'font.theme'
    tint=Optional-Read {$Font.TintAndShade} $Errors 'font.tint'
  }
}
function Read-RangeStyle($Range,[System.Collections.Generic.List[string]]$Errors) {
  $font=$null;$fill=$null;$borders=$null
  try {
    $font=Invoke-Read {,($Range.Font)} $StartupRetries
    $fill=Invoke-Read {,($Range.Interior)} $StartupRetries
    $borders=Invoke-Read {,($Range.Borders)} $StartupRetries
    $edges=[ordered]@{}
    foreach($edge in @(@('diagonalDown',5),@('diagonalUp',6),@('left',7),@('top',8),@('bottom',9),@('right',10),@('insideVertical',11),@('insideHorizontal',12))) {
      $border=$null
      try {
        $border=Invoke-Read {,($borders.Item($edge[1]))} $StartupRetries
        $edges[$edge[0]]=[ordered]@{style=Optional-Read {$border.LineStyle} $Errors ('border.'+$edge[0]+'.style');weight=Optional-Read {$border.Weight} $Errors ('border.'+$edge[0]+'.weight');color=Optional-Read {$border.Color} $Errors ('border.'+$edge[0]+'.color');colorIndex=Optional-Read {$border.ColorIndex} $Errors ('border.'+$edge[0]+'.colorIndex');theme=Optional-Read {$border.ThemeColor} $Errors ('border.'+$edge[0]+'.theme');tint=Optional-Read {$border.TintAndShade} $Errors ('border.'+$edge[0]+'.tint')}
      }finally{Release-Com $border}
    }
    return [ordered]@{
      font=Read-Font $font $Errors
      fill=[ordered]@{color=Optional-Read {$fill.Color} $Errors 'fill.color';colorIndex=Optional-Read {$fill.ColorIndex} $Errors 'fill.colorIndex';pattern=Optional-Read {$fill.Pattern} $Errors 'fill.pattern';patternColor=Optional-Read {$fill.PatternColor} $Errors 'fill.patternColor';theme=Optional-Read {$fill.ThemeColor} $Errors 'fill.theme';tint=Optional-Read {$fill.TintAndShade} $Errors 'fill.tint'}
      borders=$edges
      numberFormat=Optional-Read {$Range.NumberFormat} $Errors 'format'
      alignment=[ordered]@{horizontal=Optional-Read {$Range.HorizontalAlignment} $Errors 'align.horizontal';vertical=Optional-Read {$Range.VerticalAlignment} $Errors 'align.vertical';wrap=Optional-Read {$Range.WrapText} $Errors 'align.wrap';shrink=Optional-Read {$Range.ShrinkToFit} $Errors 'align.shrink';indent=Optional-Read {$Range.IndentLevel} $Errors 'align.indent';orientation=Optional-Read {$Range.Orientation} $Errors 'align.orientation';readingOrder=Optional-Read {$Range.ReadingOrder} $Errors 'align.readingOrder'}
    }
  } finally {Release-Com $font;Release-Com $fill;Release-Com $borders}
}
function Read-Shape($Shape,[int]$Depth,[ref]$Budget) {
  if($Budget.Value-le0 -or $Depth-gt8){return [ordered]@{status='limited';depth=$Depth}}
  $Budget.Value--;$errors=[System.Collections.Generic.List[string]]::new();$line=$null;$fill=$null;$lineColor=$null;$fillColor=$null;$group=$null;$chart=$null
  $record=[ordered]@{depth=$Depth;errors=@()}
  try {
    foreach($key in @('Name','Type','Left','Top','Width','Height','Rotation','Visible','Placement','ZOrderPosition','AlternativeText','Title')) {
      $record[$key]=Optional-Read {$Shape.$key} $errors ('shape.'+$key)
    }
    $line=Optional-Read {$Shape.Line} $errors 'shape.line';$fill=Optional-Read {$Shape.Fill} $errors 'shape.fill'
    if($line) {
      $lineColor=Optional-Read {$line.ForeColor} $errors 'shape.line.color'
      $record.line=[ordered]@{visible=Optional-Read {$line.Visible} $errors 'shape.line.visible';weight=Optional-Read {$line.Weight} $errors 'shape.line.weight';dash=Optional-Read {$line.DashStyle} $errors 'shape.line.dash';style=Optional-Read {$line.Style} $errors 'shape.line.style';transparency=Optional-Read {$line.Transparency} $errors 'shape.line.transparency';color=Read-Color $lineColor $errors 'shape.line.color'}
    }
    if($fill) {
      $fillColor=Optional-Read {$fill.ForeColor} $errors 'shape.fill.color'
      $record.fill=[ordered]@{visible=Optional-Read {$fill.Visible} $errors 'shape.fill.visible';type=Optional-Read {$fill.Type} $errors 'shape.fill.type';transparency=Optional-Read {$fill.Transparency} $errors 'shape.fill.transparency';color=Read-Color $fillColor $errors 'shape.fill.color'}
    }
    if($record.Type-eq6) {
      $group=Optional-Read {,($Shape.GroupItems)} $errors 'shape.group';$record.children=@()
      if($group) {
        $count=Required-Number {$group.Count} 'group.Count';$record.groupCount=$count
        for($j=1;$j-le$count -and $Budget.Value-gt0;$j++) {$child=$null;try{$child=Invoke-Read {,($group.Item($j))} $StartupRetries;$record.children+=,(Read-Shape $child ($Depth+1) $Budget)}finally{Release-Com $child}}
        $record.groupLimited=($record.children.Count-lt$count)
      }
    }
    if((Optional-Read {$Shape.HasChart} $errors 'shape.hasChart')-eq-1) {
      $chart=Invoke-Read {,($Shape.Chart)} $StartupRetries;$series=$null
      try {
        $series=Invoke-Read {,($chart.SeriesCollection())} $StartupRetries
        $count=Required-Number {$series.Count} 'series.Count'
        $record.chart=[ordered]@{type=Optional-Read {$chart.ChartType} $errors 'chart.type';seriesCount=$count;series=@();seriesLimited=($count-gt$MaxChartSeries)}
        for($j=1;$j-le[Math]::Min($count,$MaxChartSeries);$j++) {
          $entry=$null;$format=$null;$seriesLine=$null;$seriesFill=$null;$seriesLineColor=$null;$seriesFillColor=$null
          try {
            $entry=Invoke-Read {,($series.Item($j))} $StartupRetries;$format=Invoke-Read {,($entry.Format)} $StartupRetries
            $seriesLine=Invoke-Read {,($format.Line)} $StartupRetries;$seriesFill=Invoke-Read {,($format.Fill)} $StartupRetries
            $seriesLineColor=Optional-Read {$seriesLine.ForeColor} $errors 'chart.series.line.color';$seriesFillColor=Optional-Read {$seriesFill.ForeColor} $errors 'chart.series.fill.color'
            $record.chart.series+=,[ordered]@{index=$j;type=Optional-Read {$entry.ChartType} $errors 'chart.series.type';line=[ordered]@{visible=Optional-Read {$seriesLine.Visible} $errors 'chart.series.line.visible';weight=Optional-Read {$seriesLine.Weight} $errors 'chart.series.line.weight';color=Read-Color $seriesLineColor $errors 'chart.series.line.color'};fill=[ordered]@{visible=Optional-Read {$seriesFill.Visible} $errors 'chart.series.fill.visible';color=Read-Color $seriesFillColor $errors 'chart.series.fill.color'}}
          }finally{Release-Com $seriesLineColor;Release-Com $seriesFillColor;Release-Com $seriesLine;Release-Com $seriesFill;Release-Com $format;Release-Com $entry}
        }
      }finally{Release-Com $series}
    }
  } catch {$errors.Add($_.Exception.Message)}
  finally {Release-Com $lineColor;Release-Com $fillColor;Release-Com $line;Release-Com $fill;Release-Com $group;Release-Com $chart}
  $record.errors=@($errors);$record.status=if($errors.Count){'partial'}else{'complete'}
  return $record
}

$source=[IO.Path]::GetFullPath($InputPath);$targetPath=[IO.Path]::GetFullPath($TargetsJson);$destination=[IO.Path]::GetFullPath($OutputPath)
if($source-eq$destination){throw '원본을 덮어쓸 수 없습니다.'}
if(Test-Path -LiteralPath $destination){throw '새 OutputPath를 사용하세요. 기존 결과는 덮어쓰지 않습니다.'}
if([IO.Path]::GetExtension($source).ToLowerInvariant()-notin@('.xlsx','.xlsm','.xlsb')){throw 'xlsx/xlsm/xlsb 원본만 읽습니다.'}
$targetData=Get-Content -LiteralPath $targetPath -Raw -Encoding UTF8|ConvertFrom-Json
$targets=@((Property $targetData 'sheets')|Where-Object {$null-ne$_})
if(!$targets.Count -or $targets.Count-gt256){throw 'TargetsJson.sheets에 1~256개 시트의 bounded 샘플이 필요합니다.'}
$outputFolder=[IO.Path]::GetDirectoryName($destination);$null=New-Item -ItemType Directory -Path $outputFolder -Force
# Add-Type의 임시 compiler 작업도 지정된 출력 위치에서만 수행합니다. 시스템 환경 설정은 바꾸지 않습니다.
$env:TEMP=$outputFolder;$env:TMP=$outputFolder
if(-not('WixelNativeOracleWindow'-as[type])) {
  Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;
public static class WixelNativeOracleWindow {
  [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr hwnd, out uint processId);
}
"@
}
$before=Fingerprint $source;$inventoryBefore=@(Excel-Inventory);$beforeIds=@($inventoryBefore|ForEach-Object {$_.pid});$started=[DateTime]::UtcNow
$record=[ordered]@{schemaVersion=1;input=$source;targets=$targetPath;started=$started.ToString('o');scope='Native Excel read-only bounded direct/DisplayFormat formats, borders, dimensions, merge and drawing geometry. No visual pixel test or WIXEL equivalence verdict.';inputBefore=$before;processesBefore=$inventoryBefore;ownedPid=$null;ownershipVerified=$false;opened=$false;readOnly=$false;sheets=@();errors=@();cleanup='not-started';limits=@{cellsPerSheet=$MaxCellsPerSheet;rowsPerSheet=$MaxRowsPerSheet;columnsPerSheet=$MaxColumnsPerSheet;shapesPerSheet=$MaxShapesPerSheet;chartSeries=$MaxChartSeries;groupDepth=8};nativeUnits='Width/height/left/top and line weight are points. ColumnWidth is Normal-font character units.'}
$xl=$null;$books=$null;$seed=$null;$book=$null;$owned=$false;$ownedPid=$null;$ownedStartTicks=$null;$expectedCount=0;$partial=$destination+'.partial.json'
function Write-ProgressRecord {$record|ConvertTo-Json -Depth 50|Set-Content -LiteralPath $partial -Encoding UTF8}
try {
  $record.openPreflight=Native-OpenPreflight $source
  if($record.openPreflight.status-ne'verified'){throw 'Excel4.0 XLM 매크로는 AutomationSecurity=3으로 막지 못하므로 native 인스턴스를 시작하지 않습니다.'}
  $xl=New-Object -ComObject Excel.Application
  $hwnd=Required-Number {$xl.Hwnd} 'Excel.Hwnd';[uint32]$instancePid=0
  [void][WixelNativeOracleWindow]::GetWindowThreadProcessId([IntPtr][int64]$hwnd,[ref]$instancePid)
  if(!$instancePid -or $instancePid-in$beforeIds){throw '새 인스턴스 PID 격리를 확인하지 못했습니다. 기존 Excel은 건드리지 않습니다.'}
  $process=Get-Process -Id $instancePid -ErrorAction Stop
  if(!$process.Path -or [IO.Path]::GetFileName($process.Path)-ine'EXCEL.EXE' -or $process.StartTime.ToUniversalTime()-lt$started.AddSeconds(-5)){throw '새 Excel 실행 경로/시작시각을 확인하지 못했습니다.'}
  $books=Invoke-Read {,($xl.Workbooks)} $StartupRetries
  if($null-eq$books -or ![Runtime.InteropServices.Marshal]::IsComObject($books)){throw 'Workbooks COM collection unavailable; isolation not verified'}
  $count=Required-Number {$books.Count} 'Workbooks.Count'
  if($count-ne0){throw '새 Excel 인스턴스가 비어 있지 않습니다. 기존 통합 문서는 건드리지 않습니다.'}
  $owned=$true;$ownedPid=[int]$instancePid;$ownedStartTicks=$process.StartTime.ToUniversalTime().Ticks
  $record.ownedPid=$ownedPid;$record.ownedProcess=[ordered]@{pid=$ownedPid;path=$process.Path;startUtc=$process.StartTime.ToUniversalTime().ToString('o');hwnd=$hwnd};$record.ownershipVerified=$true
  Write-ProgressRecord
  $xl.Visible=$false;$xl.DisplayAlerts=$false;$xl.EnableEvents=$false;$xl.AskToUpdateLinks=$false;$xl.ScreenUpdating=$false;$xl.AutomationSecurity=3
  $seed=Invoke-Read {,($books.Add(-4167))} $StartupRetries;$expectedCount=1
  $xl.Calculation=-4135;$xl.CalculateBeforeSave=$false
  $record.excel=[ordered]@{version=Invoke-Read {,($xl.Version)};build=Invoke-Read {,($xl.Build)};calculation=Invoke-Read {,($xl.Calculation)};automationSecurity=Invoke-Read {,($xl.AutomationSecurity)}}
  Write-Output ('격리된 소유 Excel PID '+$ownedPid+': 읽기 전용 열기')
  # Format/Origin/Delimiter/Converter는 XLSX/XLSM/XLSB에서 무시됩니다. PowerShell7 COM의 Type.Missing 바인딩을 피합니다.
  $book=Invoke-Read {,($books.Open($source,0,$true,5,'','',$true,2,',',$false,$false,0,$false,$true,0))} $StartupRetries
  $expectedCount=2
  if((Invoke-Read {,($book.ReadOnly)})-ne$true){throw 'ReadOnly 확인 실패'}
  $fullName=Invoke-Read {,($book.FullName)}
  if([IO.Path]::GetFullPath([string]$fullName)-ine$source){throw '열린 통합 문서의 경로가 원본과 다릅니다.'}
  $record.opened=$true;$record.readOnly=$true;$record.fullNameVerified=$true
  $worksheets=Invoke-Read {,($book.Worksheets)} $StartupRetries
  if($null-eq$worksheets -or ![Runtime.InteropServices.Marshal]::IsComObject($worksheets)){throw 'Worksheets COM collection unavailable'}
  try {
    $sheetCount=Required-Number {$worksheets.Count} 'Worksheets.Count';$record.worksheetCount=$sheetCount
    foreach($target in $targets) {
      $sheet=$null;$used=$null;$shapes=$null;$errors=[System.Collections.Generic.List[string]]::new();$entry=[ordered]@{cells=@();rows=@();columns=@();drawings=@();errors=@()}
      try {
        $requestedIndex=Property $target 'index';$requestedName=Property $target 'name'
        if($null-ne$requestedIndex) {
          if($requestedIndex-is[bool] -or [int]$requestedIndex-ne$requestedIndex -or $requestedIndex-lt0 -or $requestedIndex-ge$sheetCount){throw '시트 index 범위가 올바르지 않습니다.'}
          $sheet=Invoke-Read {,($worksheets.Item([int]$requestedIndex+1))} $StartupRetries;$entry.index=[int]$requestedIndex
        } elseif($requestedName) {$sheet=Invoke-Read {,($worksheets.Item([string]$requestedName))} $StartupRetries;$entry.index=(Required-Number {$sheet.Index} 'sheet.Index')-1}
        else {throw '시트 index 또는 name이 필요합니다.'}
        $entry.name=Invoke-Read {,($sheet.Name)};if($requestedName -and [string]$requestedName-cne[string]$entry.name){throw '샘플의 시트 이름/index가 원본과 다릅니다.'}
        $entry.visible=Invoke-Read {,($sheet.Visible)}
        $used=Invoke-Read {,($sheet.UsedRange)};$entry.usedAddress=Invoke-Read {,($used.Address($false,$false))}
        $cellTargets=@((Property $target 'cells')|Where-Object {$null-ne$_}|Select-Object -Unique);$entry.requestedCells=$cellTargets.Count;$entry.cellsLimited=($cellTargets.Count-gt$MaxCellsPerSheet)
        foreach($address in @($cellTargets|Select-Object -First $MaxCellsPerSheet)) {
          $cell=$null;$display=$null;$merge=$null;$cellErrors=[System.Collections.Generic.List[string]]::new();$item=[ordered]@{address=[string]$address;errors=@()}
          try {
            if([string]$address-notmatch'^[A-Z]{1,3}[1-9][0-9]{0,6}$'){throw '단일 A1 셀 주소가 필요합니다.'}
            $cell=Invoke-Read {,($sheet.Range([string]$address))} $StartupRetries
            $item.direct=Read-RangeStyle $cell $cellErrors
            $display=Invoke-Read {,($cell.DisplayFormat)} $StartupRetries;$item.display=Read-RangeStyle $display $cellErrors
            $item.geometry=[ordered]@{left=Required-Number {$cell.Left} 'cell.Left';top=Required-Number {$cell.Top} 'cell.Top';width=Required-Number {$cell.Width} 'cell.Width';height=Required-Number {$cell.Height} 'cell.Height'}
            $item.merged=Invoke-Read {,($cell.MergeCells)}
            if($item.merged){$merge=Invoke-Read {,($cell.MergeArea)};$item.mergeAddress=Invoke-Read {,($merge.Address($false,$false))}}
          }catch{$cellErrors.Add($_.Exception.Message)}finally{Release-Com $merge;Release-Com $display;Release-Com $cell}
          $item.errors=@($cellErrors);$item.status=if($cellErrors.Count){'partial'}else{'complete'};$entry.cells+=,$item
        }
        $rowTargets=@((Property $target 'rows')|Where-Object {$null-ne$_}|Select-Object -Unique);$entry.requestedRows=$rowTargets.Count;$entry.rowsLimited=($rowTargets.Count-gt$MaxRowsPerSheet)
        foreach($r in @($rowTargets|Select-Object -First $MaxRowsPerSheet)) {
          $range=$null
          try {
            if([int]$r-ne$r -or $r-lt1 -or $r-gt1048576){throw 'rows는 Excel 1기반 행 번호여야 합니다.'}
            $range=Invoke-Read {,($sheet.Rows.Item([int]$r))};$entry.rows+=,[ordered]@{row=[int]$r;rowHeight=Optional-Read {$range.RowHeight} $errors ('row.'+$r+'.height');height=Required-Number {$range.Height} 'row.Height';top=Required-Number {$range.Top} 'row.Top';hidden=Invoke-Read {,($range.Hidden)}}
          }catch{$errors.Add('row '+$r+': '+$_.Exception.Message)}finally{Release-Com $range}
        }
        $columnTargets=@((Property $target 'columns')|Where-Object {$null-ne$_}|Select-Object -Unique);$entry.requestedColumns=$columnTargets.Count;$entry.columnsLimited=($columnTargets.Count-gt$MaxColumnsPerSheet)
        foreach($c in @($columnTargets|Select-Object -First $MaxColumnsPerSheet)) {
          $range=$null
          try {
            if([string]$c-notmatch'^[A-Z]{1,3}$'){throw 'columns는 A~XFD의 열 이름이어야 합니다.'}
            $range=Invoke-Read {,($sheet.Columns.Item([string]$c))};$entry.columns+=,[ordered]@{column=[string]$c;columnWidth=Required-Number {$range.ColumnWidth} 'column.ColumnWidth';width=Required-Number {$range.Width} 'column.Width';left=Required-Number {$range.Left} 'column.Left';hidden=Invoke-Read {,($range.Hidden)}}
          }catch{$errors.Add('column '+$c+': '+$_.Exception.Message)}finally{Release-Com $range}
        }
        $shapes=Invoke-Read {,($sheet.Shapes)};$shapeCount=Required-Number {$shapes.Count} 'Shapes.Count';$entry.shapeCount=$shapeCount;$budget=$MaxShapesPerSheet
        for($j=1;$j-le$shapeCount -and $budget-gt0;$j++) {$shape=$null;try{$shape=Invoke-Read {,($shapes.Item($j))};$entry.drawings+=,(Read-Shape $shape 0 ([ref]$budget))}finally{Release-Com $shape}}
        $entry.shapesLimited=($entry.drawings.Count-lt$shapeCount);$entry.shapeNodesRead=$MaxShapesPerSheet-$budget
      }catch{$errors.Add($_.Exception.Message)}
      finally{Release-Com $shapes;Release-Com $used;Release-Com $sheet}
      $entry.errors=@($errors);$entry.status=if($errors.Count -or $entry.cellsLimited -or $entry.rowsLimited -or $entry.columnsLimited -or $entry.shapesLimited -or @($entry.cells|Where-Object {$_.status-ne'complete'}).Count -or @($entry.drawings|Where-Object {$_.status-ne'complete'}).Count){'partial'}else{'complete'}
      $record.sheets+=,$entry;Write-ProgressRecord
    }
  }finally{Release-Com $worksheets}
} catch {$record.errors+=,$_.Exception.Message;$record.failureLocation=[ordered]@{line=$_.InvocationInfo.ScriptLineNumber;position=$_.InvocationInfo.PositionMessage;stack=$_.ScriptStackTrace}}
finally {
  # PID/시작시각/가시성/Workbook 수가 그대로일 때만 소유 인스턴스를 닫습니다.
  $canClose=$false
  if($owned -and $xl) {
    try {
      $hwnd=Required-Number {$xl.Hwnd} 'cleanup.Hwnd';[uint32]$currentPid=0;[void][WixelNativeOracleWindow]::GetWindowThreadProcessId([IntPtr][int64]$hwnd,[ref]$currentPid)
      $process=Get-Process -Id $ownedPid -ErrorAction Stop
      $currentCount=Required-Number {$books.Count} 'cleanup.Workbooks.Count'
      $canClose=($currentPid-eq$ownedPid -and $ownedPid-notin$beforeIds -and $process.StartTime.ToUniversalTime().Ticks-eq$ownedStartTicks -and [IO.Path]::GetFileName($process.Path)-ieq'EXCEL.EXE' -and (Invoke-Read {,($xl.Visible)})-eq$false -and $currentCount-eq$expectedCount)
    }catch{$record.errors+=,('cleanup identity: '+$_.Exception.Message)}
  }
  if($canClose) {
    try {
      if($book){[void](Invoke-Read {,($book.Close($false))})};if($seed){[void](Invoke-Read {,($seed.Close($false))})}
      if((Required-Number {$books.Count} 'cleanup.remaining')-ne0){throw '닫기 후 새 통합 문서가 남아 Quit을 거절했습니다.'}
      [void](Invoke-Read {,($xl.Quit())});$record.cleanup='owned-instance-closed'
    }catch{$record.cleanup='retained';$record.errors+=,('cleanup: '+$_.Exception.Message)}
  } elseif($owned) {$record.cleanup='retained-ownership-or-user-state-changed';$record.errors+=,'인스턴스 상태가 바뀌어 종료하지 않았습니다.'}
  else {$record.cleanup='unowned-instance-not-touched'}
  Release-Com $book;Release-Com $seed;Release-Com $books;Release-Com $xl
  [GC]::Collect();[GC]::WaitForPendingFinalizers()
  $record.processesAfter=@(Excel-Inventory)
  try{$record.inputAfter=Fingerprint $source;$record.originalUnchanged=($record.inputAfter.size-eq$before.size -and $record.inputAfter.mtimeUtc-eq$before.mtimeUtc -and $record.inputAfter.sha256-eq$before.sha256)}catch{$record.originalUnchanged=$false;$record.errors+=,('원본 확인: '+$_.Exception.Message)}
  $record.status=if(!$record.opened){'unavailable'}elseif(!$record.originalUnchanged -or $record.errors.Count -or @($record.sheets|Where-Object {$_.status-ne'complete'}).Count){'partial'}else{'complete'}
  $record.themeUnavailableCount=$script:themeUnavailable;$record.themeNullPolicy='ThemeColor null은 비테마/조회불가 값이며 RGB0으로 취급하지 않습니다. resolved Color는 별도로 기록합니다.'
  $record.refreshPolicy='Workbooks.Open object-model contract: no automatic PivotCache/QueryTable/OLEDB/ODBC refresh; no Refresh method called.'
  $record.finished=[DateTime]::UtcNow.ToString('o')
  $record|ConvertTo-Json -Depth 50|Set-Content -LiteralPath $destination -Encoding UTF8
  Write-Output ([ordered]@{status=$record.status;opened=$record.opened;ownedPid=$record.ownedPid;sheets=$record.sheets.Count;originalUnchanged=$record.originalUnchanged;cleanup=$record.cleanup}|ConvertTo-Json -Compress)
}
if(!$record.opened -or !$record.originalUnchanged){exit 2}
if($record.status-ne'complete'){exit 1}