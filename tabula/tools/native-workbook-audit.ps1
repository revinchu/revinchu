param([Parameter(Mandatory=$true)][string]$Manifest,[Parameter(Mandatory=$true)][string]$Id,[Parameter(Mandatory=$true)][string]$OutputRoot,[switch]$Render,[switch]$InspectSlicerItems,[switch]$TraceSteps,[switch]$InspectSmallValues)
# 사용자가 지정한 파일만 별도 Excel에서 읽기 전용으로 검사합니다. 원본 저장·복구·새로 고침·외부 링크 업데이트는 하지 않습니다.
$ErrorActionPreference='Stop'
if(Test-Path -LiteralPath (Join-Path $OutputRoot 'pause.native')){Write-Output ($Id+' 대기 표식으로 실행하지 않음');exit 0}
# 큰 파일 사이에만 작은 합성 검증을 실행하므로 Excel 메모리를 겹쳐 쓰지 않는다.
$probeQueue=Join-Path $OutputRoot 'pending-probes.json'
if($Id -match '^F[0-9]{2}$' -and (Test-Path -LiteralPath $probeQueue)){
 $probes=@(Get-Content -LiteralPath $probeQueue -Raw -Encoding UTF8|ConvertFrom-Json)
 Remove-Item -LiteralPath $probeQueue
 foreach($probe in $probes){& $PSCommandPath -Manifest $probe.manifest -Id $probe.id -OutputRoot 'D:/Codex/Temp/wixel-final-audit/native-independent' -InspectSlicerItems}
}

$cases=@(Get-Content -LiteralPath $Manifest -Raw -Encoding UTF8|ConvertFrom-Json)
$case=@($cases|Where-Object {$_.id -eq $Id})
if($case.Count -ne 1){throw '검사할 파일 ID가 없습니다.'}
$path=[IO.Path]::GetFullPath($case[0].path)
if(!(Test-Path -LiteralPath $path -PathType Leaf)){throw '입력 파일이 없습니다.'}
$OutputRoot=[IO.Path]::GetFullPath($OutputRoot)
if(!$OutputRoot.StartsWith('D:\',[StringComparison]::OrdinalIgnoreCase)){throw '검사 출력은 D:에 저장해야 합니다.'}
$null=New-Item -ItemType Directory -Path $OutputRoot -Force
$before=Get-Item -LiteralPath $path
$record=[ordered]@{id=$Id;path=$path;bytes=$before.Length;modifiedTicks=$before.LastWriteTimeUtc.Ticks;opened=$false;readOnly=$null;sheets=@();slicers=@();errors=@()}
$xl=$null;$book=$null;$owned=$false;$started=Get-Date
function Release($object){if($null -ne $object -and [Runtime.InteropServices.Marshal]::IsComObject($object)){try{[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($object)}catch{}}}
function Attempt($work,$fallback=$null){try{& $work}catch{$fallback}}
function Step($message){if($TraceSteps){($Id+' '+$message)|Add-Content -LiteralPath (Join-Path $OutputRoot ($Id+'-steps.log')) -Encoding utf8}}
function FieldInfo($fields){
 $out=@();for($i=1;$i -le $fields.Count;$i++){
  $f=$fields.Item($i);Step ('field '+$i+' name/orientation')
  $orientation=Attempt {[int]$f.Orientation};$info=@{name=Attempt {[string]$f.Name};caption=Attempt {[string]$f.Caption};orientation=$orientation;position=Attempt {$f.Position};numberFormat=Attempt {$f.NumberFormat}}
  # 해당 축에서 유효한 COM 속성만 조회한다. OLAP 행 필드에 CurrentPage/Function을 요청하지 않는다.
  if($orientation -eq 4){Step ('field '+$i+' data function');$info.function=Attempt {$f.Function};Step ('field '+$i+' data calculation');$info.calculation=Attempt {$f.Calculation}}
  if($orientation -eq 3){Step ('field '+$i+' currentPage');$info.currentPage=Attempt {[string]$f.CurrentPage};Step ('field '+$i+' multiple');$info.multiple=Attempt {$f.EnableMultiplePageItems}}
  $out+=,$info;Release $f
 };return ,$out
}
try{
 $xl=New-Object -ComObject Excel.Application
 if($xl.Workbooks.Count -ne 0){throw '별도 빈 Excel 인스턴스가 아닙니다.'};$owned=$true
 $xl.Visible=$false;$xl.DisplayAlerts=$false;$xl.EnableEvents=$false;$xl.AutomationSecurity=3;$xl.AskToUpdateLinks=$false
 $xl.ScreenUpdating=$false
 $record.excel=@{version=$xl.Version;build=$xl.Build}
 Write-Output ($Id+' 읽기 전용 열기 시작')
 $book=$xl.Workbooks.Open($path,0,$true)
 $record.opened=$true;$record.readOnly=$book.ReadOnly;$record.openMs=[math]::Round(((Get-Date)-$started).TotalMilliseconds)
 $record.saved=$book.Saved;$record.date1904=$book.Date1904;$record.calculation=$xl.Calculation;$record.names=$book.Names.Count
 $record.activeSheet=Attempt {$book.ActiveSheet.Name};$record.protectStructure=$book.ProtectStructure
 $window=Attempt {$book.Windows.Item(1)}
 $rendered=$false
 for($i=1;$i -le $book.Worksheets.Count;$i++){
  Step ("sheet $i UsedRange");$sh=$book.Worksheets.Item($i);$used=$sh.UsedRange
  try{
  $row=[ordered]@{index=$i-1;name=[string]$sh.Name;visible=[int]$sh.Visible;protected=[bool]$sh.ProtectContents;used=@{r1=[int]$used.Row-1;c1=[int]$used.Column-1;r2=[int]$used.Row+[int]$used.Rows.Count-2;c2=[int]$used.Column+[int]$used.Columns.Count-2};standardWidth=$sh.StandardWidth;standardHeight=$sh.StandardHeight;tables=@();pivots=@();shapes=@();rowHeights=@();colWidths=@();sample=@()}
  Step ("sheet $i filter/CF/page");$row.autoFilterMode=Attempt {$sh.AutoFilterMode};$row.filterMode=Attempt {$sh.FilterMode};$row.filterRange=Attempt {$sh.AutoFilter.Range.Address()}
  $row.cfCount=Attempt {$used.FormatConditions.Count}
  $row.page=Attempt {@{printArea=$sh.PageSetup.PrintArea;orientation=$sh.PageSetup.Orientation;paperSize=$sh.PageSetup.PaperSize;zoom=$sh.PageSetup.Zoom;fitWide=$sh.PageSetup.FitToPagesWide;fitTall=$sh.PageSetup.FitToPagesTall;repeatRows=$sh.PageSetup.PrintTitleRows;repeatCols=$sh.PageSetup.PrintTitleColumns}}
  Step ("sheet $i view");if($sh.Visible -eq -1){try{$sh.Activate();$row.view=@{zoom=$window.Zoom;freeze=$window.FreezePanes;splitRow=$window.SplitRow;splitCol=$window.SplitColumn;grid=$window.DisplayGridlines;headers=$window.DisplayHeadings;view=$window.View}}catch{$row.viewError=$_.Exception.Message}}
  Step ("sheet $i dimensions");for($c=1;$c -le [Math]::Min(64,$used.Column+$used.Columns.Count-1);$c++){$col=$sh.Columns.Item($c);$row.colWidths+=@{index=$c-1;chars=$col.ColumnWidth;points=$col.Width;hidden=$col.Hidden};Release $col}
  for($r=1;$r -le [Math]::Min(64,$used.Row+$used.Rows.Count-1);$r++){$rr=$sh.Rows.Item($r);$row.rowHeights+=@{index=$r-1;points=$rr.RowHeight;hidden=$rr.Hidden};Release $rr}
  Step ("sheet $i tables");$tables=$sh.ListObjects
  for($j=1;$j -le $tables.Count;$j++){$t=$tables.Item($j);$row.tables+=@{name=$t.Name;range=$t.Range.Address();style=[string]$t.TableStyle;filterButton=$t.ShowAutoFilter;headers=$t.ShowHeaders;totals=$t.ShowTotals;firstCol=$t.ShowTableStyleFirstColumn;lastCol=$t.ShowTableStyleLastColumn;rowStripes=$t.ShowTableStyleRowStripes;colStripes=$t.ShowTableStyleColumnStripes};Release $t};Release $tables
  Step ("sheet $i pivots");$pivots=$sh.PivotTables()
  for($j=1;$j -le $pivots.Count;$j++){
   Step ("sheet $i pivot $j details");$p=$pivots.Item($j);$row.pivots+=@{name=$p.Name;range=$p.TableRange2.Address();style=Attempt {[string]$p.TableStyle2};rowGrand=$p.RowGrand;colGrand=$p.ColumnGrand;fieldHeaders=Attempt {$p.DisplayFieldCaptions};preserveFormat=$p.PreserveFormatting;inGrid=Attempt {$p.InGridDropZones};pageWrap=Attempt {$p.PageFieldWrapCount};pageOrder=Attempt {$p.PageFieldOrder};rows=Attempt {FieldInfo ($p.RowFields())} @();columns=Attempt {FieldInfo ($p.ColumnFields())} @();pages=Attempt {FieldInfo ($p.PageFields())} @();values=Attempt {FieldInfo ($p.DataFields())} @();source=Attempt {$p.SourceData};cacheIndex=$p.CacheIndex;recordCount=Attempt {$p.PivotCache().RecordCount};refreshOnOpen=Attempt {$p.PivotCache().RefreshOnFileOpen};olap=Attempt {$p.PivotCache().OLAP}};Release $p
  };Release $pivots
  Step ("sheet $i shapes");$shapes=$sh.Shapes
  for($j=1;$j -le $shapes.Count;$j++){$shape=$shapes.Item($j);$row.shapes+=@{name=$shape.Name;type=$shape.Type;left=$shape.Left;top=$shape.Top;width=$shape.Width;height=$shape.Height;rotation=$shape.Rotation;visible=$shape.Visible;autoShapeType=Attempt {$shape.AutoShapeType};adjustments=Attempt {$a=$shape.Adjustments;$v=@();for($ai=1;$ai -le $a.Count;$ai++){$v+=,$a.Item($ai)};Release $a;,$v};chartType=Attempt {$shape.Chart.ChartType};groups=Attempt {$shape.GroupItems.Count}};Release $shape};Release $shapes
  Step ("sheet $i samples");$samples=@('A1','B1','A2','B2','C3','D5','H10','P20')
  foreach($address in $samples){$cell=$sh.Range($address);$row.sample+=@{address=$address;text=Attempt {$cell.Text};value=Attempt {$cell.Value2};formula=Attempt {$cell.Formula};font=Attempt {@{name=$cell.Font.Name;size=$cell.Font.Size;bold=$cell.Font.Bold;color=$cell.Font.Color}};format=Attempt {$cell.NumberFormat};fill=Attempt {$cell.Interior.Color};wrap=Attempt {$cell.WrapText};rotation=Attempt {$cell.Orientation};align=Attempt {$cell.HorizontalAlignment}};Release $cell}
  if($InspectSmallValues){$row.smallValues=@();for($vr=1;$vr -le [Math]::Min(16,$used.Row+$used.Rows.Count-1);$vr++){$line=@();for($vc=1;$vc -le [Math]::Min(8,$used.Column+$used.Columns.Count-1);$vc++){$valueCell=$sh.Cells.Item($vr,$vc);$line+=,$valueCell.Value2;Release $valueCell};$row.smallValues+=,$line}}
  if($Render -and !$rendered -and $sh.Visible -eq -1){
   try{$pdf=Join-Path $OutputRoot ($Id+'-excel-first-page.pdf');$sh.ExportAsFixedFormat(0,$pdf,0,$false,$false,1,1,$false);$row.pdf=$pdf;$rendered=$true}catch{$row.renderError=$_.Exception.Message}
  }
  $record.sheets+=,$row
  Write-Output ($Id+' sheet '+$i+'/'+$book.Worksheets.Count+' pivots='+$row.pivots.Count+' shapes='+$row.shapes.Count)
  }catch{$record.errors+=('sheet '+$i+': '+$_.Exception.Message)}
  Release $used;Release $sh
 }
 Step "slicer caches";$caches=$book.SlicerCaches
 for($i=1;$i -le $caches.Count;$i++){$cache=$caches.Item($i);$slicers=$cache.Slicers;for($j=1;$j -le $slicers.Count;$j++){$sl=$slicers.Item($j);$record.slicers+=@{name=$sl.Name;caption=$sl.Caption;cache=$cache.Name;sourceName=Attempt {$cache.SourceName};sheet=Attempt {$sl.Shape.Parent.Name};style=Attempt {[string]$sl.Style.Name};columns=$sl.NumberOfColumns;rowHeight=$sl.RowHeight;columnWidth=$sl.ColumnWidth;left=$sl.Left;top=$sl.Top;width=$sl.Width;height=$sl.Height;displayHeader=$sl.DisplayHeader;sort=Attempt {$cache.SortItems};crossFilter=Attempt {$cache.CrossFilterType};filterCleared=Attempt {$cache.FilterCleared};pivotConnectionCount=Attempt {$cache.PivotTables.Count};items=Attempt {if($InspectSlicerItems){$is=@();for($k=1;$k -le [Math]::Min(128,$cache.SlicerItems.Count);$k++){$it=$cache.SlicerItems.Item($k);$is+=@{name=$it.Name;caption=$it.Caption;selected=$it.Selected};Release $it};,$is}else{$null}};itemsCount=Attempt {$cache.SlicerItems.Count}};Release $sl};Release $slicers;Release $cache};Release $caches
 Release $window
}catch{$record.errors+= $_.Exception.Message}
finally{
 if($book){try{$book.Close($false)}catch{$record.errors+='닫기: '+$_.Exception.Message};Release $book}
 if($xl){if($owned){try{$xl.Quit()}catch{$record.errors+='종료: '+$_.Exception.Message}};Release $xl}
 [GC]::Collect();[GC]::WaitForPendingFinalizers()
 $after=Get-Item -LiteralPath $path;$record.unchanged=($after.Length -eq $before.Length -and $after.LastWriteTimeUtc.Ticks -eq $before.LastWriteTimeUtc.Ticks)
 $record.ms=[math]::Round(((Get-Date)-$started).TotalMilliseconds)
 $record|ConvertTo-Json -Depth 20|Set-Content -LiteralPath (Join-Path $OutputRoot ($Id+'.json')) -Encoding utf8
 [pscustomobject]@{id=$Id;opened=$record.opened;sheets=$record.sheets.Count;slicers=$record.slicers.Count;unchanged=$record.unchanged;ms=$record.ms;errors=$record.errors}|ConvertTo-Json -Depth 4 -Compress
}
if(!$record.opened -or !$record.unchanged -or $record.errors.Count){exit 1}
