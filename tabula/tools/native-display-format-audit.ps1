param(
 [Parameter(Mandatory=$true)][string]$Manifest,
 [Parameter(Mandatory=$true)][string]$Targets,
 [Parameter(Mandatory=$true)][string]$Id,
 [Parameter(Mandatory=$true)][string]$OutputRoot
)
# One read-only original at a time. No refresh, recalculation, save, repair or external link update.
# This collects rendered Excel formatting; it never writes cell values, formulas or source paths.
$ErrorActionPreference='Stop'
function Release($value){if($null -ne $value -and [Runtime.InteropServices.Marshal]::IsComObject($value)){try{[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($value)}catch{}}}
function ReadStyle($range){
 $font=$null;$fill=$null
 try{
  $font=$range.Font;$fill=$range.Interior
  return [ordered]@{font=[ordered]@{name=$font.Name;size=$font.Size;bold=$font.Bold;italic=$font.Italic;color=$font.Color;underline=$font.Underline};fill=[ordered]@{color=$fill.Color;colorIndex=$fill.ColorIndex;pattern=$fill.Pattern;patternColor=$fill.PatternColor};numberFormat=$range.NumberFormat}
 }finally{Release $font;Release $fill}
}
$cases=@(Get-Content -LiteralPath $Manifest -Raw -Encoding UTF8|ConvertFrom-Json)
$case=@($cases|Where-Object {$_.id -eq $Id})
if($case.Count -ne 1){throw 'Expected one manifest entry.'}
$targetsForFile=@(Get-Content -LiteralPath $Targets -Raw -Encoding UTF8|ConvertFrom-Json|Where-Object {$_.id -eq $Id})
if(!$targetsForFile.Count){throw 'No target cells.'}
$sourcePath=[IO.Path]::GetFullPath($case[0].path)
$outputPath=[IO.Path]::GetFullPath($OutputRoot)
if(!$outputPath.StartsWith('D:\',[StringComparison]::OrdinalIgnoreCase)){throw 'Output must be on D:.'}
$null=New-Item -ItemType Directory -Path $outputPath -Force
$before=Get-Item -LiteralPath $sourcePath
$record=[ordered]@{id=$Id;started=[DateTime]::UtcNow.ToString('o');opened=$false;readOnly=$false;cells=@();errors=@();scope='Range direct style vs DisplayFormat; no value/formula collection'}
$xl=$null;$book=$null;$owned=$false;$watch=[Diagnostics.Stopwatch]::StartNew()
try{
 $xl=New-Object -ComObject Excel.Application
 if($xl.Workbooks.Count -ne 0){throw 'Excel instance is not isolated and empty.'}
 $owned=$true;$xl.Visible=$false;$xl.DisplayAlerts=$false;$xl.EnableEvents=$false;$xl.AutomationSecurity=3;$xl.AskToUpdateLinks=$false;$xl.ScreenUpdating=$false
 $record.excel=@{version=$xl.Version;build=$xl.Build}
 Write-Output ($Id+' DisplayFormat read-only open')
 $book=$xl.Workbooks.Open($sourcePath,0,$true)
 $record.opened=$true;$record.readOnly=$book.ReadOnly;$record.openMs=$watch.ElapsedMilliseconds
 if(!$book.ReadOnly){throw 'Workbook did not open read-only.'}
 foreach($target in $targetsForFile){
  $sheet=$null;$cell=$null;$display=$null
  $item=[ordered]@{sheetIndex=[int]$target.sheetIndex;address=[string]$target.address;direct=$null;display=$null}
  try{
   if($item.sheetIndex -lt 0 -or $item.sheetIndex -ge $book.Worksheets.Count){throw 'Invalid sheet index.'}
   if($item.address -notmatch '^[A-Z]{1,3}[1-9][0-9]*$'){throw 'Invalid cell address.'}
   $sheet=$book.Worksheets.Item($item.sheetIndex+1)
   if($sheet.Visible -eq -1){$sheet.Activate()}
   $cell=$sheet.Range($item.address)
   $item.direct=ReadStyle $cell
   $display=$cell.DisplayFormat
   $item.display=ReadStyle $display
  }catch{$item.error=$_.Exception.Message;$record.errors+=('sheet '+$item.sheetIndex+' '+$item.address+': '+$_.Exception.Message)}
  finally{Release $display;Release $cell;Release $sheet}
  $record.cells+=,$item
 }
}catch{$record.errors+=$_.Exception.Message}
finally{
 if($book){try{$book.Close($false)}catch{$record.errors+='Close: '+$_.Exception.Message};Release $book}
 if($xl){if($owned){try{$xl.Quit()}catch{$record.errors+='Quit: '+$_.Exception.Message}};Release $xl}
 [GC]::Collect();[GC]::WaitForPendingFinalizers()
 $after=Get-Item -LiteralPath $sourcePath
 $record.originalUnchanged=($after.Length -eq $before.Length -and $after.LastWriteTimeUtc.Ticks -eq $before.LastWriteTimeUtc.Ticks)
 $record.totalMs=$watch.ElapsedMilliseconds
 $record|ConvertTo-Json -Depth 15|Set-Content -LiteralPath (Join-Path $outputPath ($Id+'.json')) -Encoding utf8
 [pscustomobject]@{id=$Id;opened=$record.opened;cells=$record.cells.Count;errors=$record.errors.Count;originalUnchanged=$record.originalUnchanged;openMs=$record.openMs;totalMs=$record.totalMs}|ConvertTo-Json -Compress
}
if(!$record.opened -or !$record.originalUnchanged -or $record.errors.Count){exit 1}
