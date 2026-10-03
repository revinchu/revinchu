param([Parameter(Mandatory=$true)][string]$Output)
$ErrorActionPreference='Stop'
function Release($x){if($null -ne $x -and [Runtime.InteropServices.Marshal]::IsComObject($x)){[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($x)}}
if(![IO.Path]::GetFullPath($Output).StartsWith('D:\',[StringComparison]::OrdinalIgnoreCase)){throw 'Output must use D:.'}
$xl=$null;$book=$null;$sheet=$null;$owned=$false;$results=@()
try{
 $xl=New-Object -ComObject Excel.Application
 if($xl.Workbooks.Count -ne 0){throw 'Not an isolated empty Excel.'};$owned=$true
 $xl.Visible=$false;$xl.DisplayAlerts=$false;$xl.EnableEvents=$false;$xl.AutomationSecurity=3;$xl.AskToUpdateLinks=$false
 $book=$xl.Workbooks.Add();$sheet=$book.Worksheets.Item(1)
 $bases=@('FFFFFF','000000','4472C4','156082','ED7D31','A5A5A5','FFC000','70AD47')
 $tints=@(0.2,0.4,0.6,0.79998168889431442,-0.249977111117893,-0.5,0.499984740745262,0.5)
 $r=1
 foreach($base in $bases){foreach($tint in $tints){
  $cell=$sheet.Cells.Item($r++,1);$fill=$cell.Interior
  try{
   $rgb=[Convert]::ToInt32($base.Substring(0,2),16)+256*[Convert]::ToInt32($base.Substring(2,2),16)+65536*[Convert]::ToInt32($base.Substring(4,2),16)
   $fill.Color=$rgb;$fill.TintAndShade=$tint;$value=[int]$fill.Color
   $hex=('{0:X2}{1:X2}{2:X2}' -f ($value -band 255),(($value -shr 8) -band 255),(($value -shr 16) -band 255))
   $results+=@{base=$base;requestedTint=$tint;storedTint=$fill.TintAndShade;rgb=$hex}
  }finally{Release $fill;Release $cell}
 }}
 $record=@{excel=@{version=$xl.Version;build=$xl.Build};scope='Synthetic unsaved workbook; direct RGB Interior.TintAndShade oracle; no user file opened';results=$results}
 $record|ConvertTo-Json -Depth 8|Set-Content -LiteralPath $Output -Encoding utf8
 Write-Output ('Native tint samples: '+$results.Count)
}finally{Release $sheet;if($book){$book.Close($false);Release $book};if($xl){if($owned){$xl.Quit()};Release $xl};[GC]::Collect();[GC]::WaitForPendingFinalizers()}
