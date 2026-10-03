# Pure metadata comparison. No COM, file writes, workbook changes or source values.
function Get-NativeAuditStructure($book) {
 $sheets=@($book.sheets | Where-Object {$null -ne $_})
 $rows=@();$totals=[ordered]@{sheets=$sheets.Count;pivots=0;slicers=0;tables=0;charts=0;shapes=0}
 foreach($sheet in $sheets){
  $row=[ordered]@{name=[string]$sheet.name;pivots=@($sheet.pivots|Where-Object {$null -ne $_}).Count;slicers=@($book.slicers|Where-Object {$null -ne $_ -and $_.sheet -ceq $sheet.name}).Count;tables=$(if($sheet.tables -is [ValueType]){[int]$sheet.tables}else{@($sheet.tables|Where-Object {$null -ne $_}).Count});charts=@($sheet.shapes|Where-Object {$null -ne $_ -and $_.type -eq 3}).Count;shapes=@($sheet.shapes|Where-Object {$null -ne $_}).Count}
  foreach($key in @('pivots','slicers','tables','charts','shapes')){$totals[$key]+=$row[$key]}
  $rows+=,$row
 }
 return @{totals=$totals;sheets=$rows}
}
function Compare-NativeAuditStructure($expectedBook,$actualBook) {
 $expected=Get-NativeAuditStructure $expectedBook;$actual=Get-NativeAuditStructure $actualBook;$differences=@()
 foreach($key in @('sheets','pivots','slicers','tables','charts','shapes')){
  if($expected.totals[$key] -ne $actual.totals[$key]){$differences+=@{field=$key;expected=$expected.totals[$key];actual=$actual.totals[$key]}}
 }
 for($i=0;$i -lt [Math]::Max($expected.sheets.Count,$actual.sheets.Count);$i++){
  $e=if($i -lt $expected.sheets.Count){$expected.sheets[$i]}else{$null};$a=if($i -lt $actual.sheets.Count){$actual.sheets[$i]}else{$null}
  foreach($key in @('name','pivots','slicers','tables','charts','shapes')){
   $ev=if($null -ne $e){$e[$key]}else{$null};$av=if($null -ne $a){$a[$key]}else{$null}
   if($ev -cne $av){$differences+=@{field=('sheets['+$i+'].'+$key);expected=$ev;actual=$av}}
  }
 }
 return @{expected=$expected;actual=$actual;differences=$differences;matches=($differences.Count -eq 0)}
}
