$ErrorActionPreference='Stop'
. (Join-Path $PSScriptRoot 'native-audit-structure.ps1')
$baseline=[pscustomobject]@{sheets=@([pscustomobject]@{name='A';pivots=@(@{name='P'});tables=@(@{name='T'});shapes=@(@{type=3},@{type=25})});slicers=@(@{sheet='A';name='S'})}
$equivalent=[pscustomobject]@{sheets=@([pscustomobject]@{name='A';pivots=@(@{name='P'});tables=1;shapes=@(@{type=3},@{type=25})});slicers=@(@{sheet='A';name='S'})}
$equal=Compare-NativeAuditStructure $baseline $equivalent
if(!$equal.matches -or $equal.differences.Count){throw 'Equivalent array/count metadata must match.'}
$lost=[pscustomobject]@{sheets=@([pscustomobject]@{name='A';pivots=@();tables=0;shapes=@()});slicers=@()}
$diff=Compare-NativeAuditStructure $baseline $lost
foreach($key in @('pivots','slicers','tables','charts','shapes')){if(!@($diff.differences|Where-Object {$_.field -eq $key}).Count){throw ('Missing count delta: '+$key)}}
if($diff.matches){throw 'Object loss cannot pass.'}
$renamed=[pscustomobject]@{sheets=@([pscustomobject]@{name='B';pivots=@(@{name='P'});tables=1;shapes=@(@{type=3},@{type=25})});slicers=@(@{sheet='B';name='S'})}
if((Compare-NativeAuditStructure $baseline $renamed).matches){throw 'Sheet identity mismatch cannot pass.'}
$empty=Compare-NativeAuditStructure ([pscustomobject]@{sheets=@();slicers=@()}) ([pscustomobject]@{sheets=@();slicers=@()})
if(!$empty.matches -or $empty.actual.totals.sheets -ne 0){throw 'Empty metadata must count zero.'}
Write-Output 'Native structural comparison: 4 cases passed (equivalence, five object counts, sheet identity, empty metadata).'
