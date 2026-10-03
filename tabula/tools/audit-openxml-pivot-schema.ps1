param(
 [Parameter(Mandatory=$true)][string]$Path,
 [Parameter(Mandatory=$true)][string]$Output,
 [Parameter(Mandatory=$true)][string]$Sdk
)
# Read-only small-part validation. Never loads worksheet or cache-record bodies.
$ErrorActionPreference='Stop'
$inputPath=[IO.Path]::GetFullPath($Path);$outputPath=[IO.Path]::GetFullPath($Output)
if($inputPath.Equals($outputPath,[StringComparison]::OrdinalIgnoreCase)){throw 'Output must differ from input.'}
if(!$outputPath.StartsWith('D:\',[StringComparison]::OrdinalIgnoreCase)){throw 'Audit output must use D:.'}
Add-Type -AssemblyName WindowsBase
Add-Type -AssemblyName System.IO.Compression.FileSystem
Add-Type -Path $Sdk
$before=Get-Item -LiteralPath $inputPath
$validator=New-Object DocumentFormat.OpenXml.Validation.OpenXmlValidator([DocumentFormat.OpenXml.FileFormatVersions]::Office2021)
$validator.MaxNumberOfErrors=300
$zip=[IO.Compression.ZipFile]::OpenRead($inputPath)
$result=[ordered]@{sdk='2.20.0';version='Office2021';scope='bounded pivot/cache-definition root elements; no worksheet or cache-record bodies';parts=@();sourceUnchanged=$null}
try{
 foreach($entry in $zip.Entries){
  $kind=if($entry.FullName -match '^xl/pivotTables/pivotTable[0-9]+\.xml$'){'PivotTableDefinition'}elseif($entry.FullName -match '^xl/pivotCache/pivotCacheDefinition[0-9]+\.xml$'){'PivotCacheDefinition'}else{$null}
  if(!$kind){continue}
  if($entry.Length -gt 33554432){throw ('Small-part limit exceeded: '+$entry.FullName)}
  $stream=$entry.Open();$reader=New-Object IO.StreamReader($stream)
  try{$xml=$reader.ReadToEnd()}finally{$reader.Dispose();$stream.Dispose()}
  $xml=[regex]::Replace($xml,'^\s*<\?xml[^?]*\?>\s*','')
  $root=New-Object ('DocumentFormat.OpenXml.Spreadsheet.'+$kind) -ArgumentList (,$xml)
  $issues=@($validator.Validate($root)|ForEach-Object {@{id=$_.Id;type=[string]$_.ErrorType;description=$_.Description;xpath=$_.Path.XPath}})
  $result.parts+=@{part=$entry.FullName;bytes=$entry.Length;issues=$issues;atErrorLimit=($issues.Count -ge 300)}
 }
}finally{$zip.Dispose()}
$after=Get-Item -LiteralPath $inputPath
$result.sourceUnchanged=($before.Length -eq $after.Length -and $before.LastWriteTimeUtc.Ticks -eq $after.LastWriteTimeUtc.Ticks)
$result|ConvertTo-Json -Depth 12|Set-Content -LiteralPath $outputPath -Encoding utf8
$all=@($result.parts|ForEach-Object {$_.issues})
[pscustomobject]@{parts=$result.parts.Count;issues=$all.Count;ids=@($all|Group-Object {$_.id}|ForEach-Object {@{id=$_.Name;count=$_.Count}});sourceUnchanged=$result.sourceUnchanged}|ConvertTo-Json -Depth 5 -Compress
