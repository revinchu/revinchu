import test from 'node:test';
import assert from 'node:assert/strict';
import { readFile, mkdir, mkdtemp, rm } from 'node:fs/promises';
import { dirname, resolve, join, sep } from 'node:path';
import { fileURLToPath } from 'node:url';
import { spawnSync } from 'node:child_process';

const tool=resolve(dirname(fileURLToPath(import.meta.url)),'../tools/native-layout-oracle.ps1');
test('native helper는 기존 PID를 설정하기 전 거절하고 소유·상태 검증 후에만 Quit한다',async()=>{
  const source=await readFile(tool,'utf8');
  assert.ok(source.includes('$instancePid-in$beforeIds'));
  assert.ok(source.includes('$count-ne0'));
  assert.ok(source.indexOf('$count-ne0')<source.indexOf('$xl.Visible=$false'));
  assert.ok(source.includes('$ownedPid-notin$beforeIds'));
  assert.ok(source.includes('$currentCount-eq$expectedCount'));
  assert.ok(source.indexOf('if($canClose)')<source.indexOf('$xl.Quit()'));
  assert.ok(source.includes('$books.Open($source,0,$true'));
  assert.ok(!/\.(Save|SaveAs|Refresh|Calculate|Activate|Select)\s*\(/.test(source));
  assert.ok(!source.includes('GetActiveObject'));
});

test('native 순수 함수는 null을0으로 바꾸지 않고 XLM/preflight만 tiny ZIP에서 검사한다',{skip:process.platform!=='win32'},async()=>{
  const root=resolve(dirname(fileURLToPath(import.meta.url)),'../.local');await mkdir(root,{recursive:true});
  const folder=await mkdtemp(join(root,'native-helper-test-'));
  const command=String.raw`
$ErrorActionPreference='Stop'
$script:themeUnavailable=0;$StartupRetries=1
$tokens=$null;$errors=$null
$ast=[Management.Automation.Language.Parser]::ParseFile($env:WIXEL_NATIVE_TEST_TOOL,[ref]$tokens,[ref]$errors)
if($errors.Count){throw 'PowerShell parse failed'}
$names=@('Invoke-Read','Required-Number','Optional-Read','Native-OpenPreflight','Property')
$functions=$ast.FindAll({param($node)$node -is [Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -in $names},$true)
foreach($function in $functions){. ([scriptblock]::Create($function.Extent.Text))}
$empty=Invoke-Read {,@()};$collection=Invoke-Read {,@(1,2)}
if($null-eq$empty -or $empty.Count-ne0 -or $collection-isnot[array] -or $collection.Count-ne2){throw 'collection was enumerated or null'}
if((Required-Number {23} 'mock')-ne23 -or (Required-Number {0} 'mock')-ne0){throw 'number changed'}
foreach($invalid in @($null,$true,[double]::NaN,[double]::PositiveInfinity,'not-number')){
  $rejected=$false;try{[void](Required-Number {$invalid} 'mock')}catch{$rejected=$true};if(!$rejected){throw 'invalid became number'}
}
$readErrors=[System.Collections.Generic.List[string]]::new()
$optionalCollection=Optional-Read {,@(1,2)} $readErrors 'mock.collection'
if($optionalCollection-isnot[array] -or $optionalCollection.Count-ne2){throw 'optional collection was enumerated'}
$value=Optional-Read {throw 'sample unavailable'} $readErrors 'font.color'
if($null-ne$value -or $readErrors.Count-ne1){throw 'unavailable was hidden'}
if(@((Property ([pscustomobject]@{}) 'cells')|Where-Object {$null-ne$_}).Count-ne0){throw 'missing samples were not empty'}
Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem
function Synthetic-Zip($name,$entry,$text){
  $path=Join-Path $env:WIXEL_NATIVE_TEST_FOLDER $name
  $zip=[IO.Compression.ZipFile]::Open($path,[IO.Compression.ZipArchiveMode]::Create)
  try{$part=$zip.CreateEntry($entry);$stream=$part.Open();$writer=[IO.StreamWriter]::new($stream);try{$writer.Write($text)}finally{$writer.Dispose();$stream.Dispose()}}finally{$zip.Dispose()}
  return $path
}
$clean=Synthetic-Zip 'clean.xlsx' 'xl/connections.xml' '<connections><connection refreshOnLoad="0"/></connections>'
$flagged=Synthetic-Zip 'flagged.xlsx' 'xl/pivotCache/pivotCacheDefinition1.xml' ('<pivotCacheDefinition '+(' '.PadLeft(8190))+'refreshOnLoad="true"/>')
$xlm=Synthetic-Zip 'xlm.xlsx' 'xl/macrosheets/sheet1.xml' '<worksheet/>'
if((Native-OpenPreflight $clean).status-ne'verified'){throw 'clean preflight failed'}
$report=Native-OpenPreflight $flagged
if($report.status-ne'verified' -or !$report.refreshOnLoad){throw 'Open-contract refresh flag was lost or falsely blocked'}
if((Native-OpenPreflight $xlm).status-ne'xlm-macros-not-covered-by-force-disable'){throw 'XLM not refused before COM'}
Write-Output 'pure-native-helper-tests-ok'
`;
  try {
    const result=spawnSync('powershell.exe',['-NoProfile','-NonInteractive','-EncodedCommand',Buffer.from(command,'utf16le').toString('base64')],{encoding:'utf8',timeout:30000,env:{...process.env,TEMP:folder,TMP:folder,WIXEL_NATIVE_TEST_TOOL:tool,WIXEL_NATIVE_TEST_FOLDER:folder}});
    assert.equal(result.status,0,result.stdout+'\n'+result.stderr);assert.match(result.stdout,/pure-native-helper-tests-ok/);
  } finally {assert.ok(resolve(folder).startsWith(root+sep));await rm(folder,{recursive:true,force:true});}
});