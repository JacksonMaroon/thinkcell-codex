# Set native defaults on a new copy. Existing chart formatting is not rewritten.
param(
 [Parameter(Mandatory=$true)][string]$InputFile,
 [Parameter(Mandatory=$true)][string]$OutputDirectory,
 [Parameter(Mandatory=$true)][string]$StyleFile
)
$ErrorActionPreference='Stop'
if($PSVersionTable.PSEdition-ne'Desktop'){throw 'Requires Windows PowerShell 5.1.'}
Import-Module (Join-Path $PSHOME 'Modules/Microsoft.PowerShell.Utility')
$source=(Resolve-Path -LiteralPath $InputFile).Path
$style=(Resolve-Path -LiteralPath $StyleFile).Path
$out=[IO.Path]::GetFullPath($OutputDirectory)
if(Test-Path -LiteralPath $out){throw 'Use a new output directory.'}
if([IO.Path]::GetExtension($source)-ne'.pptx'){throw 'Extract the donor to PPTX first.'}
$digest=(Get-FileHash -LiteralPath $source).Hash
$styleDigest=(Get-FileHash -LiteralPath $style).Hash
[xml]$styleXml=Get-Content -LiteralPath $style -Raw
$styleName=[string]$styleXml.DocumentElement.GetAttribute('name')
if(-not$styleName){throw 'Style XML has no name.'}
$app=New-Object -ComObject PowerPoint.Application
function State {
 $items=@();for($i=1;$i-le$app.Presentations.Count;$i++){$p=$app.Presentations.Item($i);$items+=,[ordered]@{path=[string]$p.FullName;saved=[int]$p.Saved;slides=[int]$p.Slides.Count;windows=[int]$p.Windows.Count}}
 return ($items|ConvertTo-Json -Depth 5 -Compress)
}
$before=State;$owned=$null
New-Item -ItemType Directory -Path $out|Out-Null
$copy=Join-Path $out 'styled.pptx'
$r=[ordered]@{status='FAILED';source_sha256=$digest;style_sha256=$styleDigest;expected_style=$styleName;output=$copy;scope='Defaults for new elements; existing element styles are preserved.'}
try {
 if(-not$app.COMAddIns.Item('thinkcell.addin').Connect){throw 'think-cell disconnected.'}
 $tc=$app.COMAddIns.Item('thinkcell.addin').Object
 Copy-Item -LiteralPath $source -Destination $copy
 $owned=$app.Presentations.Open($copy,$false,$false,$false)
 # Only retain designs actually used by the copied slides.
 $used=@();for($i=1;$i-le$owned.Slides.Count;$i++){$used+=[string]$owned.Slides.Item($i).Design.Name}
 for($i=$owned.Designs.Count;$i-ge1;$i--){if([string]$owned.Designs.Item($i).Name-notin$used){$owned.Designs.Item($i).Delete()}}
 for($i=1;$i-le$owned.Designs.Count;$i++){$tc.LoadStyle($owned.Designs.Item($i).SlideMaster,$style)}
 $owned.Save();$owned.Close();$owned=$null
 $owned=$app.Presentations.Open($copy,$false,$false,$false)
 $names=@();for($i=1;$i-le$owned.Designs.Count;$i++){$names+=[string]$tc.GetStyleName($owned.Designs.Item($i).SlideMaster)}
 if(@($names|Where-Object{$_-ne$styleName}).Count){throw 'Style name did not persist.'}
 $height=[int][Math]::Round(1920*$owned.PageSetup.SlideHeight/$owned.PageSetup.SlideWidth)
 $owned.Slides.Item(1).Export((Join-Path $out 'preview.png'),'PNG',1920,$height)
 $owned.Close();$owned=$null
 $r.style_names_after_reopen=$names
 $r.status='STYLE_DEFAULTS_VERIFIED'
}catch{$r.error=$_.Exception.Message}
finally {
 if($null-ne$owned){try{$owned.Saved=$true;$owned.Close()}catch{}}
 $r.other_presentations_unchanged=(State)-eq$before
 $r.source_unchanged=(Get-FileHash -LiteralPath $source).Hash-eq$digest
 $r.style_file_unchanged=(Get-FileHash -LiteralPath $style).Hash-eq$styleDigest
 $r|ConvertTo-Json -Depth 5|Set-Content -LiteralPath (Join-Path $out 'style.json') -Encoding utf8
 $r|ConvertTo-Json -Depth 5
}
if($r.status-eq'FAILED'-or-not$r.other_presentations_unchanged-or-not$r.source_unchanged-or-not$r.style_file_unchanged){exit 1}
