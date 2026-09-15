param([Parameter(Mandatory=$true)][string]$InputFile,[Parameter(Mandatory=$true)][string]$OutputFile,[Parameter(Mandatory=$true)][string]$ReportFile,[Parameter(Mandatory=$true)][string]$RenderFile,[int]$FirstSlideNumber=0)
$ErrorActionPreference='Stop'
if($PSVersionTable.PSEdition-ne'Desktop'){throw 'Requires Windows PowerShell 5.1.'}
Import-Module (Join-Path $PSHOME 'Modules/Microsoft.PowerShell.Utility')
$source=(Resolve-Path -LiteralPath $InputFile).Path
$dest=[IO.Path]::GetFullPath($OutputFile)
$RenderFile=[IO.Path]::GetFullPath($RenderFile)
$ReportFile=[IO.Path]::GetFullPath($ReportFile)
$paths=@($source,$dest,$RenderFile,$ReportFile)
if(@($paths|Select-Object -Unique).Count-ne4){throw 'Input, output, render and report paths must differ.'}
foreach($target in @($dest,$RenderFile,$ReportFile)){
 if(Test-Path -LiteralPath $target){throw 'Distinct new output files required.'}
 if(-not(Test-Path -LiteralPath ([IO.Path]::GetDirectoryName($target)) -PathType Container)){throw 'Output parent missing.'}
}
if($source-eq$dest){throw 'Cannot overwrite input.'}
$app=New-Object -ComObject PowerPoint.Application
function State {
 $items=@();for($i=1;$i-le$app.Presentations.Count;$i++){$p=$app.Presentations.Item($i);$items+=,[ordered]@{path=[string]$p.FullName;saved=[int]$p.Saved;slides=[int]$p.Slides.Count;windows=[int]$p.Windows.Count}}
 return ($items|ConvertTo-Json -Depth 5 -Compress)
}
$before=State;$owned=$null
$r=[ordered]@{source=$source;source_sha256=(Get-FileHash -LiteralPath $source).Hash;output=$dest;render=$RenderFile;other_presentations_before=$before;native_reopen_pass=$false}
try {
 if(-not$app.COMAddIns.Item('thinkcell.addin').Connect){throw 'think-cell disconnected.'}
 Copy-Item -LiteralPath $source -Destination $dest
 $owned=$app.Presentations.Open($dest,$false,$false,$false)
 $r.slides=[int]$owned.Slides.Count
 if($FirstSlideNumber-gt0){$owned.PageSetup.FirstSlideNumber=$FirstSlideNumber}
 $owned.Saved=$false;$owned.Save()
 $owned.Close();$owned=$null
 $owned=$app.Presentations.Open($dest,$false,$false,$false)
 if($owned.Slides.Count-ne$r.slides){throw 'Reopen slide count mismatch.'}
 $height=[int][Math]::Round(1920*$owned.PageSetup.SlideHeight/$owned.PageSetup.SlideWidth)
 $owned.Slides.Item(1).Export($RenderFile,'PNG',1920,$height)
 $owned.Close();$owned=$null
 $r.native_reopen_pass=$true
}catch{$r.error=$_.Exception.Message}
finally {
 if($null-ne$owned){try{$owned.Saved=$true;$owned.Close()}catch{}}
 $r.other_presentations_after=State
 $r.other_presentations_unchanged=$r.other_presentations_before-eq$r.other_presentations_after
 $r.source_unchanged=(Get-FileHash -LiteralPath $source).Hash-eq$r.source_sha256
 $r|ConvertTo-Json -Depth 8|Set-Content -LiteralPath $ReportFile -Encoding UTF8
 $r|ConvertTo-Json -Depth 8
}
if(-not$r.native_reopen_pass-or-not$r.other_presentations_unchanged-or-not$r.source_unchanged){exit 2}


