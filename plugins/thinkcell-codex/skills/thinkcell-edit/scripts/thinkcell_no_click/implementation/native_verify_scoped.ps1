param(
    [Parameter(Mandatory=$true)][string]$InputFile,
    [Parameter(Mandatory=$true)][string]$OutputFile,
    [Parameter(Mandatory=$true)][string]$ReportFile,
    [Parameter(Mandatory=$true)][string]$RenderFile,
    [int]$FirstSlideNumber=0,
    [int]$RenderSlideIndex=1,
    [string]$ScriptsRoot='',
    [string]$CheckpointFile=''
)
$ErrorActionPreference='Stop'
if($PSVersionTable.PSEdition-ne'Desktop'){throw 'Requires Windows PowerShell 5.1.'}
Import-Module (Join-Path $PSHOME 'Modules/Microsoft.PowerShell.Utility')

function Resolve-ScriptsRoot {
 param([string]$Requested)
 $candidates=@()
 if(-not[string]::IsNullOrWhiteSpace($Requested)){$candidates+=[IO.Path]::GetFullPath($Requested)}
 $here=[IO.Path]::GetFullPath($PSScriptRoot)
 $candidates+=Join-Path $here 'scripts'
 $candidates+=Join-Path $here '../scripts'
 $candidates+=Join-Path $here '../../'
 foreach($candidate in $candidates){
  $resolved=$null
  try{$resolved=(Resolve-Path -LiteralPath $candidate -ErrorAction Stop).Path}catch{continue}
  if((Test-Path -LiteralPath (Join-Path $resolved 'office_state.ps1') -PathType Leaf) -and
     (Test-Path -LiteralPath (Join-Path $resolved 'reliable_office_state.ps1') -PathType Leaf)){
   return $resolved
  }
 }
 throw 'Cannot locate office_state.ps1 and reliable_office_state.ps1. Pass -ScriptsRoot.'
}

$scripts=Resolve-ScriptsRoot $ScriptsRoot
. (Join-Path $scripts 'office_state.ps1')
. (Join-Path $scripts 'reliable_office_state.ps1')

$source=(Resolve-Path -LiteralPath $InputFile).Path
$dest=[IO.Path]::GetFullPath($OutputFile)
$RenderFile=[IO.Path]::GetFullPath($RenderFile)
$ReportFile=[IO.Path]::GetFullPath($ReportFile)
$checkpointPath=if([string]::IsNullOrWhiteSpace($CheckpointFile)){''}else{[IO.Path]::GetFullPath($CheckpointFile)}
$paths=@($source,$dest,$RenderFile,$ReportFile)
if($checkpointPath){$paths+=$checkpointPath}
if(@($paths|Select-Object -Unique).Count-ne$paths.Count){throw 'Input, output, render, report and checkpoint paths must differ.'}
foreach($target in @($dest,$RenderFile,$ReportFile)+$(if($checkpointPath){@($checkpointPath)}else{@()})){
 if(Test-Path -LiteralPath $target){throw 'Distinct new output files required.'}
 if(-not(Test-Path -LiteralPath ([IO.Path]::GetDirectoryName($target)) -PathType Container)){throw 'Output parent missing.'}
}
if($source-eq$dest){throw 'Cannot overwrite input.'}

$app=$null;$addins=$null;$addin=$null;$presentations=$null
$owned=$null;$slides=$null;$pageSetup=$null;$slide=$null
$before=$null
$cleanupErrors=@();$released=@()
$r=[ordered]@{source=$source;source_sha256=(Get-FileHash -LiteralPath $source).Hash;output=$dest;render=$RenderFile;other_presentations_before=$null;native_reopen_pass=$false;cleanup_release_complete=$false}

function State {
 return (Get-ReliableOfficeState -Application $app)
}
function Record-CleanupError([string]$Name,[object]$ErrorRecord){
 $script:cleanupErrors+=($Name+': '+[string]$ErrorRecord.Exception.Message)
}
function Release-OwnedReference([string]$Name,[object]$Value){
 if($null-eq$Value){return}
 try{
  [void][Runtime.InteropServices.Marshal]::ReleaseComObject($Value)
  $script:released+=$Name
 }catch{Record-CleanupError $Name $_}
}
function Write-ProgressCheckpoint([string]$Phase,[bool]$IsReleased){
 if(-not$checkpointPath){return}
 $payload=[ordered]@{phase=$Phase;pid=$PID;utc=[DateTime]::UtcNow.ToString('o');cleanup_release_complete=$IsReleased;released_references=@($script:released);cleanup_errors=@($cleanupErrors)}
 $payload|ConvertTo-Json -Depth 7|Set-Content -LiteralPath $checkpointPath -Encoding UTF8
}

try {
 $app=New-Object -ComObject PowerPoint.Application
 $addins=$app.COMAddIns
 $addin=$addins.Item('thinkcell.addin')
 if(-not$addin.Connect){throw 'think-cell disconnected.'}
 $presentations=$app.Presentations
 $before=State
 $r.other_presentations_before=$before
 Write-ProgressCheckpoint 'setup_ready' $false

 Copy-Item -LiteralPath $source -Destination $dest
 $owned=$presentations.Open($dest,$false,$false,$false)
 $slides=$owned.Slides
 $r.slides=[int]$slides.Count
 $pageSetup=$owned.PageSetup
 if($FirstSlideNumber-gt0){$pageSetup.FirstSlideNumber=$FirstSlideNumber}
 $owned.Saved=$false;$owned.Save()
 # Release child RCWs before closing the owned presentation. The presentation
 # itself is always closed first and then released before its variable is null.
 Release-OwnedReference 'pageSetup' $pageSetup;$pageSetup=$null
 Release-OwnedReference 'slides' $slides;$slides=$null
 $owned.Close();Release-OwnedReference 'owned' $owned;$owned=$null

 $owned=$presentations.Open($dest,$false,$false,$false)
 $slides=$owned.Slides
 if($slides.Count-ne$r.slides){throw 'Reopen slide count mismatch.'}
 if($RenderSlideIndex -lt 1 -or $RenderSlideIndex -gt $slides.Count){throw 'Render slide index is outside the opened presentation.'}
 $pageSetup=$owned.PageSetup
 $height=[int][Math]::Round(1920*$pageSetup.SlideHeight/$pageSetup.SlideWidth)
 $slide=$slides.Item($RenderSlideIndex)
 $r.render_slide_index=$RenderSlideIndex
 $slide.Export($RenderFile,'PNG',1920,$height)
 Release-OwnedReference 'slide' $slide;$slide=$null
 Release-OwnedReference 'pageSetup' $pageSetup;$pageSetup=$null
 Release-OwnedReference 'slides' $slides;$slides=$null
 $owned.Close();Release-OwnedReference 'owned' $owned;$owned=$null
 $r.native_reopen_pass=$true
}catch{$r.error=$_.Exception.Message}
finally {
 Write-ProgressCheckpoint 'cleanup_pending' $false

 # Close each owned presentation before releasing its own RCW. Child RCWs are
 # released first; no Quit, FinalReleaseComObject, visibility change or GC.
 if($null-ne$slide){Release-OwnedReference 'slide' $slide;$slide=$null}
 if($null-ne$pageSetup){Release-OwnedReference 'pageSetup' $pageSetup;$pageSetup=$null}
 if($null-ne$slides){Release-OwnedReference 'slides' $slides;$slides=$null}
 if($null-ne$owned){try{$owned.Saved=$true;$owned.Close()}catch{Record-CleanupError 'owned.Close' $_};Release-OwnedReference 'owned' $owned;$owned=$null}
 # Capture the same full state/sibling/source gates after owned objects are
 # closed, while the app/collection references remain valid. A state failure
 # remains visible in the report and never passes.
 try {
  if($null-ne$before){$r.other_presentations_after=Wait-OfficeState $before {State};$r.other_presentations_unchanged=Compare-OfficeState $before $r.other_presentations_after}
  else{$r.other_presentations_after=$null;$r.other_presentations_unchanged=$false}
 }catch{$r.other_presentations_after=$null;$r.other_presentations_unchanged=$false;$r.state_error=$_.Exception.Message}
 $r.state_scope='Visible, named and unsaved presentations; full snapshots retain saved hidden unnamed native transients.'
 $r.source_unchanged=(Get-FileHash -LiteralPath $source).Hash-eq$r.source_sha256
 if($null-ne$presentations){Release-OwnedReference 'presentations' $presentations;$presentations=$null}
 if($null-ne$addin){Release-OwnedReference 'addin' $addin;$addin=$null}
 if($null-ne$addins){Release-OwnedReference 'addins' $addins;$addins=$null}
 if($null-ne$app){Release-OwnedReference 'app' $app;$app=$null}
 $r.cleanup_release_complete=($cleanupErrors.Count-eq0)
 if($cleanupErrors.Count-gt0){$r.cleanup_release_errors=@($cleanupErrors)}
 Write-ProgressCheckpoint 'cleanup_returned' ([bool]$r.cleanup_release_complete)
 $r|ConvertTo-Json -Depth 10|Set-Content -LiteralPath $ReportFile -Encoding UTF8
 $r|ConvertTo-Json -Depth 10
}
if(-not$r.native_reopen_pass-or-not$r.other_presentations_unchanged-or-not$r.source_unchanged-or-not$r.cleanup_release_complete){exit 2}
