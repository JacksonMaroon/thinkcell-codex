# Extract a saved slide without applying a foreign template or touching the source.
param(
 [Parameter(Mandatory=$true)][string]$InputFile,
 [Parameter(Mandatory=$true)][int]$SlideNumber,
 [Parameter(Mandatory=$true)][string]$OutputDirectory
)
$ErrorActionPreference='Stop'
if($PSVersionTable.PSEdition-ne'Desktop'){throw 'Requires Windows PowerShell 5.1.'}
Import-Module (Join-Path $PSHOME 'Modules/Microsoft.PowerShell.Utility')
$source=(Resolve-Path -LiteralPath $InputFile).Path
$out=[IO.Path]::GetFullPath($OutputDirectory)
if(Test-Path -LiteralPath $out){throw 'Use a new output directory.'}
if($SlideNumber-lt1){throw 'SlideNumber must be positive.'}
$digest=(Get-FileHash -LiteralPath $source).Hash
$app=New-Object -ComObject PowerPoint.Application
function State {
 $items=@();for($i=1;$i-le$app.Presentations.Count;$i++){$p=$app.Presentations.Item($i);$items+=,[ordered]@{path=[string]$p.FullName;saved=[int]$p.Saved;slides=[int]$p.Slides.Count;windows=[int]$p.Windows.Count}}
 return ($items|ConvertTo-Json -Depth 5 -Compress)
}
$otherBefore=State
$owned=$null
for($i=1;$i-le$app.Presentations.Count;$i++){
 $p=$app.Presentations.Item($i)
 if([string]$p.FullName-eq$source-and[int]$p.Saved-eq0){throw 'Source has unsaved edits; use the connected plugin export to include them.'}
}
New-Item -ItemType Directory -Path $out|Out-Null
$copy=Join-Path $out 'slide.pptx'
$staged=Join-Path $out ('source'+[IO.Path]::GetExtension($source))
$before=Join-Path $out 'before.png'
$after=Join-Path $out 'after.png'
$report=[ordered]@{status='FAILED';source_sha256=$digest;slide_number=$SlideNumber;output=$copy}
try {
 Copy-Item -LiteralPath $source -Destination $staged
 $owned=$app.Presentations.Open($staged,$false,$false,$false)
 if($SlideNumber-gt$owned.Slides.Count){throw 'SlideNumber out of range.'}
 $display=[int]$owned.PageSetup.FirstSlideNumber+$SlideNumber-1
 $height=[int][Math]::Round(1920*$owned.PageSetup.SlideHeight/$owned.PageSetup.SlideWidth)
 $owned.Slides.Item($SlideNumber).Export($before,'PNG',1920,$height)
 for($i=$owned.Slides.Count;$i-ge1;$i--){if($i-ne$SlideNumber){$owned.Slides.Item($i).Delete()}}
 $usedDesign=[string]$owned.Slides.Item(1).Design.Name
 for($i=$owned.Designs.Count;$i-ge1;$i--){if([string]$owned.Designs.Item($i).Name-ne$usedDesign){$owned.Designs.Item($i).Delete()}}
 $owned.PageSetup.FirstSlideNumber=$display
 $owned.SaveAs($copy,24);$owned.Close();$owned=$null
 $owned=$app.Presentations.Open($copy,$false,$false,$false)
 if($owned.Slides.Count-ne1){throw 'Extraction did not preserve one slide.'}
 $owned.Slides.Item(1).Export($after,'PNG',1920,$height)
 $owned.Close();$owned=$null
 if((Get-FileHash -LiteralPath $source).Hash-ne$digest){throw 'Source changed during extraction.'}
 $report.status='EXTRACTED_REVIEW_REQUIRED'
 $report.source_unchanged=$true
 $report.output_sha256=(Get-FileHash -LiteralPath $copy).Hash
 $report.before_render=$before;$report.after_render=$after
 $report.instruction='Compare renders and inspect exact chart data before using the extracted slide. Existing internal slide links may not remain meaningful in a standalone slide.'
}catch{$report.error=$_.Exception.Message}
finally{
 if($null-ne$owned){try{$owned.Saved=$true;$owned.Close()}catch{}}
 $report.other_presentations_unchanged=(State)-eq$otherBefore
 if(-not$report.other_presentations_unchanged){$report.status='FAILED';$report.error='Other presentation state changed during extraction; inspect before continuing.'}
 $report|ConvertTo-Json -Depth 5|Set-Content -LiteralPath (Join-Path $out 'extraction.json') -Encoding utf8
 $report|ConvertTo-Json -Depth 5
}
if($report.status-eq'FAILED'){exit 1}
