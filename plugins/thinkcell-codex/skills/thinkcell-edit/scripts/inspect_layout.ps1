# Read native geometry from an owned read-only staging copy; never save a deck.
param(
 [Parameter(Mandatory=$true)][string]$InputFile,
 [Parameter(Mandatory=$true)][string]$ExpectedSha256,
 [Parameter(Mandatory=$true)][int]$SlideNumber,
 [Nullable[int]]$SlideId,
 [Parameter(Mandatory=$true)][string]$OutputDirectory
)
$ErrorActionPreference='Stop'
if($PSVersionTable.PSEdition-ne'Desktop'){throw 'Requires Windows PowerShell 5.1.'}
Import-Module (Join-Path $PSHOME 'Modules/Microsoft.PowerShell.Utility')
$source=(Resolve-Path -LiteralPath $InputFile).Path
$out=[IO.Path]::GetFullPath($OutputDirectory)
if(Test-Path -LiteralPath $out){throw 'Use a new output directory.'}
if($SlideNumber-lt1){throw 'SlideNumber must be positive.'}
$digest=(Get-FileHash -LiteralPath $source -Algorithm SHA256).Hash
if($ExpectedSha256-notmatch'^[a-fA-F0-9]{64}$'-or$digest-ne$ExpectedSha256){throw 'Source SHA256 does not match.'}
$app=New-Object -ComObject PowerPoint.Application
function State {
 $items=@()
 for($i=1;$i-le$app.Presentations.Count;$i++){
  $p=$app.Presentations.Item($i)
  $items+=,[ordered]@{path=[string]$p.FullName;saved=[int]$p.Saved;slides=[int]$p.Slides.Count;windows=[int]$p.Windows.Count}
 }
 return ($items|ConvertTo-Json -Depth 5 -Compress)
}
function ShapeText($shape) {
 $parts=@()
 try{if([int]$shape.HasTextFrame-ne0-and[int]$shape.TextFrame.HasText-ne0){$parts+=,[string]$shape.TextFrame.TextRange.Text}}catch{}
 if([int]$shape.Type-eq6){for($g=1;$g-le$shape.GroupItems.Count;$g++){$parts+=,(ShapeText $shape.GroupItems.Item($g))}}
 return ($parts-join"`n")
}
function Normalize-NoteRuns([string]$Text,$Runs) {
 if($Text.Length-gt0){return $Runs}
 return @()
}
function DescribeShape($shape,[double]$canvasWidth,[double]$canvasHeight) {
 $x=[double]$shape.Left;$y=[double]$shape.Top;$w=[double]$shape.Width;$h=[double]$shape.Height
 $rotation=[double]$shape.Rotation;$radians=$rotation*[Math]::PI/180
 $bw=[Math]::Abs($w*[Math]::Cos($radians))+[Math]::Abs($h*[Math]::Sin($radians))
 $bh=[Math]::Abs($w*[Math]::Sin($radians))+[Math]::Abs($h*[Math]::Cos($radians))
 $bx=$x+($w-$bw)/2;$by=$y+($h-$bh)/2
 $visible=[int]$shape.Visible;$text=ShapeText $shape;$shapeType=[int]$shape.Type
 $emptyInvisibleHelper=$false
 if($shapeType-in@(1,5,14,17)-and[string]::IsNullOrWhiteSpace($text)){
  try{$emptyInvisibleHelper=([int]$shape.Fill.Visible-eq0-and[int]$shape.Line.Visible-eq0)}catch{}
 }
 $cl=[Math]::Max([double]0,$bx);$ct=[Math]::Max([double]0,$by)
 $cr=[Math]::Min($canvasWidth,$bx+$bw);$cb=[Math]::Min($canvasHeight,$by+$bh)
 $outside=($cr-lt$cl-or$cb-lt$ct)
 $clipped=$null
 if(-not$outside){$clipped=[ordered]@{left=$cl;top=$ct;width=($cr-$cl);height=($cb-$ct)}}
 return [ordered]@{id=[int]$shape.Id;name=[string]$shape.Name;type=$shapeType;text=$text;visible=$visible;
  rotation=$rotation;native_frame=[ordered]@{left=$x;top=$y;width=$w;height=$h};
  bounds=[ordered]@{left=$bx;top=$by;width=$bw;height=$bh};clipped_bounds=$clipped;
  offcanvas=($bx-lt0-or$by-lt0-or$bx+$bw-gt$canvasWidth-or$by+$bh-gt$canvasHeight);
  obstacle=($visible-ne0-and-not$emptyInvisibleHelper-and-not$outside);empty_invisible_helper=$emptyInvisibleHelper}
}
$otherBefore=State
for($i=1;$i-le$app.Presentations.Count;$i++){
 $p=$app.Presentations.Item($i)
 if([string]$p.FullName-eq$source-and[int]$p.Saved-eq0){throw 'Source has unsaved edits; export the current connected document first.'}
}
New-Item -ItemType Directory -Path $out|Out-Null
$staged=Join-Path $out ('source'+[IO.Path]::GetExtension($source))
$preview=Join-Path $out 'slide.png';$owned=$null
$report=[ordered]@{status='FAILED';source_sha256=$digest.ToLowerInvariant();slide_number=$SlideNumber;source_unchanged=$false}
try {
 Copy-Item -LiteralPath $source -Destination $staged
 if((Get-FileHash -LiteralPath $staged).Hash-ne$digest){throw 'Staging copy hash does not match source.'}
 $owned=$app.Presentations.Open($staged,$true,$false,$false)
 if($SlideNumber-gt$owned.Slides.Count){throw 'SlideNumber out of range.'}
 $slide=$owned.Slides.Item($SlideNumber)
 if($null-ne$SlideId-and[int]$slide.SlideID-ne$SlideId){throw 'Slide number and stable slide ID do not match.'}
 $width=[double]$owned.PageSetup.SlideWidth;$height=[double]$owned.PageSetup.SlideHeight
 $report.slide_id=[int]$slide.SlideID;$report.width=$width;$report.height=$height
 $report.first_slide_number=[int]$owned.PageSetup.FirstSlideNumber
 $report.follow_master_background=[int]$slide.FollowMasterBackground
 $report.show_master_shapes=[int]$slide.DisplayMasterShapes
 $all=@();$obstacles=@();$offcanvas=@()
 for($i=1;$i-le$slide.Shapes.Count;$i++){
  $shape=DescribeShape $slide.Shapes.Item($i) $width $height
  $all+=,$shape
  if($shape.offcanvas){$offcanvas+=,$shape}
  if($shape.obstacle){
   $obstacles+=,[ordered]@{id=$shape.id;name=$shape.name;text=$shape.text;visible=$shape.visible;
    left=$shape.clipped_bounds.left;top=$shape.clipped_bounds.top;width=$shape.clipped_bounds.width;height=$shape.clipped_bounds.height}
  }
 }
 $report.shapes=$obstacles;$report.all_shapes=$all;$report.offcanvas_shapes=$offcanvas;$report.reserved=@()
 $inherited=@()
 foreach($layer in @('layout','master')){
  # Assign directly: sending COM Shapes through an if-expression pipeline unrolls it.
  if($layer-eq'layout'){$collection=$slide.CustomLayout.Shapes}else{$collection=$slide.Master.Shapes}
  for($i=1;$i-le$collection.Count;$i++){$entry=DescribeShape $collection.Item($i) $width $height;$entry.layer=$layer;$inherited+=,$entry}
 }
 $report.inherited_shapes=$inherited
 $report.inherited_occupancy_requires_review=$true
 $report.design_name=[string]$slide.Design.Name;$report.layout_name=[string]$slide.CustomLayout.Name
 $notes=@();$noteRuns=@()
 for($i=1;$i-le$slide.NotesPage.Shapes.Count;$i++){
  $ns=$slide.NotesPage.Shapes.Item($i)
  $isNotesBody=$false
  try{$isNotesBody=[int]$ns.PlaceholderFormat.Type-eq2}catch{continue}
  try{
   if($isNotesBody){
    $notes+=,(ShapeText $ns)
    $range=$ns.TextFrame.TextRange
    for($j=1;$j-le$range.Runs().Count;$j++){
     $run=$range.Runs($j,1)
     $noteRuns+=,[ordered]@{text=[string]$run.Text;font=[string]$run.Font.Name;size=[double]$run.Font.Size;
      bold=[int]$run.Font.Bold;italic=[int]$run.Font.Italic;underline=[int]$run.Font.Underline;color=[int]$run.Font.Color.RGB}
    }
   }
  }catch{throw ('Cannot read original speaker-note font runs: '+$_.Exception.Message)}
 }
 $report.notes_text=$notes-join"`n"
 $report.notes_font_runs=@(Normalize-NoteRuns $report.notes_text $noteRuns)
 $slide.Export($preview,'PNG',1920,[int][Math]::Round(1920*$height/$width))
 $owned.Saved=$true;$owned.Close();$owned=$null
 if((Get-FileHash -LiteralPath $source).Hash-ne$digest){throw 'Source changed during inspection.'}
 $report.source_unchanged=$true;$report.preview=$preview
 $report.preview_sha256=(Get-FileHash -LiteralPath $preview).Hash
 $report.status='LAYOUT_INSPECTED_REVIEW_REQUIRED'
 $report.instruction='Review rendered slide, inherited master/layout content, background decorations and offcanvas shapes before planning. Add occupied inherited rails to reserved. Bounding rectangles are conservative and do not establish all rendered occupancy.'
}catch{$report.error=$_.Exception.Message}
finally{
 if($null-ne$owned){try{$owned.Saved=$true;$owned.Close()}catch{}}
 $report.other_presentations_unchanged=(State)-eq$otherBefore
 if(-not$report.other_presentations_unchanged){$report.status='FAILED';$report.error='Other presentation state changed during inspection.'}
 $report|ConvertTo-Json -Depth 15|Set-Content -LiteralPath (Join-Path $out 'layout.json') -Encoding utf8
 $report|ConvertTo-Json -Depth 15
}
if($report.status-eq'FAILED'){exit 1}
