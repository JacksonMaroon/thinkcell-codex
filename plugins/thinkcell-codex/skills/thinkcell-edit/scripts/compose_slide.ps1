param(
 [Parameter(Mandatory=$true)][string]$Target,
 [Parameter(Mandatory=$true)][string]$Donor,
 [Parameter(Mandatory=$true)][string]$TargetSha256,
 [Parameter(Mandatory=$true)][string]$DonorSha256,
 [Parameter(Mandatory=$true)][int]$FirstSlideNumber,
 [Parameter(Mandatory=$true)][string]$OutputDirectory
)
$ErrorActionPreference='Stop'
if($PSVersionTable.PSEdition-ne'Desktop'){throw 'Windows PowerShell 5.1 required'}
Import-Module (Join-Path $PSHOME 'Modules/Microsoft.PowerShell.Utility')
. (Join-Path $PSScriptRoot 'office_state.ps1')
$Target=(Resolve-Path -LiteralPath $Target).Path;$Donor=(Resolve-Path -LiteralPath $Donor).Path
$OutputDirectory=[IO.Path]::GetFullPath($OutputDirectory)
if(Test-Path -LiteralPath $OutputDirectory){throw 'New directory required'}
if((Get-FileHash -LiteralPath $Target).Hash-ne$TargetSha256-or(Get-FileHash -LiteralPath $Donor).Hash-ne$DonorSha256){throw 'Input hash changed'}
$app=New-Object -ComObject PowerPoint.Application
function State {
 $items=@();for($i=1;$i-le$app.Presentations.Count;$i++){$p=$app.Presentations.Item($i);$items+=,[ordered]@{path=[string]$p.FullName;saved=[int]$p.Saved;slides=[int]$p.Slides.Count;windows=[int]$p.Windows.Count}}
 return ($items|ConvertTo-Json -Depth 5 -Compress)
}
$before=State;$hashT=$TargetSha256;$hashD=$DonorSha256
for($i=1;$i-le$app.Presentations.Count;$i++){
 $open=$app.Presentations.Item($i)
 if(([string]$open.FullName-eq$Target-or[string]$open.FullName-eq$Donor)-and[int]$open.Saved-eq0){throw 'An input has unsaved changes'}
}
function Signature($shape){
 $tags=@();for($i=1;$i-le$shape.Tags.Count;$i++){$tags+=,([string]$shape.Tags.Name($i)+'='+[string]$shape.Tags.Value($i))}
 $text='';if([int]$shape.HasTextFrame-ne0){$text=[string]$shape.TextFrame.TextRange.Text}
 $children=@();if([int]$shape.Type-eq6){for($i=1;$i-le$shape.GroupItems.Count;$i++){$children+=,(Signature $shape.GroupItems.Item($i))}}
 return [ordered]@{name=[string]$shape.Name;type=[int]$shape.Type;text=$text;visible=[int]$shape.Visible;
  left=[Math]::Round([double]$shape.Left,3);top=[Math]::Round([double]$shape.Top,3);
  width=[Math]::Round([double]$shape.Width,3);height=[Math]::Round([double]$shape.Height,3);
  rotation=[Math]::Round([double]$shape.Rotation,3);tags=$tags;children=$children}
}
function NotesBody($slide,[bool]$create){
 for($i=1;$i-le$slide.NotesPage.Shapes.Placeholders.Count;$i++){
  $shape=$slide.NotesPage.Shapes.Placeholders.Item($i)
  if([int]$shape.PlaceholderFormat.Type-eq2){return $shape}
 }
 if($create){return $slide.NotesPage.Shapes.AddPlaceholder(2)}
 return $null
}
function IsNative($shape){
 if($shape.Tags.Item('THINKCELLSHAPEDONOTDELETE')-or$shape.Name-eq'thinkcellActiveDocDoNotDelete'){return $true}
 if([int]$shape.Type-eq6){
  $count=0;for($i=1;$i-le$shape.GroupItems.Count;$i++){if(IsNative $shape.GroupItems.Item($i)){$count++}}
  if($count-gt0-and$count-ne$shape.GroupItems.Count){throw 'Mixed ordinary/native group unsupported'}
  return $count-gt0
 }
 return $false
}
function NoteStyles($body){
 $runs=@();if($null-eq$body){return '[]'}
 $range=$body.TextFrame.TextRange
 if([string]::IsNullOrEmpty([string]$range.Text)){return '[]'}
 for($i=1;$i-le$range.Runs().Count;$i++){
  $run=$range.Runs($i,1)
  $runs+=,[ordered]@{text=[string]$run.Text;font=[string]$run.Font.Name;size=[double]$run.Font.Size;
    bold=[int]$run.Font.Bold;italic=[int]$run.Font.Italic;underline=[int]$run.Font.Underline;color=[int]$run.Font.Color.RGB}
 }
 return ($runs|ConvertTo-Json -Depth 5 -Compress)
}
New-Item -ItemType Directory -Path $OutputDirectory|Out-Null
$p=$null;$r=[ordered]@{status='FAILED'}
try{
 $out=Join-Path $OutputDirectory 'composed.pptx';Copy-Item -LiteralPath $Donor -Destination $out
 $p=$app.Presentations.Open($out,$false,$false,$false)
 if($p.Slides.Count-ne1){throw 'One donor slide required'}
 [void]$p.Slides.InsertFromFile($Target,1,1,1)
 if($p.Slides.Count-ne2){throw 'Target import failed'}
 $d=$p.Slides.Item(1);$t=$p.Slides.Item(2)
 $p.PageSetup.FirstSlideNumber=$FirstSlideNumber
 $renderHeight=[int][Math]::Round(1920*$p.PageSetup.SlideHeight/$p.PageSetup.SlideWidth)
 if(-not[bool]$t.FollowMasterBackground){throw 'Custom slide background requires a separate supported route'}
 # Only ordinary source shapes are copied. The complete donor chart and carrier stay in place.
 for($i=1;$i-le$t.Shapes.Count;$i++){
  $sh=$t.Shapes.Item($i)
  if(IsNative $sh){throw 'Target contains think-cell'}
 }
 for($i=$d.Shapes.Count;$i-ge1;$i--){
  $sh=$d.Shapes.Item($i)
  $keep=IsNative $sh
  if(-not$keep){$sh.Delete()}
 }
 $d.Design=$t.Design;$d.CustomLayout=$t.CustomLayout
 $d.FollowMasterBackground=$t.FollowMasterBackground
 $d.DisplayMasterShapes=$t.DisplayMasterShapes
 if($t.Shapes.Count-eq0){throw 'Use ordinary donor creation for a completely empty slide'}
 $t.Shapes.Range().Copy();$pasted=$d.Shapes.Paste()
 $r.source_shapes=$t.Shapes.Count;$r.pasted_shapes=$pasted.Count
 if($pasted.Count-ne$t.Shapes.Count){throw 'Shape copy count mismatch'}
 $preserved=@()
 for($i=1;$i-le$t.Shapes.Count;$i++){
  $original=$t.Shapes.Item($i);$copy=$pasted.Item($i);$copy.Name=$original.Name
  $oldSignature=Signature $original;$newSignature=Signature $copy
  if(($oldSignature|ConvertTo-Json -Depth 15 -Compress)-ne($newSignature|ConvertTo-Json -Depth 15 -Compress)){throw ('Copied shape changed: '+$original.Name)}
  $preserved+=,[ordered]@{source_id=[int]$original.Id;output_id=[int]$copy.Id;signature=$oldSignature}
 }
 $r.preserved_shapes=$preserved
 # Copy formatted speaker-note text through the native text API.
 $sourceNotes=NotesBody $t $false;$destinationNotes=NotesBody $d $true
 $expectedNoteStyles=NoteStyles $sourceNotes
 $expectedNotes='';$destinationNotes.TextFrame.TextRange.Text=''
 if($null-ne$sourceNotes){
  $expectedNotes=[string]$sourceNotes.TextFrame.TextRange.Text
  if($expectedNotes.Length-gt0){$sourceNotes.TextFrame.TextRange.Copy();[void]$destinationNotes.TextFrame.TextRange.Paste()}
 }
 $r.notes_unchanged=$expectedNotes-eq[string]$destinationNotes.TextFrame.TextRange.Text
 if(-not$r.notes_unchanged){throw 'Notes changed'}
 $t.Export((Join-Path $OutputDirectory 'before.png'),'PNG',1920,$renderHeight)
 $t.Delete();$p.Save();$p.Close();$p=$null
 $p=$app.Presentations.Open($out,$false,$false,$false)
 $saved=$p.Slides.Item(1)
 $savedNotes=NotesBody $saved $false
 if($null-eq$savedNotes-or[string]$savedNotes.TextFrame.TextRange.Text-ne$expectedNotes){throw 'Notes changed after native reopen'}
 if($expectedNotes.Length-gt0-and(NoteStyles $savedNotes)-ne$expectedNoteStyles){throw 'Note text formatting changed after native reopen'}
 $r.notes_text_style_unchanged=$true
 $decodedNoteStyles=(NoteStyles $savedNotes)|ConvertFrom-Json
 $r.notes_font_runs=@($decodedNoteStyles)
 foreach($entry in $preserved){
  $matches=@();for($i=1;$i-le$saved.Shapes.Count;$i++){if([int]$saved.Shapes.Item($i).Id-eq$entry.output_id){$matches+=,$saved.Shapes.Item($i)}}
  if($matches.Count-ne1){throw 'Copied shape identity lost after native reopen'}
  if((Signature $matches[0]|ConvertTo-Json -Depth 15 -Compress)-ne($entry.signature|ConvertTo-Json -Depth 15 -Compress)){throw 'Copied shape changed after native reopen'}
 }
 $saved.Export((Join-Path $OutputDirectory 'after.png'),'PNG',1920,$renderHeight)
 $p.Close();$p=$null
 $r.status='COMPOSED_REVIEW_REQUIRED';$r.output=$out
}catch{$r.error=$_.Exception.Message}
finally{
 if($null-ne$p){$p.Saved=$true;$p.Close()}
 $r.other_presentations_before=$before;$r.other_presentations_after=Wait-OfficeState $before {State}
 $r.other_presentations_unchanged=Compare-OfficeState $before $r.other_presentations_after
 $r.state_scope='Visible, named and unsaved presentations; full snapshots retain saved hidden unnamed native transients.'
 $r.sources_unchanged=((Get-FileHash -LiteralPath $Target).Hash-eq$hashT)-and((Get-FileHash -LiteralPath $Donor).Hash-eq$hashD)
 $r|ConvertTo-Json -Depth 7|Set-Content -LiteralPath (Join-Path $OutputDirectory 'composition.json') -Encoding utf8
 $r|ConvertTo-Json -Depth 7
}
if($r.status-eq'FAILED'-or-not$r.other_presentations_unchanged-or-not$r.sources_unchanged){exit 1}
