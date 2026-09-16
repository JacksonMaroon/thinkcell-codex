param([Parameter(Mandatory=$true)][string]$InputFile,[Parameter(Mandatory=$true)][string]$OutputFile,[Parameter(Mandatory=$true)][string]$ReportFile,[Parameter(Mandatory=$true)][string]$ExpectedSha256,[Parameter(Mandatory=$true)][string]$AbsoluteText,[Parameter(Mandatory=$true)][string]$RelativeText)
$ErrorActionPreference='Stop'
$source=(Resolve-Path -LiteralPath $InputFile).Path
$output=[IO.Path]::GetFullPath($OutputFile);$report=[IO.Path]::GetFullPath($ReportFile)
if(@($source,$output,$report|Select-Object -Unique).Count-ne3){throw 'Input, output and report paths must differ.'}
if((Get-FileHash -LiteralPath $source).Hash-ne$ExpectedSha256){throw 'Source SHA-256 mismatch.'}
if((Test-Path -LiteralPath $output) -or (Test-Path -LiteralPath $report)){throw 'Output/report must be new.'}
Copy-Item -LiteralPath $source -Destination $output
$app=New-Object -ComObject PowerPoint.Application;$owned=$null
function CharCodes($range){$a=@();for($i=1;$i-le$range.Length;$i++){$t=[string]$range.Characters($i,1).Text;$a+=,[int][char]$t[0]};return $a}
function State {$items=@();for($i=1;$i-le$app.Presentations.Count;$i++){$p=$app.Presentations.Item($i);$items+=,[ordered]@{path=[string]$p.FullName;saved=[int]$p.Saved;slides=[int]$p.Slides.Count;windows=[int]$p.Windows.Count}};return ($items|ConvertTo-Json -Depth 5 -Compress)}
$beforeState=State;$result=[ordered]@{status='NATIVE_PARTIAL_DELETE_FAILED';source_sha256=(Get-FileHash -LiteralPath $source).Hash;shape_tag='tbUpiE_yCia4NQkLBVWFCIA';other_presentations_before=$beforeState}
$failure=$null
try{
 $owned=$app.Presentations.Open($output,$false,$false,$false)
 $matches=@();foreach($shape in $owned.Slides.Item(1).Shapes){for($i=1;$i-le$shape.Tags.Count;$i++){if($shape.Tags.Value($i)-eq'tbUpiE_yCia4NQkLBVWFCIA'){$matches+=,$shape}}}
 if($matches.Count-ne1){throw 'Exact native label tag unresolved.'};$shape=$matches[0]
 if(([int]$shape.HasTextFrame -ne -1) -or ([int]$shape.TextFrame.HasText -ne -1)){throw 'Selected native label has no text frame.'}
 $text=$shape.TextFrame.TextRange;$before=[string]$text.Text;$codes=CharCodes $text
 $expectedBefore=$AbsoluteText+[char]11+'('+$RelativeText+')';$expectedAfter='('+$RelativeText+')';$prefixLength=$AbsoluteText.Length+1
 if($before-ne$expectedBefore -or $text.Length-ne$expectedBefore.Length){throw ('Unexpected dual-label character sequence: '+($codes-join ','))}
 # Remove only the absolute field and its paragraph separator.  Keep the
 # original opening/closing literal parentheses and the relative field in its
 # original trailing position, so the native field is not the first run.
 $text.Characters(1,$prefixLength).Delete()
 $after=[string]$shape.TextFrame.TextRange.Text
 if($after-ne$expectedAfter -or $shape.TextFrame.TextRange.Length-ne$expectedAfter.Length){throw 'Partial deletion did not retain the exact relative field and parentheses.'}
 $owned.Save();$owned.Close();$owned=$null
 $result.status='NATIVE_PARTIAL_DELETE_KEEP_PARENS_COMPLETE_STATIC_READBACK_REQUIRED';$result.output_sha256=(Get-FileHash -LiteralPath $output).Hash;$result.before_text=$before;$result.before_character_codes=$codes;$result.deleted_character_range=@(1,$prefixLength);$result.after_text=$after;$result.relative_character_range_preserved=@(($prefixLength + 1),$expectedBefore.Length)
}catch{$failure=$_.Exception.Message;$result.error=$failure
}finally{if($null-ne$owned){try{$owned.Saved=$true;$owned.Close()}catch{}};$result.other_presentations_after=State;$result.other_presentations_unchanged=($result.other_presentations_before-eq$result.other_presentations_after);$result.source_unchanged=((Get-FileHash -LiteralPath $source).Hash-eq$ExpectedSha256);$result|ConvertTo-Json -Depth 8|Set-Content -LiteralPath $report -Encoding UTF8}
if($failure){throw $failure}
