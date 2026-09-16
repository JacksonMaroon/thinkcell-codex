param([string]$PlanFile,[string]$OutputFile,[string]$ReportFile)
$ErrorActionPreference='Stop'
$plan=Get-Content -Raw -Encoding UTF8 -LiteralPath $PlanFile|ConvertFrom-Json
$src=(Resolve-Path -LiteralPath $plan.input).Path
if((Get-FileHash -LiteralPath $src).Hash -ne $plan.input_sha256){throw 'Source hash mismatch'}
if(Test-Path -LiteralPath $OutputFile){throw 'Output exists'}
$app=New-Object -ComObject PowerPoint.Application
$pres=$app.Presentations.Open($src,0,0,0)
try{
 $slide=$pres.Slides.Item(1);$ops=@()
 foreach($target in $plan.targets){
  $found=@()
  foreach($s in $slide.Shapes){if($s.HasChart -eq -1){for($i=1;$i -le $s.Tags.Count;$i++){if($s.Tags.Value($i) -eq $target.shape_tag){$found+=,$s}}}}
  if($found.Count -ne 1){throw 'Exact native chart tag unresolved'}
  if($target.series_fills.PSObject.Properties.Count -gt 0 -or $target.point_fills.PSObject.Properties.Count -gt 0){throw 'Fill editing is not supported'}
 }
 if($plan.native_text_font_size){foreach($s in $slide.Shapes){if($s.Tags.Count -gt 0 -and $s.HasTextFrame -eq -1 -and $s.TextFrame.HasText -eq -1){$s.TextFrame.TextRange.Font.Size=$plan.native_text_font_size}}}
 $pres.SaveAs([IO.Path]::GetFullPath($OutputFile),24)
 @{status='API_APPLIED_REGENERATION_REQUIRED';operations=$ops;source_unchanged=((Get-FileHash -LiteralPath $src).Hash -eq $plan.input_sha256)}|ConvertTo-Json -Depth 8|Set-Content -Encoding UTF8 -LiteralPath $ReportFile
}finally{$pres.Close()}
