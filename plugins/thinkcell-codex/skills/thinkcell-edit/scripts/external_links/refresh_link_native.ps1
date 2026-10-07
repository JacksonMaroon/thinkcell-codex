param([Parameter(Mandatory=$true)][string]$Plan,[Parameter(Mandatory=$true)][string]$Python,[ValidateRange(5,120)][int]$TimeoutSeconds=45)
$ErrorActionPreference='Stop'
if($PSVersionTable.PSEdition-ne'Desktop'){throw 'Windows PowerShell 5.1 required'}
$p=Get-Content -LiteralPath $Plan -Raw|ConvertFrom-Json
if($p.schema-ne'thinkcell-refresh-plan-v1'){throw 'Unsupported plan'}
$r=$p.request
$scripts=Split-Path $PSScriptRoot -Parent
. (Join-Path $scripts 'office_state.ps1')
. (Join-Path $scripts 'reliable_office_state.ps1')
$validated=$false;$ppt=$null;$excel=$null;$pres=$null;$book=$null;$sheet=$null
$proof=$null;$beforePpt=$null;$beforeExcel=$null;$cleanup=@()
$result=[ordered]@{schema='thinkcell-refresh-result-v1';status='FAILED';automatic_refresh_verified=$false;native_reopen_pass=$false;request=$r}
function ExcelState {
 $items=@()
 for($i=1;$i-le$excel.Workbooks.Count;$i++){
  $b=$excel.Workbooks.Item($i)
  try{$items+=@{path=[string]$b.FullName;saved=[bool]$b.Saved;read_only=[bool]$b.ReadOnly;worksheets=[int]$b.Worksheets.Count}}finally{[void][Runtime.InteropServices.Marshal]::ReleaseComObject($b)}
 }
 return ConvertTo-Json -InputObject $items -Depth 5 -Compress
}
function Release([object]$Object){if($null-ne$Object){[void][Runtime.InteropServices.Marshal]::ReleaseComObject($Object)}}
function Invoke-PythonCheck([string]$Mode){
 # Windows PowerShell 5.1 surfaces native stderr as ErrorRecords. Capture the
 # whole diagnostic rather than allowing EAP Stop to abort at its first line.
 $prior=$ErrorActionPreference
 try{
  $ErrorActionPreference='Continue'
  if($Mode-eq'verify'){$output=& $Python (Join-Path $PSScriptRoot 'refresh_link.py') $Mode --plan $Plan --presentation $r.output_presentation 2>&1}
  else{$output=& $Python (Join-Path $PSScriptRoot 'refresh_link.py') $Mode --plan $Plan 2>&1}
  return @{exit_code=$LASTEXITCODE;text=($output|Out-String)}
 }finally{$ErrorActionPreference=$prior}
}
try{
 $validation=Invoke-PythonCheck 'validate-plan'
 if($validation.exit_code-ne0){throw ('Plan validation failed: '+$validation.text)}
 $validated=$true
 if((Get-FileHash -LiteralPath $r.input_presentation).Hash-ne$r.input_sha256-or(Get-FileHash -LiteralPath $r.source_workbook).Hash-ne$r.source_workbook_sha256){throw 'Sealed source hash mismatch'}
 if((Get-FileHash -LiteralPath $r.output_presentation).Hash-ne$p.prepared_presentation_sha256-or(Get-FileHash -LiteralPath $r.output_workbook).Hash-ne$p.prepared_workbook_sha256){throw 'Prepared candidate hash mismatch'}
 if((Test-Path -LiteralPath $r.report)-or(Test-Path -LiteralPath $r.preview)){throw 'Fresh report and preview required'}
 $ppt=New-Object -ComObject PowerPoint.Application
 try{$excel=[Runtime.InteropServices.Marshal]::GetActiveObject('Excel.Application')}
 catch [Runtime.InteropServices.COMException]{if($_.Exception.HResult-ne-2147221021){throw};$excel=New-Object -ComObject Excel.Application}
 $addin=$ppt.COMAddIns.Item('thinkcell.addin')
 try{if(-not$addin.Connect){throw 'think-cell disconnected'}}finally{Release $addin}
 $beforePpt=Get-ReliableOfficeState $ppt;$beforeExcel=ExcelState
 foreach($b in @($beforeExcel|ConvertFrom-Json)){if($b.path-eq$r.output_workbook){throw 'Candidate workbook already open'}}
 foreach($b in @($beforePpt|ConvertFrom-Json)){if($b.path-eq$r.output_presentation){throw 'Candidate presentation already open'}}
 # These are task-owned candidates only. Do not change application-wide options.
 $book=$excel.Workbooks.Open($r.output_workbook,0,$false)
 if($book.ReadOnly){throw 'Candidate workbook is read-only'}
 $sheet=$book.Worksheets.Item($r.sheet)
 $range=$sheet.Names.Item($p.rebind.guards.named_range_compatibility.target.name).RefersToRange
 try{if($range.Address($false,$false)-ne$r.expected_range){throw 'Native range guard mismatch'}}finally{Release $range}
 # Validate every old value and formula before the first write.
 foreach($u in $r.updates){
  $cell=$sheet.Range($u.cell)
  try{
   if($cell.HasFormula){throw 'Formula cells are unsupported'}
   $old=$cell.Value2
   if(($u.expected-is[string])-ne($old-is[string])-or$old-cne$u.expected){throw ('Cell expectation mismatch: '+$u.cell)}
  }finally{Release $cell}
 }
 $pres=$ppt.Presentations.Open($r.output_presentation,$false,$false,$false)
 foreach($u in $r.updates){$cell=$sheet.Range($u.cell);try{$cell.Value2=$u.value}finally{Release $cell}}
 $book.Save()
 foreach($u in $r.updates){$cell=$sheet.Range($u.cell);try{if($cell.Value2-cne$u.value){throw 'Saved workbook value mismatch'}}finally{Release $cell}}
 # No Send/UpdateBatch/internal datasheet write. Wait for the native advise sink.
 $timer=[Diagnostics.Stopwatch]::StartNew()
 do{
  Start-Sleep -Milliseconds 500
  $pres.Save()
  $proof=Invoke-PythonCheck 'verify'
  $verified=$proof.exit_code-eq0
 }while(-not$verified-and$timer.Elapsed.TotalSeconds-lt$TimeoutSeconds)
 if(-not$verified){throw ('Automatic refresh timed out: '+$proof.text)}
 $result.automatic_refresh_verified=$true
 Release $sheet;$sheet=$null
 $book.Close($false);Release $book;$book=$null
 $pres.Close();Release $pres;$pres=$null
 $pres=$ppt.Presentations.Open($r.output_presentation,$true,$false,$false)
 $slide=$pres.Slides.Item([int]$r.slide)
 $setup=$pres.PageSetup
 try{$height=[int][Math]::Round(1920*$setup.SlideHeight/$setup.SlideWidth)}finally{Release $setup}
 try{$slide.Export($r.preview,'PNG',1920,$height)}finally{Release $slide}
 $pres.Close();Release $pres;$pres=$null
 $proof=Invoke-PythonCheck 'verify'
 if($proof.exit_code-ne0){throw ('Saved reopened verification failed: '+$proof.text)}
 $result.proof=($proof.text|ConvertFrom-Json)
 $book=$excel.Workbooks.Open($r.output_workbook,0,$true)
 $sheet=$book.Worksheets.Item($r.sheet)
 foreach($u in $r.updates){$cell=$sheet.Range($u.cell);try{if($cell.Value2-cne$u.value){throw 'Reopened workbook mismatch'}}finally{Release $cell}}
 Release $sheet;$sheet=$null;$book.Close($false);Release $book;$book=$null
 $result.native_reopen_pass=$true
 $result.output_presentation_sha256=(Get-FileHash -LiteralPath $r.output_presentation).Hash
 $result.output_workbook_sha256=(Get-FileHash -LiteralPath $r.output_workbook).Hash
}catch{$result.error=$_.Exception.Message}
finally{
 try{Release $sheet}catch{$cleanup+=$_.Exception.Message}
 if($null-ne$book){try{$book.Close($false);Release $book}catch{$cleanup+=$_.Exception.Message}}
 if($null-ne$pres){try{$pres.Saved=$true;$pres.Close();Release $pres}catch{$cleanup+=$_.Exception.Message}}
 if($null-ne$beforePpt){try{$afterPpt=Wait-OfficeState $beforePpt {Get-ReliableOfficeState $ppt};$result.other_presentations_before=$beforePpt;$result.other_presentations_after=$afterPpt;$result.other_presentations_unchanged=Compare-OfficeState $beforePpt $afterPpt}catch{$cleanup+=$_.Exception.Message}}
 if($null-ne$beforeExcel){try{$afterExcel=ExcelState;$result.other_workbooks_before=$beforeExcel;$result.other_workbooks_after=$afterExcel;$result.other_workbooks_unchanged=$beforeExcel-eq$afterExcel}catch{$cleanup+=$_.Exception.Message}}
 try{$result.sources_unchanged=(Get-FileHash -LiteralPath $r.input_presentation).Hash-eq$r.input_sha256-and(Get-FileHash -LiteralPath $r.source_workbook).Hash-eq$r.source_workbook_sha256}catch{$cleanup+=$_.Exception.Message}
 try{Release $excel;Release $ppt}catch{$cleanup+=$_.Exception.Message}
 $result.cleanup_errors=$cleanup
 if($result.native_reopen_pass-and$result.other_presentations_unchanged-and$result.other_workbooks_unchanged-and$result.sources_unchanged-and$cleanup.Count-eq0){$result.status='VERIFIED_PERSISTENT_AUTOMATIC_REFRESH'}
 if($validated){
  $stream=[IO.File]::Open($r.report,[IO.FileMode]::CreateNew,[IO.FileAccess]::Write,[IO.FileShare]::None)
  try{$writer=New-Object IO.StreamWriter($stream,[Text.UTF8Encoding]::new($false));try{$writer.Write(($result|ConvertTo-Json -Depth 20))}finally{$writer.Dispose()}}finally{$stream.Dispose()}
 }
 $result|ConvertTo-Json -Depth 20
}
if($result.status-ne'VERIFIED_PERSISTENT_AUTOMATIC_REFRESH'){exit 2}
