# PowerPoint/think-cell can create saved, hidden, unnamed transient presentations.
# Keep the full snapshots in reports. Compare all visible, named or unsaved work.
function Compare-OfficeState([string]$Before,[string]$After){
 function Retained([string]$Json){
  $decoded=$Json|ConvertFrom-Json
  $items=@($decoded|Where-Object {
   -not($_.windows-eq0-and$_.saved-eq-1-and$_.path-match'^Presentation\d+$')
  })
  return ($items|ConvertTo-Json -Depth 7 -Compress)
 }
 return (Retained $Before)-eq(Retained $After)
}

# Allow a short native settling interval only after an actual state mismatch.
# Return the raw final snapshot so callers can retain complete evidence.
function Wait-OfficeState([string]$Before,[scriptblock]$Snapshot){
 $after=[string](& $Snapshot)
 if(Compare-OfficeState $Before $after){return $after}
 for($attempt=0;$attempt-lt20;$attempt++){
  Start-Sleep -Milliseconds 250
  $after=[string](& $Snapshot)
  if(Compare-OfficeState $Before $after){break}
 }
 return $after
}
