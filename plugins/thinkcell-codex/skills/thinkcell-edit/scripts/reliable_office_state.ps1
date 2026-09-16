# Candidate read-only Office state snapshot helper.
#
# Dot-source this file from a native verifier, then replace its State function
# with Get-ReliableOfficeState. It snapshots every presentation, including
# named, visible and unsaved work. A retry is allowed only when a known
# PowerPoint collection operation reports an index/range failure, which is the
# signature of a collection changing between Count and Item. Other COM and
# automation errors are rethrown immediately.

$ErrorActionPreference = 'Stop'

class TransientOfficeCollectionException : System.Exception {
    TransientOfficeCollectionException([string]$message, [System.Exception]$inner) : base($message, $inner) {}
}

function Test-TransientOfficeCollectionError {
    param(
        [Parameter(Mandatory=$true)][string]$Operation,
        [Parameter(Mandatory=$true)][System.Exception]$Exception
    )
    # PowerPoint's COM wrapper can surface this as RuntimeException rather than
    # COMException. Operation context and the collection/range wording are
    # therefore both required; generic COM failures never qualify.
    $message = [string]$Exception.Message
    if ($null -ne $Exception.InnerException) {
        $message += ' ' + [string]$Exception.InnerException.Message
    }
    $knownCollectionOperation = $Operation -match '^(Presentations|Slides|Windows)\.(Count|Item)$'
    $rangeFailure = $message -match '(?i)(out\s+of\s+range|integer\s+out\s+of\s+range|invalid\s+(index|argument))'
    $comLike = ($Exception -is [System.Runtime.InteropServices.COMException]) -or
               ($Exception -is [System.ArgumentOutOfRangeException]) -or
               ($Exception.GetType().FullName -like 'System.Management.Automation.*')
    # The PowerPoint wrapper's observed message is: "Integer out of range.
    # 1 is not in Index's valid range of 1 to 0." It does not mention the
    # collection name, so the already-known operation supplies that context.
    return ($knownCollectionOperation -and $rangeFailure -and $comLike)
}

function Invoke-OfficeCollectionRead {
    param(
        [Parameter(Mandatory=$true)][string]$Operation,
        [Parameter(Mandatory=$true)][scriptblock]$Action
    )
    try {
        return & $Action
    } catch {
        if (Test-TransientOfficeCollectionError -Operation $Operation -Exception $_.Exception) {
            $message = "Transient PowerPoint collection race during $($Operation): $($_.Exception.Message)"
            throw [TransientOfficeCollectionException]::new($message, $_.Exception)
        }
        throw
    }
}

function Get-OfficeStateSnapshotOnce {
    param([Parameter(Mandatory=$true)][object]$Application)
    $presentations = $Application.Presentations
    $count = [int](Invoke-OfficeCollectionRead -Operation 'Presentations.Count' -Action {
        [int]$presentations.Count
    })
    $items = @()
    for ($index = 1; $index -le $count; $index++) {
        $presentation = Invoke-OfficeCollectionRead -Operation 'Presentations.Item' -Action {
            $presentations.Item($index)
        }
        $items += [ordered]@{
            path = [string]$presentation.FullName
            saved = [int]$presentation.Saved
            slides = [int](Invoke-OfficeCollectionRead -Operation 'Slides.Count' -Action {
                [int]$presentation.Slides.Count
            })
            windows = [int](Invoke-OfficeCollectionRead -Operation 'Windows.Count' -Action {
                [int]$presentation.Windows.Count
            })
        }
    }
    # InputObject keeps the full array together. Piping the array would make
    # ConvertTo-Json serialize only one presentation per pipeline invocation.
    return (ConvertTo-Json -InputObject $items -Depth 5 -Compress)
}

function Get-ReliableOfficeState {
    param(
        [Parameter(Mandatory=$true)][object]$Application,
        [ValidateRange(1,10)][int]$MaxAttempts = 4,
        [ValidateRange(0,2000)][int]$RetryDelayMilliseconds = 125
    )
    for ($attempt = 1; $attempt -le $MaxAttempts; $attempt++) {
        try {
            return [string](Get-OfficeStateSnapshotOnce -Application $Application)
        } catch [TransientOfficeCollectionException] {
            if ($attempt -ge $MaxAttempts) {
                throw
            }
            if ($RetryDelayMilliseconds -gt 0) {
                Start-Sleep -Milliseconds $RetryDelayMilliseconds
            }
        }
    }
    throw 'Office state snapshot did not complete.'
}
