param([Parameter(Mandatory)][string]$Directory)

$ErrorActionPreference = 'Stop'
$ownedPath = Join-Path $Directory 'owned.xlsx'
$peerPath = Join-Path $Directory 'peer.xlsx'
$references = [System.Collections.Generic.List[object]]::new()
$excel = $null
$owned = $null
$peer = $null
function Track-Com {
    param([object]$Value)
    if ($null -ne $Value -and [Runtime.InteropServices.Marshal]::IsComObject($Value)) {
        $references.Add($Value)
    }
    return ,$Value
}
try {
    $excel = Track-Com (New-Object -ComObject Excel.Application)
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $books = Track-Com $excel.Workbooks
    $owned = Track-Com ($books.Add())
    $owned.SaveAs($ownedPath)
    $owned.Close($false)
    $before = (Get-FileHash -LiteralPath $ownedPath).Hash
    $owned = Track-Com ($books.Open($ownedPath))
    $sheets = Track-Com $owned.Worksheets
    $sheet = Track-Com ($sheets.Item(1))
    $cell = Track-Com ($sheet.Range('B7'))
    $cell.Value2 = 'This must not be saved'
    $peer = Track-Com ($books.Add())
    $peer.SaveAs($peerPath)
    $result = & pwsh -NoProfile -File (Join-Path $PSScriptRoot 'Close-OwnedWorkbook.ps1') -Path $ownedPath | ConvertFrom-Json
    if ($LASTEXITCODE -ne 0) { throw 'Owned cleanup failed.' }
    if ($result.closed -ne 1 -or $books.Count -ne 1) { throw 'Owned cleanup did not preserve the peer workbook.' }
    $remaining = $peer
    if ($remaining.FullName -ne $peerPath) { throw 'Cleanup closed the wrong workbook.' }
    if ((Get-FileHash -LiteralPath $ownedPath).Hash -ne $before) { throw 'Cleanup saved unauthorized changes.' }
    $again = & pwsh -NoProfile -File (Join-Path $PSScriptRoot 'Close-OwnedWorkbook.ps1') -Path $ownedPath | ConvertFrom-Json
    if ($LASTEXITCODE -ne 0) { throw 'Repeated cleanup failed.' }
    if ($again.matched -ne 0 -or $books.Count -ne 1) { throw 'Repeated cleanup touched another workbook.' }
    for ($i = 0; $i -lt $references.Count; $i++) {
        for ($j = $i + 1; $j -lt $references.Count; $j++) {
            if ([object]::ReferenceEquals($references[$i], $references[$j])) {
                throw 'Cleanup tracked the same COM wrapper more than once.'
            }
        }
    }
    $peer.Close($false)
    $peer = $null
    $excel.Quit()
    $excel = $null
    @{ closed_only_owned = $true; discarded_changes = $true; repeated_cleanup_safe = $true } | ConvertTo-Json -Compress
} finally {
    if ($null -ne $peer) { $peer.Close($false) }
    if ($null -ne $excel) { $excel.Quit() }
    for ($i = $references.Count - 1; $i -ge 0; $i--) {
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($references[$i])
    }
}
