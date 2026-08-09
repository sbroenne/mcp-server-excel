<#
.SYNOPSIS
    Verifies aggregate star-history validation, snapshot updates, and SVG rendering.
#>

$ErrorActionPreference = "Stop"

$scriptPath = Join-Path $PSScriptRoot "Update-StarHistory.ps1"
$testRoot = Join-Path ([System.IO.Path]::GetTempPath()) "excelmcp-star-history-tests-$([Guid]::NewGuid().ToString('N'))"
$utf8NoBom = [System.Text.UTF8Encoding]::new($false)
$testsRun = 0

function Assert-True {
    param(
        [Parameter(Mandatory = $true)]
        [bool]$Condition,

        [Parameter(Mandatory = $true)]
        [string]$Message
    )

    if (-not $Condition) {
        throw $Message
    }
}

function Assert-Throws {
    param(
        [Parameter(Mandatory = $true)]
        [scriptblock]$Action,

        [Parameter(Mandatory = $true)]
        [string]$ExpectedMessage
    )

    try {
        & $Action
    }
    catch {
        Assert-True -Condition ($_.Exception.Message -like "*$ExpectedMessage*") `
            -Message "Expected error containing '$ExpectedMessage', got '$($_.Exception.Message)'."
        return
    }

    throw "Expected an error containing '$ExpectedMessage', but no error was thrown."
}

function New-HistoryFile {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Name,

        [Parameter(Mandatory = $true)]
        [string]$Content
    )

    $path = Join-Path $testRoot $Name
    [System.IO.File]::WriteAllText($path, $Content.TrimStart(), $utf8NoBom)
    return $path
}

function Invoke-StarHistory {
    param(
        [Parameter(Mandatory = $true)]
        [string]$HistoryPath,

        [Parameter(Mandatory = $true)]
        [string]$OutputPath,

        [int]$CurrentCount,

        [DateTimeOffset]$SnapshotDate,

        [switch]$WithoutSnapshot
    )

    $parameters = @{
        Repository = "owner/repository"
        HistoryPath = $HistoryPath
        OutputPath = $OutputPath
    }

    if (-not $WithoutSnapshot) {
        $parameters.CurrentCount = $CurrentCount
        $parameters.SnapshotDate = $SnapshotDate
    }

    & $scriptPath @parameters
}

New-Item -ItemType Directory -Path $testRoot | Out-Null

try {
    $historyPath = New-HistoryFile -Name "render.csv" -Content @"
date,count
2026-01-01,1
2026-01-03,3
"@
    $outputPath = Join-Path $testRoot "render.svg"
    Invoke-StarHistory -HistoryPath $historyPath -OutputPath $outputPath `
        -CurrentCount 5 -SnapshotDate ([DateTimeOffset]"2026-01-05T12:00:00Z")
    $svg = [System.IO.File]::ReadAllText($outputPath)
    [xml]$svgDocument = $svg
    Assert-True -Condition ($svg -like "*owner/repository - 5 stars*") `
        -Message "The SVG subtitle did not use the latest aggregate count."
    Assert-True -Condition ($svgDocument.DocumentElement.LocalName -eq "svg") `
        -Message "The generated file is not valid SVG XML."
    $rows = @(Import-Csv $historyPath)
    Assert-True -Condition ($rows.Count -eq 3) -Message "A new daily snapshot was not appended."
    Assert-True -Condition ($rows[-1].date -eq "2026-01-05" -and $rows[-1].count -eq "5") `
        -Message "The appended daily snapshot is incorrect."
    $testsRun++

    $historyPath = New-HistoryFile -Name "replace.csv" -Content @"
date,count
2026-01-01,1
2026-01-05,4
"@
    Invoke-StarHistory -HistoryPath $historyPath -OutputPath (Join-Path $testRoot "replace.svg") `
        -CurrentCount 5 -SnapshotDate ([DateTimeOffset]"2026-01-05T23:59:00Z")
    $rows = @(Import-Csv $historyPath)
    Assert-True -Condition ($rows.Count -eq 2) -Message "A same-day update created a duplicate snapshot."
    Assert-True -Condition ($rows[-1].count -eq "5") -Message "A same-day update did not replace the count."
    $testsRun++

    $historyPath = New-HistoryFile -Name "decrease.csv" -Content @"
date,count
2026-01-01,5
2026-01-02,4
"@
    Invoke-StarHistory -HistoryPath $historyPath -OutputPath (Join-Path $testRoot "decrease.svg") `
        -CurrentCount 3 -SnapshotDate ([DateTimeOffset]"2026-01-03T00:00:00Z")
    $svg = [System.IO.File]::ReadAllText((Join-Path $testRoot "decrease.svg"))
    Assert-True -Condition ($svg -like "*owner/repository - 3 stars*") `
        -Message "A valid count decrease was not rendered accurately."
    $testsRun++

    $historyPath = New-HistoryFile -Name "past.csv" -Content @"
date,count
2026-01-02,2
"@
    Assert-Throws -ExpectedMessage "precedes the latest history date" -Action {
        Invoke-StarHistory -HistoryPath $historyPath -OutputPath (Join-Path $testRoot "past.svg") `
            -CurrentCount 1 -SnapshotDate ([DateTimeOffset]"2026-01-01T00:00:00Z")
    }
    $testsRun++

    $historyPath = New-HistoryFile -Name "duplicate.csv" -Content @"
date,count
2026-01-01,1
2026-01-01,2
"@
    Assert-Throws -ExpectedMessage "strictly increasing" -Action {
        Invoke-StarHistory -HistoryPath $historyPath -OutputPath (Join-Path $testRoot "duplicate.svg") `
            -WithoutSnapshot
    }
    $testsRun++

    $historyPath = New-HistoryFile -Name "malformed.csv" -Content @"
date,count
2026-01-01,not-a-number
"@
    Assert-Throws -ExpectedMessage "invalid count" -Action {
        Invoke-StarHistory -HistoryPath $historyPath -OutputPath (Join-Path $testRoot "malformed.svg") `
            -WithoutSnapshot
    }
    $testsRun++

    $historyPath = New-HistoryFile -Name "empty.csv" -Content "date,count`n"
    Assert-Throws -ExpectedMessage "does not contain any aggregate rows" -Action {
        Invoke-StarHistory -HistoryPath $historyPath -OutputPath (Join-Path $testRoot "empty.svg") `
            -WithoutSnapshot
    }
    $testsRun++
}
finally {
    if (Test-Path $testRoot) {
        Remove-Item -Path $testRoot -Recurse -Force
    }
}

Write-Host "Passed $testsRun star-history tests." -ForegroundColor Green
