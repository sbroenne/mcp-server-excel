function ConvertTo-ExcelRunnerUtc {
    param($Value)
    if ($Value -is [DateTime] -and $Value.Kind -eq [DateTimeKind]::Utc) { return $Value }
    if ($Value -isnot [string]) { throw 'Runner evidence requires an explicit UTC timestamp.' }
    $parsed = [DateTime]::MinValue
    if (-not [DateTime]::TryParseExact($Value, [string[]]@('o', 'yyyy-MM-ddTHH:mm:ssK'), [Globalization.CultureInfo]::InvariantCulture,
        [Globalization.DateTimeStyles]::RoundtripKind, [ref]$parsed) -or $parsed.Kind -ne [DateTimeKind]::Utc) {
        throw 'Runner evidence requires an explicit UTC timestamp.'
    }
    return $parsed
}

function Assert-ExcelRunnerFreshness {
    param($CheckedAt, [DateTime]$Now = [DateTime]::UtcNow, [int]$MaximumAgeDays = 8)
    $checked = ConvertTo-ExcelRunnerUtc $CheckedAt
    if ($checked -gt $Now.AddMinutes(1) -or $checked -lt $Now.AddDays(-$MaximumAgeDays)) {
        throw 'Runner evidence is missing, from the future or overdue; keep the listener stopped.'
    }
}

function Assert-ExcelRunnerReadiness {
    param($State, [DateTime]$Now = [DateTime]::UtcNow, [DateTime]$BootTime = [DateTime]::MinValue)
    if ($State.state -ne 'ready' -or $State.license -ne 'Licensed' -or
        $State.formula -ne 42 -or $State.persistedFormula -ne 42) {
        throw 'The non-admin desktop must establish licensed Excel, calculation and saved readback.'
    }
    Assert-ExcelRunnerFreshness $State.checkedAt -Now $Now
    $boot = ConvertTo-ExcelRunnerUtc $State.bootTime
    $checked = ConvertTo-ExcelRunnerUtc $State.checkedAt
    if ($boot -gt $checked -or ($BootTime -ne [DateTime]::MinValue -and $boot -ne $BootTime.ToUniversalTime())) {
        throw 'Desktop evidence belongs to a different boot.'
    }
}

function Assert-ExcelRunnerPatchState {
    param($State, [DateTime]$Now = [DateTime]::UtcNow)
    if ($State.state -ne 'patched' -or $State.pendingRestart -isnot [bool] -or $State.pendingRestart -or
        $State.officeVersion -notmatch '^16\.0\.\d+\.\d+$' -or
        $State.officeTarget -notmatch '^16\.0\.\d+\.\d+$' -or
        [version]$State.officeVersion -lt [version]$State.officeTarget) {
        throw 'Windows must have a clean scan after restart, and Excel must meet the published Current release.'
    }
    Assert-ExcelRunnerFreshness $State.checkedAt -Now $Now
}

function Test-ExcelRunnerCloudRun {
    param($Run, [string]$Repository)
    return $Run.event -eq 'dynamic' -and $Run.path -eq 'dynamic/copilot-swe-agent/copilot' -and
        $Run.actor.id -eq 198982749 -and $Run.actor.type -eq 'Bot' -and
        $Run.head_repository.full_name -eq $Repository
}

function Assert-ExcelRunnerMaintenanceHistory {
    param([object[]]$Runs, [string]$Repository, [string]$DefaultBranch, [DateTime]$Now = [DateTime]::UtcNow)
    $latest = @($Runs | Where-Object {
        $_.head_branch -eq $DefaultBranch -and $_.head_repository.full_name -eq $Repository -and
        $_.path -eq '.github/workflows/excel-runner-maintenance.yml'
    } | Sort-Object {
        ConvertTo-ExcelRunnerUtc $_.created_at
    } -Descending | Select-Object -First 1)
    if ($latest.Count -ne 1 -or $latest[0].status -ne 'completed' -or $latest[0].conclusion -ne 'success') {
        throw 'Protected maintenance has not succeeded; do not repeatedly wake the VM for queued work.'
    }
    $checked = ConvertTo-ExcelRunnerUtc $latest[0].updated_at
    Assert-ExcelRunnerFreshness $checked.ToString('o') -Now $Now
}

function Test-ExcelRunnerJobTarget {
    param($Job, [string]$RunnerName = 'azure-excel-copilot', [string]$Label = 'excel-copilot')
    return $Job.runner_name -eq $RunnerName -or $Job.labels -contains $Label
}

function Assert-ExcelRunnerWorkspace {
    param([string]$Path)
    if ([IO.Path]::GetFullPath($Path).TrimEnd('\') -ine 'C:\actions-runner\_work\mcp-server-excel\mcp-server-excel') {
        throw 'Cleanup is restricted to this runner''s exact repository workspace.'
    }
    if (-not (Test-Path -LiteralPath $Path)) { return }
    $item = Get-Item -LiteralPath $Path -Force
    while ($item -and $item.FullName -ine 'C:\') {
        if ($item.Attributes -band [IO.FileAttributes]::ReparsePoint) { throw 'Do not follow a workspace directory link.' }
        $item = $item.Parent
    }
}

function Assert-ExcelRunnerIdle {
    param([object[]]$Jobs, $Guest, [string]$RunnerName = 'azure-excel-copilot', [switch]$AllowQueued)
    if (@($Jobs | Where-Object {
        (Test-ExcelRunnerJobTarget $_ -RunnerName $RunnerName) -and $_.status -ne 'completed' -and
        (-not $AllowQueued -or $_.status -ne 'queued')
    }).Count -or $Guest.listeners -ne 0 -or $Guest.workers -ne 0 -or $Guest.excel -ne 0 -or
        $Guest.taskRunning -ne $false) {
        throw 'Runner is not conclusively idle; do not update, restart or deallocate it.'
    }
}
