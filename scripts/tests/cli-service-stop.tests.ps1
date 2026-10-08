$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
$stopScript = Join-Path $root 'scripts\Stop-ExcelCliService.ps1'
$release = Join-Path $root 'src\ExcelMcp.CLI\bin\Release\net10.0-windows\excelcli.exe'
$debug = Join-Path $root 'src\ExcelMcp.CLI\bin\Debug\net10.0-windows\excelcli.exe'

function New-Candidate([int]$Id, [string]$Path, [string]$Arguments) {
    [pscustomobject]@{
        ProcessId = $Id
        ExecutablePath = $Path
        CommandLine = "`"$Path`" $Arguments"
        CreationDate = [datetime]'2026-01-01T00:00:00Z'
    }
}

# Execute the real script with only the OS process boundary substituted.
function Get-CimInstance {
    [CmdletBinding()]
    param([string]$ClassName, [string]$Filter)
    if ($ClassName -ne 'Win32_Process') { throw 'Unexpected process provider.' }
    if ($global:queryFailure) { throw 'process-query-denied' }
    if ($Filter -eq "Name = 'excelcli.exe'") { return $global:candidates }
    if ($Filter -notmatch '^ProcessId = (\d+)$') { throw "Unexpected query: $Filter" }
    $global:revalidation[[int]$Matches[1]]
}

function Stop-Process {
    [CmdletBinding()]
    param([int]$Id, [switch]$Force, [string]$Name)
    if ($Name -or -not $Force -or $Id -le 0) { throw 'Unsafe process termination.' }
    $global:stopped.Add($Id)
    if ($global:exitDuringStop) {
        $global:revalidation.Remove($Id)
        throw 'process-exited'
    }
    if ($global:stopFailure) { throw 'process-stop-denied' }
}

function dotnet { throw 'A service stop must not compile or restore anything.' }

$local = @(
    (New-Candidate 11 $release 'service run --pipe-name one'),
    (New-Candidate 12 $debug 'service run --pipe-name "pipe two"')
)
$excluded = @(
    (New-Candidate 21 "$root-other\src\ExcelMcp.CLI\bin\Release\net10.0-windows\excelcli.exe" 'service run --pipe-name one'),
    (New-Candidate 22 $release 'range get-values --sheet-name "service run"'),
    (New-Candidate 23 $release 'service status'),
    (New-Candidate 24 (Join-Path $root 'src\ExcelMcp.McpServer\bin\Release\net10.0-windows\mcp-excel.exe') 'service run --pipe-name one'),
    (New-Candidate 25 'C:\Program Files\Microsoft Office\EXCEL.EXE' 'service run --pipe-name one')
)

foreach ($case in @(
    @{ Name = 'all-local'; Candidates = $local + $excluded; Expected = @(11, 12) },
    @{ Name = 'case-insensitive-pipe'; Candidates = $local + $excluded; Pipe = 'ONE'; Expected = @(11) },
    @{ Name = 'quoted-pipe'; Candidates = $local; Pipe = 'pipe two'; Expected = @(12) },
    @{ Name = 'exact-pipe-only'; Candidates = $local; Pipe = 'on'; Expected = @() },
    @{ Name = 'no-process'; Candidates = @(); Expected = @() },
    @{ Name = 'already-exited'; Candidates = @($local[0]); Missing = $true; Expected = @() },
    @{ Name = 'reused-pid'; Candidates = @($local[0]); Reused = $true; Expected = @() },
    @{ Name = 'changed-path'; Candidates = @($local[0]); ChangedPath = $true; Expected = @() },
    @{ Name = 'changed-command'; Candidates = @($local[0]); ChangedCommand = $true; Expected = @() },
    @{ Name = 'exit-during-stop'; Candidates = @($local[0]); ExitDuringStop = $true; Expected = @(11) },
    @{ Name = 'stop-denied'; Candidates = @($local[0]); StopFailure = $true; Expected = @(11); Error = 'process-stop-denied' },
    @{ Name = 'query-denied'; Candidates = @(); QueryFailure = $true; Expected = @(); Error = 'process-query-denied' }
)) {
    $global:candidates = $case.Candidates
    $global:revalidation = @{}
    foreach ($candidate in $global:candidates) {
        $global:revalidation[$candidate.ProcessId] = $candidate.PSObject.Copy()
    }
    if ($case.Missing) { $global:revalidation.Clear() }
    if ($case.Reused) { $global:revalidation[11].CreationDate = $local[0].CreationDate.AddSeconds(1) }
    if ($case.ChangedPath) { $global:revalidation[11].ExecutablePath = $debug }
    if ($case.ChangedCommand) { $global:revalidation[11].CommandLine = "`"$release`" service status" }
    $global:stopped = [Collections.Generic.List[int]]::new()
    $global:stopFailure = $case.StopFailure
    $global:queryFailure = $case.QueryFailure
    $global:exitDuringStop = $case.ExitDuringStop
    $failure = $null
    try { & $stopScript -PipeName $case.Pipe }
    catch { $failure = $_.Exception.Message }
    if ($case.Error) {
        if ($failure -ne $case.Error) { throw "$($case.Name): expected $($case.Error), got $failure." }
    }
    elseif ($failure) { throw "$($case.Name): $failure" }
    if (($global:stopped -join ',') -ne ($case.Expected -join ',')) {
        throw "$($case.Name): expected PIDs $($case.Expected -join ','), stopped $($global:stopped -join ',')."
    }
}
Write-Output 'CLI service stop identity, isolation, no-op and failure cases passed.'
