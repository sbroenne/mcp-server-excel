$ErrorActionPreference = 'Stop'

function Stop-ExcelRunnerJobProcess {
    param($Entry, [string]$ComputerName)
    $process = Get-Process -Id $Entry.ProcessId -ErrorAction SilentlyContinue
    if (-not $process) { return }
    try {
        try { $owner = Invoke-CimMethod $Entry -MethodName GetOwner }
        catch {
            if ($process.HasExited) { return }
            throw
        }
        if ($owner.ReturnValue -ne 0 -or [string]::IsNullOrWhiteSpace($owner.User) -or
            [string]::IsNullOrWhiteSpace($owner.Domain)) {
            if ($process.HasExited) { return }
            throw 'A remaining process owner could not be established.'
        }
        if ($owner.User -ine 'excelrunner' -or $owner.Domain -ine $ComputerName) { return }
        try { $null = $process.Handle }
        catch {
            if ($process.HasExited) { return }
            throw
        }
        if ($process.HasExited) { return }
        if ($process.StartTime -isnot [DateTime]) { throw 'A job-owned process start time could not be established.' }
        $ticks = $process.StartTime.ToUniversalTime().Ticks
        if (($ticks - $ticks % 10) -ne $Entry.CreationDate.ToUniversalTime().Ticks) { return }
        if (-not $process.HasExited) {
            Stop-Process -Id $process.Id -Force
            if (-not $process.WaitForExit(10000)) { throw 'A job-owned process did not stop.' }
        }
    }
    finally { $process.Dispose() }
}

function Remove-ExcelRunnerJobWorkspace {
    param([string]$Path)
    Assert-ExcelRunnerWorkspace $Path
    if (-not (Test-Path -LiteralPath $Path)) { return }
    $directories = [Collections.Generic.Stack[string]]::new()
    $directories.Push($Path)
    while ($directories.Count) {
        foreach ($item in Get-ChildItem -LiteralPath $directories.Pop() -Force) {
            if ($item.Attributes -band [IO.FileAttributes]::ReparsePoint) { throw 'Do not delete through a job workspace link.' }
            if ($item.PSIsContainer) { $directories.Push($item.FullName) }
        }
    }
    # Runner.Worker and this hook can still hold the checkout as their working directory.
    foreach ($item in Get-ChildItem -LiteralPath $Path -Force) {
        Remove-Item -LiteralPath $item.FullName -Recurse -Force
    }
    if (@(Get-ChildItem -LiteralPath $Path -Force).Count) { throw 'Job workspace contents remain after cleanup.' }
}

. (Join-Path $PSScriptRoot 'ExcelRunnerPolicy.ps1')
$recordPath = Join-Path $env:LOCALAPPDATA "ExcelMcp\Jobs\$env:GITHUB_RUN_ID-$env:GITHUB_RUN_ATTEMPT.json"
$record = Get-Content -LiteralPath $recordPath -Raw | ConvertFrom-Json
if ($record.runId -ne $env:GITHUB_RUN_ID -or $record.attempt -ne $env:GITHUB_RUN_ATTEMPT) { throw 'Job cleanup record has the wrong identity.' }
$boot = (Get-CimInstance Win32_OperatingSystem).LastBootUpTime.ToUniversalTime()
$sameBoot = (ConvertTo-ExcelRunnerUtc $record.bootTime) -eq $boot
if ($sameBoot -and $record.session -ne (Get-Process -Id $PID).SessionId) { throw 'Job cleanup has a different desktop session.' }
Assert-ExcelRunnerWorkspace $env:GITHUB_WORKSPACE
$snapshot = if ($sameBoot) { @(Get-CimInstance Win32_Process -Filter "SessionId=$($record.session)") } else { @() }
$ancestors = [Collections.Generic.HashSet[uint32]]::new()
[uint32]$ancestor = $PID
while ($ancestor -gt 0 -and $ancestors.Add($ancestor)) {
    $parent = @($snapshot | Where-Object ProcessId -EQ $ancestor)
    if ($parent.Count -ne 1) { break }
    $ancestor = $parent[0].ParentProcessId
}
$failures = [Collections.Generic.List[Exception]]::new()
foreach ($entry in $snapshot) {
    if ($ancestors.Contains($entry.ProcessId)) { continue }
    $stamp = $entry.CreationDate.ToUniversalTime().ToString('o')
    if (@($record.processes | Where-Object { $_.id -eq $entry.ProcessId -and $_.startedAt -eq $stamp }).Count) { continue }
    try {
        Stop-ExcelRunnerJobProcess $entry $env:COMPUTERNAME
    }
    catch { $failures.Add($_.Exception) }
}
if ($failures.Count) { throw [AggregateException]::new('Job cleanup failed; keep the runner quarantined.', $failures) }
Remove-ExcelRunnerJobWorkspace $env:GITHUB_WORKSPACE
Remove-Item -LiteralPath $recordPath -Force
