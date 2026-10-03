$ErrorActionPreference = 'Stop'
. (Join-Path (Split-Path -Parent $PSScriptRoot) 'ExcelRunnerPolicy.ps1')
$now = [DateTime]::UtcNow
$run = @{
    head_branch = 'main'; head_repository = @{ full_name = 'synthetic/repository' }
    path = '.github/workflows/excel-runner-maintenance.yml'
    status = 'completed'; conclusion = 'success'
    created_at = $now.AddHours(-1).ToString('o'); updated_at = $now.ToString('o')
}
Assert-ExcelRunnerMaintenanceHistory @($run) 'synthetic/repository' 'main' -Now $now
foreach ($field in @('head_branch', 'path', 'status', 'conclusion', 'updated_at')) {
    $bad = $run.Clone()
    $bad[$field] = if ($field -eq 'updated_at') { $now.AddDays(-9).ToString('o') } else { 'untrusted-or-failed' }
    $failed = $false
    try { Assert-ExcelRunnerMaintenanceHistory @($bad) 'synthetic/repository' 'main' -Now $now }
    catch { $failed = $true }
    if (-not $failed) { throw "Unsafe maintenance history was accepted: $field" }
}
$failedRun = $run.Clone()
$failedRun.created_at = $now.ToString('o')
$failedRun.conclusion = 'failure'
$failed = $false
try { Assert-ExcelRunnerMaintenanceHistory @($failedRun, $run) 'synthetic/repository' 'main' -Now $now }
catch { $failed = $true }
if (-not $failed) { throw 'An older success must not conceal the most recent failed maintenance.' }
Write-Output 'Protected maintenance history blocks stale or failed repeated VM wakes.'
