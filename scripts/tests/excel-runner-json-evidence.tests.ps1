$ErrorActionPreference = 'Stop'
. (Join-Path (Split-Path -Parent $PSScriptRoot) 'ExcelRunnerPolicy.ps1')
$now = [DateTime]::UtcNow
$state = @{
    state = 'ready'; license = 'Licensed'; formula = 42; persistedFormula = 42
    checkedAt = $now.ToString('o'); bootTime = $now.AddMinutes(-5).ToString('o')
}
$json = $state | ConvertTo-Json -Compress
$parsed = $json | ConvertFrom-Json
Assert-ExcelRunnerReadiness $parsed -Now $now -BootTime $now.AddMinutes(-5)
$failed = $false
try { ConvertTo-ExcelRunnerUtc ([DateTime]::SpecifyKind($now, [DateTimeKind]::Unspecified)) }
catch { $failed = $true }
if (-not $failed) { throw 'Typed JSON timestamps must still have explicit UTC identity.' }
Write-Output 'Actual JSON evidence survives native date conversion without accepting ambiguous timestamps.'
