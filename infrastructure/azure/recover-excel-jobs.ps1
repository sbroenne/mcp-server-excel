param([ValidatePattern('^[a-f0-9]{32}$')][string]$OperationId)
$ErrorActionPreference = 'Stop'
$identity = [Security.Principal.WindowsIdentity]::GetCurrent()
try {
    if (-not $OperationId -or -not [Environment]::UserInteractive -or
        (Get-Process -Id $PID).SessionId -le 0 -or
        $identity.Name -ine "$env:COMPUTERNAME\excelrunner" -or
        ([Security.Principal.WindowsPrincipal]::new($identity)).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
        throw 'Recovery requires a unique operation and the limited interactive account.'
    }
}
finally { $identity.Dispose() }
if (@(Get-Process -Name Runner.Listener, Runner.Worker -ErrorAction SilentlyContinue).Count) {
    throw 'Recovery must not overlap coding work.'
}
$directory = Join-Path $env:LOCALAPPDATA 'ExcelMcp\Recovery'
New-Item -ItemType Directory -Path $directory -Force | Out-Null
$state = @{ operationId = $OperationId; state = 'failed' }
try {
    $records = Join-Path $env:LOCALAPPDATA 'ExcelMcp\Jobs'
    if (Test-Path -LiteralPath $records) {
        foreach ($file in Get-ChildItem -LiteralPath $records -File -Filter '*.json') {
            $record = Get-Content -LiteralPath $file.FullName -Raw | ConvertFrom-Json
            if ($record.runId -notmatch '^\d+$' -or $record.attempt -notmatch '^\d+$' -or
                $file.Name -ne "$($record.runId)-$($record.attempt).json") { throw 'Unexpected cleanup record.' }
            $env:GITHUB_RUN_ID = "$($record.runId)"
            $env:GITHUB_RUN_ATTEMPT = "$($record.attempt)"
            $env:GITHUB_WORKSPACE = 'C:\actions-runner\_work\mcp-server-excel\mcp-server-excel'
            & (Join-Path $PSScriptRoot 'excel-job-completed.ps1')
        }
    }
    $state.state = 'recovered'
}
catch { $state.errorType = $_.Exception.GetType().FullName; throw }
finally {
    $state | ConvertTo-Json -Compress | Set-Content -LiteralPath (Join-Path $directory "$OperationId.json") -Encoding UTF8
}
