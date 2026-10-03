<#
.SYNOPSIS
Registers the repository runner in the non-admin interactive profile, without starting it.
.DESCRIPTION
Adapts mcp-windows runner setup (MIT, Copyright (c) 2025 Sbroenne).
Only a one-hour registration token is transported, encrypted to a transient guest key.
No GitHub administration token or Azure control credential is sent to the VM.
#>
param(
    [ValidateSet('Prepare', 'Configure')][string]$Action,
    [ValidatePattern('^[a-f0-9]{32}$')][string]$OperationId
)
$ErrorActionPreference = 'Stop'

function Assert-ExcelRunnerPackage {
    param($Asset, [string]$Version)
    if ($Version -notmatch '^\d+\.\d+\.\d+$' -or
        $Asset.name -ne "actions-runner-win-x64-$Version.zip" -or
        $Asset.browser_download_url -ne "https://github.com/actions/runner/releases/download/v$Version/actions-runner-win-x64-$Version.zip" -or
        $Asset.digest -notmatch '^sha256:[a-f0-9]{64}$') {
        throw 'Runner package requires the official versioned Windows x64 URL and published SHA256 digest.'
    }

}

function New-ExcelRunnerRegistrationKey {
    Add-Type -AssemblyName System.Security
    $key = [Security.Cryptography.RSACryptoServiceProvider]::new(2048)
    $key.PersistKeyInCsp = $false
    return $key
}

if ($MyInvocation.InvocationName -eq '.') { return }
if (-not $Action -or -not $OperationId) { throw 'A registration action and unique operation ID are required.' }
$identity = [Security.Principal.WindowsIdentity]::GetCurrent()
try {
    if ($identity.Name -ine "$env:COMPUTERNAME\excelrunner" -or
        ([Security.Principal.WindowsPrincipal]::new($identity)).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator) -or
        -not [Environment]::UserInteractive -or (Get-Process -Id $PID).SessionId -le 0) {
        throw 'Registration requires the dedicated non-admin interactive account.'
    }
}
finally { $identity.Dispose() }
if (@(Get-Process -Name Runner.Listener, Runner.Worker -ErrorAction SilentlyContinue).Count) {
    throw 'Registration must not overlap coding work.'
}
$directory = Join-Path $env:LOCALAPPDATA 'ExcelMcp\Registration'
New-Item -ItemType Directory -Path $directory -Force | Out-Null
$report = Join-Path $directory "$OperationId.json"
$keyPath = Join-Path $directory "$OperationId.key"
$entropy = [Text.Encoding]::UTF8.GetBytes($OperationId)
$state = @{ operationId = $OperationId; state = 'failed' }
$rsa = New-ExcelRunnerRegistrationKey
$keyBytes = $null
$plainToken = $null
$process = $null
try {
    switch ($Action) {
        'Prepare' {
            if ((Test-Path -LiteralPath $keyPath) -or (Test-Path -LiteralPath $report)) { throw 'Registration operation already exists.' }
            $keyBytes = [Text.Encoding]::UTF8.GetBytes($rsa.ToXmlString($true))
            $protected = [Security.Cryptography.ProtectedData]::Protect(
                $keyBytes, $entropy, [Security.Cryptography.DataProtectionScope]::CurrentUser)
            [IO.File]::WriteAllBytes($keyPath, $protected)
            $public = $rsa.ExportParameters($false)
            $state.modulus = [Convert]::ToBase64String($public.Modulus)
            $state.exponent = [Convert]::ToBase64String($public.Exponent)
            $state.state = 'prepared'
        }
        'Configure' {
            $request = Get-Content "C:\ProgramData\ExcelMcp\Desktop\registration-$OperationId.json" -Raw | ConvertFrom-Json
            if ($request.operationId -ne $OperationId -or $request.repository -ne 'sbroenne/mcp-server-excel') {
                throw 'Registration request belongs to another operation or repository.'
            }
            if (Test-Path -LiteralPath 'C:\actions-runner\.runner') { throw 'Do not replace an existing runner registration.' }
            $keyBytes = [Security.Cryptography.ProtectedData]::Unprotect(
                [IO.File]::ReadAllBytes($keyPath), $entropy, [Security.Cryptography.DataProtectionScope]::CurrentUser)
            $rsa.FromXmlString([Text.Encoding]::UTF8.GetString($keyBytes))
            $plainToken = $rsa.Decrypt([Convert]::FromBase64String($request.encryptedToken),
                [Security.Cryptography.RSAEncryptionPadding]::OaepSHA1)
            $info = [Diagnostics.ProcessStartInfo]::new()
            $info.FileName = 'C:\actions-runner\bin\Runner.Listener.exe'
            $info.WorkingDirectory = 'C:\actions-runner'
            $info.Arguments = 'configure --unattended --url https://github.com/sbroenne/mcp-server-excel --name azure-excel-copilot --labels excel-copilot --no-default-labels --work _work'
            $info.UseShellExecute = $false
            $info.CreateNoWindow = $true
            $info.RedirectStandardOutput = $true
            $info.RedirectStandardError = $true
            $info.EnvironmentVariables['ACTIONS_RUNNER_INPUT_TOKEN'] = [Text.Encoding]::UTF8.GetString($plainToken)
            $process = [Diagnostics.Process]::Start($info)
            $null = $process.Handle
            $startedAt = $process.StartTime
            $stdout = $process.StandardOutput.ReadToEndAsync()
            $stderr = $process.StandardError.ReadToEndAsync()
            if (-not $process.WaitForExit(300000)) {
                $current = Get-Process -Id $process.Id -ErrorAction SilentlyContinue
                try {
                    if ($current -and $current.StartTime -eq $startedAt) { Stop-Process -Id $current.Id -Force }
                }
                finally { if ($current) { $current.Dispose() } }
                throw 'Runner configuration exceeded its deadline.'
            }
            $info.EnvironmentVariables.Remove('ACTIONS_RUNNER_INPUT_TOKEN')
            $stdout.GetAwaiter().GetResult() | Set-Content (Join-Path $directory "$OperationId.stdout.log") -Encoding UTF8
            $stderr.GetAwaiter().GetResult() | Set-Content (Join-Path $directory "$OperationId.stderr.log") -Encoding UTF8
            if ($process.ExitCode -ne 0) { throw "Runner configuration failed with exit code $($process.ExitCode)." }
            $configuration = Get-Content 'C:\actions-runner\.runner' -Raw | ConvertFrom-Json
            if ($configuration.agentName -ne 'azure-excel-copilot' -or
                $configuration.gitHubUrl.TrimEnd('/') -ne 'https://github.com/sbroenne/mcp-server-excel' -or
                $configuration.agentId -le 0) { throw 'Unexpected registered runner identity.' }
            $state.state = 'configured'
            $state.agentId = $configuration.agentId
        }
    }
}
catch {
    $state.errorType = $_.Exception.GetType().FullName
    $state.errorId = $_.FullyQualifiedErrorId
    $state.errorLine = $_.InvocationInfo.ScriptLineNumber
    throw
}
finally {
    if ($info) { $info.EnvironmentVariables.Remove('ACTIONS_RUNNER_INPUT_TOKEN') }
    if ($Action -eq 'Configure' -or $state.state -eq 'failed') {
        if (Test-Path -LiteralPath $keyPath) { Remove-Item -LiteralPath $keyPath -Force }
    }
    if ($keyBytes) { [Array]::Clear($keyBytes, 0, $keyBytes.Length) }
    if ($plainToken) { [Array]::Clear($plainToken, 0, $plainToken.Length) }
    if ($process) { $process.Dispose() }
    $rsa.Dispose()
    $state | ConvertTo-Json -Compress | Set-Content -LiteralPath "$report.tmp" -Encoding UTF8
    if (Test-Path -LiteralPath $report) { [IO.File]::Replace("$report.tmp", $report, [NullString]::Value) }
    else { [IO.File]::Move("$report.tmp", $report) }
}
