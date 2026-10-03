<#
.SYNOPSIS
Installs 64-bit retail Excel 2024 on the dedicated development VM.
.DESCRIPTION
Uses the same SYSTEM scheduled-worker pattern as mcp-windows maintenance.
Installation is separate from activation and interactive Excel qualification.
#>
param(
    [ValidateSet('Start', 'Status', 'Worker')]
    [string]$Action
)

$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'

function Assert-MicrosoftOfficeSignature {
    param([string]$Path)
    $signature = Get-AuthenticodeSignature -LiteralPath $Path
    if ($signature.Status -ne 'Valid' -or
        $signature.SignerCertificate.Subject -notmatch '^CN=Microsoft Corporation,') {
        throw 'The Office deployment executable must have a valid Microsoft Corporation signature.'
    }
}

function Get-VerifiedExcelInstallation {
    $configuration = Get-ItemProperty -LiteralPath 'HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration'
    $products = @($configuration.ProductReleaseIds -split ',')
    if ($products -notcontains 'Excel2024Retail' -or $configuration.Platform -ne 'x64') {
        throw 'The installed product must include Excel2024Retail, 64-bit.'
    }
    if (-not $configuration.VersionToReport -or -not $configuration.InstallationPath) {
        throw 'The installed Office version or installation path is missing.'
    }
    $excelPath = Join-Path $configuration.InstallationPath 'root\Office16\EXCEL.EXE'
    if (-not (Test-Path -LiteralPath $excelPath -PathType Leaf)) {
        throw 'Excel is missing from the reported Office installation.'
    }
    return @{
        product = 'Excel2024Retail'
        platform = 'x64'
        version = $configuration.VersionToReport
    }
}

function Invoke-OfficeSetupProcess {
    param([string]$Path, [string]$Arguments, [int]$TimeoutSeconds = 1200)
    $startInfo = New-Object Diagnostics.ProcessStartInfo
    $startInfo.FileName = $Path
    $startInfo.Arguments = $Arguments
    $startInfo.UseShellExecute = $false
    $startInfo.CreateNoWindow = $true
    $process = [Diagnostics.Process]::Start($startInfo)
    try {
        # Retain the handle before exit; PowerShell's Start-Process can lose ExitCode.
        $null = $process.Handle
        $startTime = $process.StartTime
        if (-not $process.WaitForExit($TimeoutSeconds * 1000)) {
            $owned = Get-Process -Id $process.Id -ErrorAction SilentlyContinue
            if ($owned -and $owned.StartTime -eq $startTime) {
                Stop-Process -Id $owned.Id -Force
                if (-not $process.WaitForExit(10000)) { throw 'Timed-out installer could not be stopped.' }
            }
            $failure = [TimeoutException]::new('Office installer exceeded its deadline; inspect Click-to-Run before retrying.')
            $failure.Data['ProcessId'] = $process.Id
            $failure.Data['StartTime'] = $startTime
            throw $failure
        }
        return $process.ExitCode
    }
    finally { $process.Dispose() }
}

function Write-OfficeInstallState {
    param([hashtable]$State)
    $State.checkedAt = [DateTime]::UtcNow.ToString('o')
    $State | ConvertTo-Json -Compress | Set-Content -LiteralPath "$statePath.tmp" -Encoding UTF8
    if (Test-Path -LiteralPath $statePath) {
        [IO.File]::Replace("$statePath.tmp", $statePath, [NullString]::Value)
    }
    else { [IO.File]::Move("$statePath.tmp", $statePath) }
}

if ($MyInvocation.InvocationName -eq '.') { return }
if (-not $Action) { throw 'An explicit installation action is required.' }

$directory = 'C:\ProgramData\ExcelMcp\Provisioning'
$statePath = Join-Path $directory 'office-install.json'
$taskName = 'ExcelMcp-Install-Office'

switch ($Action) {
    'Start' {
        $task = Get-ScheduledTask -TaskName $taskName -ErrorAction SilentlyContinue
        if ($task -and $task.State -eq 'Running') { throw 'Excel installation is already running.' }
        if (@(Get-Process -Name Runner.Listener -ErrorAction SilentlyContinue).Count -gt 0) {
            throw 'Office installation requires the coding runner to be stopped.'
        }
        if (@(Get-Process -Name EXCEL -ErrorAction SilentlyContinue).Count -gt 0) {
            throw 'Close Excel before installation; do not interrupt an active workbook.'
        }
        if (Test-Path -LiteralPath 'HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration') {
            $installation = Get-VerifiedExcelInstallation
            Write-OfficeInstallState @{
                state = 'installed'
                installation = $installation
                activation = 'not-verified'
            }
        }
        else {
            Write-OfficeInstallState @{ state = 'running' }
            $taskAction = New-ScheduledTaskAction -Execute 'powershell.exe' -Argument (
                "-NoProfile -NonInteractive -ExecutionPolicy Bypass -File `"$PSCommandPath`" -Action Worker"
            )
            $principal = New-ScheduledTaskPrincipal -UserId 'SYSTEM' -LogonType ServiceAccount -RunLevel Highest
            $settings = New-ScheduledTaskSettingsSet -ExecutionTimeLimit (New-TimeSpan -Minutes 30)
            Register-ScheduledTask -TaskName $taskName -Action $taskAction -Principal $principal `
                -Settings $settings -Force | Out-Null
            Start-ScheduledTask -TaskName $taskName
        }
    }
    'Worker' {
        try {
            $setup = Join-Path $directory 'office-deployment-setup.exe'
            $configurationPath = Join-Path $directory 'excel-retail.xml'
            Invoke-WebRequest -Uri 'https://officecdn.microsoft.com/pr/wsus/setup.exe' `
                -OutFile $setup -UseBasicParsing
            Assert-MicrosoftOfficeSignature -Path $setup
            @'
<Configuration>
  <Add OfficeClientEdition="64" Channel="Current">
    <Product ID="Excel2024Retail">
      <Language ID="en-us" />
    </Product>
  </Add>
  <Display Level="None" AcceptEULA="TRUE" />
  <Updates Enabled="TRUE" Channel="Current" />
</Configuration>
'@ | Set-Content -LiteralPath $configurationPath -Encoding UTF8
            $exitCode = Invoke-OfficeSetupProcess -Path $setup -Arguments "/configure `"$configurationPath`""
            if ($exitCode -notin @(0, 3010)) {
                throw "Office installation failed with exit code $exitCode."
            }
            $installation = Get-VerifiedExcelInstallation
            Write-OfficeInstallState @{
                state = 'installed'
                installation = $installation
                rebootRequired = $exitCode -eq 3010
                activation = 'not-verified'
            }
        }
        catch {
            Write-OfficeInstallState @{ state = 'failed'; error = $_.Exception.Message }
            throw
        }
        finally {
            foreach ($path in @($setup, $configurationPath)) {
                if ($path -and (Test-Path -LiteralPath $path)) { Remove-Item -LiteralPath $path -Force }
            }
        }
        return
    }
    'Status' {
        if (-not (Test-Path -LiteralPath $statePath)) { throw 'No Office installation state exists.' }
        $state = Get-Content -LiteralPath $statePath -Raw | ConvertFrom-Json
        if ($state.state -eq 'running') {
            $task = Get-ScheduledTask -TaskName $taskName
            if ($task.State -ne 'Running') {
                $info = Get-ScheduledTaskInfo -TaskName $taskName
                Write-OfficeInstallState @{
                    state = 'failed'
                    error = "Installation stopped without a result; task exit code=$($info.LastTaskResult)."
                }
            }
        }
    }
}

Write-Output ('EXCELMCP_OFFICE=' + (Get-Content -LiteralPath $statePath -Raw).Trim())
