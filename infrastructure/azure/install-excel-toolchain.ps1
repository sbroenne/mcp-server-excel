<#
.SYNOPSIS
Installs development prerequisites without registering a coding runner.
.DESCRIPTION
Adapts mcp-windows setup-runner.ps1 at a76b3eeb075419986d0b00721bae5cdc896c2f81.
Copyright (c) 2025 Sbroenne. MIT licence; see LICENSE in this repository.
Uses the existing bounded SYSTEM installer worker and protected directory.
#>
param(
    [ValidateSet('Start', 'Status', 'Worker')]
    [string]$Action,
    [ValidatePattern('^\d+\.\d+\.\d+$')]
    [string]$SdkVersion,
    [ValidateSet('latestFeature', 'disable')]
    [string]$RollForward = 'latestFeature'
)
$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'

function Assert-ToolchainInstallerSignature {
    param([string]$Path, [ValidateSet('Microsoft', 'DotNet', 'Git', 'Node')][string]$Publisher)
    $signature = Get-AuthenticodeSignature -LiteralPath $Path
    $subject = switch ($Publisher) {
        Microsoft { '^CN=Microsoft Corporation,' }
        DotNet { '^CN=(\.NET|Microsoft Corporation),' }
        Git { '^CN=(Open Source Developer, )?Johannes Schindelin,' }
        Node { '^CN=OpenJS Foundation,' }
    }
    $publisherMatches = $signature.SignerCertificate.Subject -match $subject
    if ($Publisher -eq 'DotNet') {
        $publisherMatches = $publisherMatches -and
            $signature.SignerCertificate.Subject -match '(^|, )O=Microsoft Corporation(,|$)'
    }
    if ($signature.Status -ne 'Valid' -or -not $publisherMatches) {
        $microsoftOrganization = $signature.SignerCertificate.Subject -match '(^|, )O=Microsoft Corporation(,|$)'
        $dotnetCommonName = $signature.SignerCertificate.Subject -match '^CN=\.NET(,|$)'
        throw ("The $Publisher installer must have a valid signature from the expected publisher. " +
            "status=$($signature.Status); publisherMatches=$publisherMatches; " +
            "microsoftOrganization=$microsoftOrganization; dotnetCommonName=$dotnetCommonName")
    }
}

function Assert-RequiredRunnerSdk {
    param([string]$Required, [string[]]$Installed,
        [ValidateSet('latestFeature', 'disable')][string]$RollForward = 'latestFeature')
    $minimum = [version]$Required
    $compatible = @(
        foreach ($line in $Installed) {
            if ($line -notmatch '^(\d+\.\d+\.\d+)\s') { continue }
            $candidate = [version]$Matches[1]
            if (($RollForward -eq 'disable' -and $candidate -eq $minimum) -or
                ($RollForward -eq 'latestFeature' -and $candidate.Major -eq $minimum.Major -and
                 $candidate.Minor -eq $minimum.Minor -and $candidate -ge $minimum)) { $line }
        }
    )
    if (-not $compatible.Count) {
        throw "No installed .NET SDK satisfies $Required with rollForward=$RollForward."
    }
}

function Get-RunnerJqRelease {
    @{
        version = 'jq-1.8.2'
        uri = 'https://github.com/jqlang/jq/releases/download/jq-1.8.2/jq-windows-amd64.exe'
        sha256 = 'a6fc67fedaf9128a3309a1e2ebb8b986aeccf70122ee46d2cb4849e423f0c627'
    }
}

function Assert-RunnerJqPackage {
    param([string]$Path)
    if ((Get-FileHash -LiteralPath $Path -Algorithm SHA256).Hash -ne (Get-RunnerJqRelease).sha256) {
        throw 'The jq executable must match the pinned official Windows x64 release checksum.'
    }
}

function Get-RunnerCloudToolState {
    $developmentMode = Get-ItemPropertyValue -LiteralPath 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\AppModelUnlock' `
        -Name AllowDevelopmentWithoutDevLicense -ErrorAction Stop
    if ($developmentMode -ne 1) { throw 'Windows Developer Mode is required for limited-user runtime symbolic-link extraction.' }
    $bash = Join-Path $env:ProgramFiles 'Git\bin\bash.exe'
    $jq = Join-Path $env:ProgramFiles 'ExcelMcp\Tools\jq.exe'
    foreach ($path in @($bash, $jq)) {
        if (-not (Test-Path -LiteralPath $path -PathType Leaf)) { throw 'A required cloud initialization tool is missing.' }
    }
    Assert-RunnerJqPackage $jq
    if ((Get-Command bash -CommandType Application -ErrorAction Stop).Source -ine $bash -or
        (Get-Command jq -CommandType Application -ErrorAction Stop).Source -ine $jq) {
        throw 'The runner PATH must resolve the protected Git Bash and pinned jq executables.'
    }
    $bashOutput = @(& $bash --version)
    if ($LASTEXITCODE -ne 0 -or -not $bashOutput.Count -or $bashOutput[0] -notmatch '^GNU bash, version \d+\.\d+') {
        throw 'Git Bash verification failed.'
    }
    $bashVersion = $bashOutput[0]
    $jqVersion = (& $bash --noprofile --norc -c 'jq --version') -join ''
    if ($LASTEXITCODE -ne 0 -or $jqVersion -ne (Get-RunnerJqRelease).version) {
        throw 'The cloud initialization Bash shell cannot execute the required jq version.'
    }
    return @{ bash = $bashVersion; jq = $jqVersion; developmentMode = $true }
}

function Set-RunnerCloudToolPath {
    $bashDirectory = Join-Path $env:ProgramFiles 'Git\bin'
    $toolsDirectory = Join-Path $env:ProgramFiles 'ExcelMcp\Tools'
    $existing = @([Environment]::GetEnvironmentVariable('Path', 'Machine') -split ';' |
        Where-Object { $_ -and $_.TrimEnd('\') -ine $bashDirectory -and $_.TrimEnd('\') -ine $toolsDirectory })
    [Environment]::SetEnvironmentVariable('Path', (@($bashDirectory, $toolsDirectory) + $existing) -join ';', 'Machine')
    $env:Path = [Environment]::GetEnvironmentVariable('Path', 'Machine') + ';' +
        [Environment]::GetEnvironmentVariable('Path', 'User')
}

function Install-RunnerCloudPrerequisites {
    $bashDirectory = Join-Path $env:ProgramFiles 'Git\bin'
    if (-not (Test-Path -LiteralPath (Join-Path $bashDirectory 'bash.exe') -PathType Leaf)) {
        throw 'Install the verified Git for Windows package before cloud initialization tools.'
    }
    $toolsDirectory = Join-Path $env:ProgramFiles 'ExcelMcp\Tools'
    New-Item -ItemType Directory -Path $toolsDirectory -Force | Out-Null
    $jq = Join-Path $toolsDirectory 'jq.exe'
    if (-not (Test-Path -LiteralPath $jq -PathType Leaf)) {
        $download = Join-Path $toolsDirectory "jq-download-$([Guid]::NewGuid().ToString('N')).exe"
        try {
            Invoke-WebRequest -Uri (Get-RunnerJqRelease).uri -OutFile $download -UseBasicParsing -TimeoutSec 180
            Assert-RunnerJqPackage $download
            [IO.File]::Move($download, $jq)
        }
        finally { if (Test-Path -LiteralPath $download) { Remove-Item -LiteralPath $download -Force } }
    }
    Assert-RunnerJqPackage $jq
    $developmentSettings = 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\AppModelUnlock'
    New-Item -Path $developmentSettings -Force | Out-Null
    New-ItemProperty -LiteralPath $developmentSettings -Name AllowDevelopmentWithoutDevLicense `
        -PropertyType DWord -Value 1 -Force | Out-Null
    Set-RunnerCloudToolPath
    Get-RunnerCloudToolState
}

function Get-RunnerToolchainState {
    $dotnet = Join-Path $env:ProgramFiles 'dotnet\dotnet.exe'
    $git = Join-Path $env:ProgramFiles 'Git\cmd\git.exe'
    $pwsh = Join-Path $env:ProgramFiles 'PowerShell\7\pwsh.exe'
    $node = Join-Path $env:ProgramFiles 'nodejs\node.exe'
    foreach ($path in @($dotnet, $git, $pwsh, $node)) {
        if (-not (Test-Path -LiteralPath $path -PathType Leaf)) { throw 'A required development tool is missing.' }
    }
    $sdks = @(& $dotnet --list-sdks)
    if ($LASTEXITCODE -ne 0) { throw '.NET SDK enumeration failed.' }
    Assert-RequiredRunnerSdk $SdkVersion $sdks -RollForward $RollForward
    Push-Location $directory
    try {
        $selectedSdk = (& $dotnet --version) -join ''
        if ($LASTEXITCODE -ne 0) { throw 'The SDK resolver could not honor global.json.' }
    }
    finally { Pop-Location }
    Assert-RequiredRunnerSdk $SdkVersion @("$selectedSdk [resolved]") -RollForward $RollForward
    $gitVersion = (& $git --version) -join ''
    if ($LASTEXITCODE -ne 0) { throw 'Git verification failed.' }
    $powershellVersion = (& $pwsh -NoProfile -Command '$PSVersionTable.PSVersion.ToString()') -join ''
    if ($LASTEXITCODE -ne 0 -or $powershellVersion -notmatch '^7\.') { throw 'PowerShell 7 verification failed.' }
    $nodeVersion = (& $node --version) -join ''
    if ($LASTEXITCODE -ne 0 -or $nodeVersion -notmatch '^v22\.') { throw 'Node.js 22 verification failed.' }
    $cloud = Get-RunnerCloudToolState
    return @{
        sdk = $selectedSdk; requiredSdk = $SdkVersion; rollForward = $RollForward
        git = $gitVersion; powershell = $powershellVersion; node = $nodeVersion
        bash = $cloud.bash; jq = $cloud.jq; developmentMode = $cloud.developmentMode
    }
}

function Write-ToolchainState {
    param([hashtable]$State)
    $State.checkedAt = [DateTime]::UtcNow.ToString('o')
    $State | ConvertTo-Json -Depth 5 -Compress | Set-Content -LiteralPath "$statePath.tmp" -Encoding UTF8
    if (Test-Path -LiteralPath $statePath) { [IO.File]::Replace("$statePath.tmp", $statePath, [NullString]::Value) }
    else { [IO.File]::Move("$statePath.tmp", $statePath) }
}

function Install-RunnerPrerequisite {
    param([string]$Uri, [string]$FileName, [string]$Publisher, [string]$Arguments, [switch]$Msi)
    $installer = Join-Path $directory $FileName
    try {
        Invoke-WebRequest -Uri $Uri -OutFile $installer -UseBasicParsing -TimeoutSec 180
        Assert-ToolchainInstallerSignature $installer $Publisher
        if ($Msi) {
            $exitCode = Invoke-OfficeSetupProcess -Path (Join-Path $env:SystemRoot 'System32\msiexec.exe') `
                -Arguments "/i `"$installer`" /qn /norestart $Arguments" -TimeoutSeconds 600
        }
        else { $exitCode = Invoke-OfficeSetupProcess -Path $installer -Arguments $Arguments -TimeoutSeconds 600 }
        if ($exitCode -notin @(0, 3010)) { throw "Prerequisite installation failed with exit code $exitCode." }
        if ($exitCode -eq 3010) { $script:toolchainRebootRequired = $true }
    }
    finally { if (Test-Path -LiteralPath $installer) { Remove-Item -LiteralPath $installer -Force } }
}

if ($MyInvocation.InvocationName -eq '.') { return }
if (-not $Action -or -not $SdkVersion) { throw 'An explicit action and repository SDK version are required.' }
$requestedAction = $Action
. (Join-Path $PSScriptRoot 'install-excel-office.ps1')
$Action = $requestedAction
$directory = 'C:\ProgramData\ExcelMcp\Provisioning'
$statePath = Join-Path $directory 'toolchain-install.json'
$taskName = 'ExcelMcp-Install-Toolchain'

switch ($Action) {
    Start {
        $task = Get-ScheduledTask -TaskName $taskName -ErrorAction SilentlyContinue
        if ($task -and $task.State -eq 'Running') { throw 'Toolchain installation is already running.' }
        if ((Test-Path -LiteralPath 'C:\actions-runner\.runner') -or
            @(Get-Process -Name Runner.Listener -ErrorAction SilentlyContinue).Count -gt 0) {
            throw 'Provisioning requires an unregistered, stopped coding runner.'
        }
        Write-ToolchainState @{ state = 'running'; sdk = $SdkVersion }
        $taskAction = New-ScheduledTaskAction -Execute 'powershell.exe' -Argument (
            "-NoProfile -NonInteractive -ExecutionPolicy Bypass -File `"$PSCommandPath`" -Action Worker -SdkVersion $SdkVersion -RollForward $RollForward"
        )
        $principal = New-ScheduledTaskPrincipal -UserId SYSTEM -LogonType ServiceAccount -RunLevel Highest
        $settings = New-ScheduledTaskSettingsSet -ExecutionTimeLimit (New-TimeSpan -Minutes 45)
        Register-ScheduledTask -TaskName $taskName -Action $taskAction -Principal $principal -Settings $settings -Force | Out-Null
        Start-ScheduledTask -TaskName $taskName
    }
    Worker {
        try {
            $script:toolchainRebootRequired = $false
            $targetSdk = $SdkVersion
            if ($RollForward -eq 'latestFeature') {
                $minimum = [version]$SdkVersion
                $channel = "$($minimum.Major).$($minimum.Minor)"
                $releaseMetadata = Invoke-RestMethod "https://builds.dotnet.microsoft.com/dotnet/release-metadata/$channel/releases.json" -TimeoutSec 60
                $targetSdk = $releaseMetadata.'latest-sdk'
                if ($targetSdk -notmatch '^\d+\.\d+\.\d+$' -or
                    ([version]$targetSdk).Major -ne $minimum.Major -or
                    ([version]$targetSdk).Minor -ne $minimum.Minor -or ([version]$targetSdk) -lt $minimum) {
                    throw 'The stable SDK release does not satisfy global.json.'
                }
            }
            $dotnet = Join-Path $env:ProgramFiles 'dotnet\dotnet.exe'
            $sdks = if (Test-Path -LiteralPath $dotnet) {
                $installed = @(& $dotnet --list-sdks)
                if ($LASTEXITCODE -ne 0) { throw 'Existing .NET SDK enumeration failed.' }
                $installed
            } else { @() }
            if (-not @($sdks | Where-Object { $_ -match ('^' + [regex]::Escape($targetSdk) + '\s') }).Count) {
                Install-RunnerPrerequisite "https://builds.dotnet.microsoft.com/dotnet/Sdk/$targetSdk/dotnet-sdk-$targetSdk-win-x64.exe" `
                    'dotnet-sdk.exe' DotNet '/quiet /norestart'
            }
            $headers = @{ 'User-Agent' = 'ExcelMcp-Runner-Setup' }
            if (-not (Test-Path -LiteralPath (Join-Path $env:ProgramFiles 'Git\cmd\git.exe'))) {
                $release = Invoke-RestMethod 'https://api.github.com/repos/git-for-windows/git/releases/latest' -Headers $headers -TimeoutSec 60
                $assets = @($release.assets | Where-Object name -Match '^Git-.*-64-bit\.exe$')
                if ($assets.Count -ne 1) { throw 'Expected exactly one official Git x64 installer.' }
                Install-RunnerPrerequisite $assets[0].browser_download_url 'git.exe' Git '/VERYSILENT /NORESTART /NOCANCEL /SP- /ALLUSERS'
            }
            if (-not (Test-Path -LiteralPath (Join-Path $env:ProgramFiles 'PowerShell\7\pwsh.exe'))) {
                $release = Invoke-RestMethod 'https://api.github.com/repos/PowerShell/PowerShell/releases/latest' -Headers $headers -TimeoutSec 60
                $assets = @($release.assets | Where-Object name -Match '^PowerShell-.*-win-x64\.msi$')
                if ($assets.Count -ne 1) { throw 'Expected exactly one official PowerShell x64 installer.' }
                Install-RunnerPrerequisite $assets[0].browser_download_url 'powershell.msi' Microsoft 'ADD_PATH=1' -Msi
            }
            $node = Join-Path $env:ProgramFiles 'nodejs\node.exe'
            $nodeMatches = $false
            if (Test-Path -LiteralPath $node) {
                $current = (& $node --version) -join ''
                if ($LASTEXITCODE -ne 0) { throw 'Existing Node.js version check failed.' }
                $nodeMatches = $current -match '^v22\.'
                if (-not $nodeMatches) { throw 'An unexpected Node.js major version is installed; do not silently replace it.' }
            }
            if (-not $nodeMatches) {
                $releases = Invoke-RestMethod 'https://nodejs.org/dist/index.json' -TimeoutSec 60
                $release = $releases | Where-Object {
                    $_.version -match '^v22\.' -and $_.lts -is [string] -and $_.files -contains 'win-x64-msi'
                } | Select-Object -First 1
                if (-not $release) { throw 'Could not resolve a supported Node.js 22 x64 installer.' }
                Install-RunnerPrerequisite "https://nodejs.org/dist/$($release.version)/node-$($release.version)-x64.msi" `
                    'node.msi' Node '' -Msi
            }
            $null = Install-RunnerCloudPrerequisites
            Write-ToolchainState @{
                state = 'installed'; tools = (Get-RunnerToolchainState); rebootRequired = $script:toolchainRebootRequired
                runnerRegistered = $false
            }
        }
        catch {
            Write-ToolchainState @{ state = 'failed'; error = $_.Exception.Message }
            throw
        }
        return
    }
    Status {
        if (-not (Test-Path -LiteralPath $statePath)) { throw 'No toolchain installation state exists.' }
        $state = Get-Content -LiteralPath $statePath -Raw | ConvertFrom-Json
        if ($state.sdk -and $state.sdk -ne $SdkVersion) { throw 'The installation belongs to another SDK request.' }
        if ($state.state -eq 'running') {
            $task = Get-ScheduledTask -TaskName $taskName
            if ($task.State -ne 'Running') {
                $info = Get-ScheduledTaskInfo -TaskName $taskName
                Write-ToolchainState @{ state = 'failed'; error = "Installation stopped without a result; task exit code=$($info.LastTaskResult)." }
            }
        }
    }
}
Write-Output ('EXCELMCP_TOOLCHAIN=' + (Get-Content -LiteralPath $statePath -Raw).Trim())
