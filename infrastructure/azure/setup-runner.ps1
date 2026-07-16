# Installs the GitHub Actions runner on the Azure Excel VM.
# Run as Administrator after the VM is provisioned. Excel window automation
# requires an interactive desktop, so the runner starts after secure auto-logon
# instead of running in Windows service session 0.

param(
    [Parameter(Mandatory = $true)]
    [string]$GithubRepoUrl,

    [Parameter(Mandatory = $true)]
    [string]$GithubRunnerToken,

    [Parameter(Mandatory = $true)]
    [string]$WindowsAccount,

    [Parameter(Mandatory = $true)]
    [string]$WindowsPassword,

    [string]$RunnerName = "azure-excel-runner"
)

$ErrorActionPreference = "Stop"
$ProgressPreference = "SilentlyContinue"
$logPath = "C:\runner-setup.log"
$runnerDir = "C:\actions-runner"

function Write-SetupLog {
    param([string]$Message)

    $entry = "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] $Message"
    Write-Host $entry
    Add-Content -Path $logPath -Value $entry
}

function Assert-ValidAuthenticodeSignature {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    $signature = Get-AuthenticodeSignature -FilePath $Path
    if ($signature.Status -ne [System.Management.Automation.SignatureStatus]::Valid) {
        throw "Authenticode verification failed for '$Path': $($signature.StatusMessage)"
    }

    Write-SetupLog "Verified Authenticode signature for $(Split-Path $Path -Leaf): $($signature.SignerCertificate.Subject)"
}

function Assert-Sha256Digest {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path,

        [Parameter(Mandatory = $true)]
        [string]$ExpectedDigest
    )

    if (-not $ExpectedDigest.StartsWith("sha256:", [StringComparison]::OrdinalIgnoreCase)) {
        throw "Release metadata did not provide a SHA-256 digest for '$Path'."
    }

    $expectedHash = $ExpectedDigest.Substring("sha256:".Length)
    $actualHash = (Get-FileHash -Path $Path -Algorithm SHA256).Hash
    if (-not $actualHash.Equals($expectedHash, [StringComparison]::OrdinalIgnoreCase)) {
        throw "SHA-256 verification failed for '$Path'. Expected $expectedHash, got $actualHash."
    }

    Write-SetupLog "Verified SHA-256 digest for $(Split-Path $Path -Leaf)."
}

try {
    Write-SetupLog "Starting GitHub Actions runner setup."

    $accountParts = $WindowsAccount -split "\\", 2
    $accountQualifier = if ($accountParts.Count -eq 2) { $accountParts[0] } else { "." }
    $accountUser = if ($accountParts.Count -eq 2) { $accountParts[1] } else { $accountParts[0] }
    if ([string]::IsNullOrWhiteSpace($accountUser)) {
        throw "WindowsAccount must identify a local Windows user."
    }
    if ($accountQualifier -ne "." -and $accountQualifier -ine $env:COMPUTERNAME) {
        throw "WindowsAccount must be a local account (for example, '.\azureuser'). Domain accounts are not supported."
    }

    $localUser = Get-LocalUser -Name $accountUser -ErrorAction Stop
    $accountDomain = $env:COMPUTERNAME
    $qualifiedAccount = "$accountDomain\$accountUser"

    $dotnetExe = Join-Path $env:ProgramFiles "dotnet\dotnet.exe"
    $installedSdk = if (Test-Path $dotnetExe) { & $dotnetExe --list-sdks 2>$null } else { @() }
    if (-not ($installedSdk -match "^10\.")) {
        Write-SetupLog "Installing .NET 10 SDK."
        $dotnetInstaller = Join-Path $env:TEMP "dotnet-sdk.exe"
        Invoke-WebRequest `
            -Uri "https://aka.ms/dotnet/10.0/dotnet-sdk-win-x64.exe" `
            -OutFile $dotnetInstaller `
            -UseBasicParsing
        Assert-ValidAuthenticodeSignature -Path $dotnetInstaller
        Start-Process `
            -FilePath $dotnetInstaller `
            -ArgumentList "/quiet", "/norestart" `
            -Wait `
            -NoNewWindow
        Remove-Item $dotnetInstaller -Force
    }

    $env:Path = [Environment]::GetEnvironmentVariable("Path", "Machine") +
        ";" + [Environment]::GetEnvironmentVariable("Path", "User")

    if (-not (Get-Command git -ErrorAction SilentlyContinue)) {
        Write-SetupLog "Installing Git for Windows."
        $gitRelease = Invoke-RestMethod `
            -Uri "https://api.github.com/repos/git-for-windows/git/releases/latest" `
            -Headers @{ "User-Agent" = "ExcelMcp-Runner-Setup" }
        $gitInstallerAsset = $gitRelease.assets |
            Where-Object { $_.name -match "^Git-.*-64-bit\.exe$" } |
            Select-Object -First 1
        if (-not $gitInstallerAsset) {
            throw "Could not locate the Git for Windows 64-bit installer."
        }

        $gitInstaller = Join-Path $env:TEMP "git-for-windows.exe"
        Invoke-WebRequest `
            -Uri $gitInstallerAsset.browser_download_url `
            -OutFile $gitInstaller `
            -UseBasicParsing
        Assert-ValidAuthenticodeSignature -Path $gitInstaller
        Start-Process `
            -FilePath $gitInstaller `
            -ArgumentList "/VERYSILENT", "/NORESTART", "/NOCANCEL", "/SP-" `
            -Wait `
            -NoNewWindow
        Remove-Item $gitInstaller -Force
        $env:Path = [Environment]::GetEnvironmentVariable("Path", "Machine") +
            ";" + [Environment]::GetEnvironmentVariable("Path", "User")
    }

    $pwshExe = Join-Path $env:ProgramFiles "PowerShell\7\pwsh.exe"
    if (-not (Test-Path $pwshExe)) {
        Write-SetupLog "Installing PowerShell 7."
        $pwshRelease = Invoke-RestMethod `
            -Uri "https://api.github.com/repos/PowerShell/PowerShell/releases/latest" `
            -Headers @{ "User-Agent" = "ExcelMcp-Runner-Setup" }
        $pwshInstallerAsset = $pwshRelease.assets |
            Where-Object { $_.name -match "^PowerShell-.*-win-x64\.msi$" } |
            Select-Object -First 1
        if (-not $pwshInstallerAsset) {
            throw "Could not locate the PowerShell 7 x64 MSI."
        }

        $pwshInstaller = Join-Path $env:TEMP "powershell-7.msi"
        Invoke-WebRequest `
            -Uri $pwshInstallerAsset.browser_download_url `
            -OutFile $pwshInstaller `
            -UseBasicParsing
        Assert-ValidAuthenticodeSignature -Path $pwshInstaller
        Start-Process `
            -FilePath "msiexec.exe" `
            -ArgumentList "/i", "`"$pwshInstaller`"", "/qn", "/norestart", "ADD_PATH=1" `
            -Wait `
            -NoNewWindow
        Remove-Item $pwshInstaller -Force
        if (-not (Test-Path $pwshExe)) {
            throw "PowerShell 7 installation did not create $pwshExe."
        }
    }

    # Office desktop applications require these folders when launched from a
    # Windows service session. Without them, Excel starts but workbook open/save
    # calls fail with COM error 0x800A03EC.
    Write-SetupLog "Ensuring Office service-profile Desktop folders exist."
    @(
        "$env:WINDIR\System32\config\systemprofile\Desktop",
        "$env:WINDIR\SysWOW64\config\systemprofile\Desktop"
    ) | ForEach-Object {
        New-Item -Path $_ -ItemType Directory -Force | Out-Null
    }

    New-Item -Path $runnerDir -ItemType Directory -Force | Out-Null
    Set-Location $runnerDir

    if (-not (Test-Path (Join-Path $runnerDir ".runner"))) {
        Write-SetupLog "Resolving the latest GitHub Actions runner release."
        $release = Invoke-RestMethod `
            -Uri "https://api.github.com/repos/actions/runner/releases/latest" `
            -Headers @{ "User-Agent" = "ExcelMcp-Runner-Setup" }
        $runnerVersion = $release.tag_name.TrimStart("v")
        $runnerAssetName = "actions-runner-win-x64-$runnerVersion.zip"
        $runnerAsset = $release.assets |
            Where-Object { $_.name -eq $runnerAssetName } |
            Select-Object -First 1
        if (-not $runnerAsset) {
            throw "Could not locate the GitHub Actions runner asset '$runnerAssetName'."
        }

        $runnerArchive = Join-Path $runnerDir "actions-runner.zip"

        Write-SetupLog "Installing GitHub Actions runner v$runnerVersion."
        Invoke-WebRequest -Uri $runnerAsset.browser_download_url -OutFile $runnerArchive -UseBasicParsing
        Assert-Sha256Digest -Path $runnerArchive -ExpectedDigest $runnerAsset.digest
        Expand-Archive -Path $runnerArchive -DestinationPath $runnerDir -Force
        Remove-Item $runnerArchive -Force

        & .\config.cmd `
            --url $GithubRepoUrl `
            --token $GithubRunnerToken `
            --name $RunnerName `
            --labels "excel" `
            --runnergroup "Default" `
            --work "_work" `
            --unattended `
            --replace
        if ($LASTEXITCODE -ne 0) {
            throw "Runner configuration failed with exit code $LASTEXITCODE."
        }
    }
    else {
        Write-SetupLog "Runner is already registered."
    }

    Write-SetupLog "Configuring en-US locale for $qualifiedAccount."
    Set-WinSystemLocale -SystemLocale "en-US"
    $localeMarker = "C:\runner-locale-configured.txt"
    $localeTaskName = "ExcelMcp-Configure-Locale"
    $localeScript = @'
Set-Culture -CultureInfo "en-US"
Set-WinUserLanguageList -LanguageList "en-US" -Force
Set-WinHomeLocation -GeoId 244
Set-Content -Path "C:\runner-locale-configured.txt" -Value "configured"
'@
    $localeEncoded = [Convert]::ToBase64String([Text.Encoding]::Unicode.GetBytes($localeScript))
    Remove-Item $localeMarker -Force -ErrorAction SilentlyContinue
    Unregister-ScheduledTask -TaskName $localeTaskName -Confirm:$false -ErrorAction SilentlyContinue
    $localeAction = New-ScheduledTaskAction `
        -Execute "powershell.exe" `
        -Argument "-NoProfile -NonInteractive -EncodedCommand $localeEncoded"
    Register-ScheduledTask `
        -TaskName $localeTaskName `
        -Action $localeAction `
        -User $qualifiedAccount `
        -Password $WindowsPassword `
        -RunLevel Highest `
        -Force | Out-Null
    Start-ScheduledTask -TaskName $localeTaskName
    $localeDeadline = (Get-Date).AddMinutes(2)
    while (-not (Test-Path $localeMarker) -and (Get-Date) -lt $localeDeadline) {
        Start-Sleep -Seconds 2
    }
    Unregister-ScheduledTask -TaskName $localeTaskName -Confirm:$false
    if (-not (Test-Path $localeMarker)) {
        throw "Timed out configuring the runner user's locale."
    }

    Write-SetupLog "Configuring secure automatic logon for $qualifiedAccount."
    $toolsDir = "C:\ProgramData\ExcelMcp"
    $autologonExe = Join-Path $toolsDir "Autologon64.exe"
    if (-not (Test-Path $autologonExe)) {
        $autologonArchive = Join-Path $env:TEMP "Autologon.zip"
        $autologonExtract = Join-Path $env:TEMP "Autologon"
        Invoke-WebRequest `
            -Uri "https://download.sysinternals.com/files/AutoLogon.zip" `
            -OutFile $autologonArchive `
            -UseBasicParsing
        Remove-Item $autologonExtract -Recurse -Force -ErrorAction SilentlyContinue
        Expand-Archive -Path $autologonArchive -DestinationPath $autologonExtract -Force
        New-Item -Path $toolsDir -ItemType Directory -Force | Out-Null
        Copy-Item (Join-Path $autologonExtract "Autologon64.exe") $autologonExe -Force
        Remove-Item $autologonArchive -Force
        Remove-Item $autologonExtract -Recurse -Force
    }
    Assert-ValidAuthenticodeSignature -Path $autologonExe
    & $autologonExe $accountUser $accountDomain $WindowsPassword "/accepteula"
    if ($LASTEXITCODE -ne 0) {
        throw "Sysinternals Autologon configuration failed with exit code $LASTEXITCODE."
    }

    $localUser | Set-LocalUser -PasswordNeverExpires $true

    $runnerServices = @(Get-Service -Name "actions.runner.*" -ErrorAction SilentlyContinue)
    foreach ($runnerService in $runnerServices) {
        Write-SetupLog "Removing non-interactive runner service $($runnerService.Name)."
        if ($runnerService.Status -ne "Stopped") {
            Stop-Service -InputObject $runnerService -Force
        }
        & sc.exe delete $runnerService.Name | Out-Null
        if ($LASTEXITCODE -ne 0) {
            throw "Failed to remove runner service $($runnerService.Name)."
        }
    }
    Remove-Item (Join-Path $runnerDir ".service") -Force -ErrorAction SilentlyContinue

    Write-SetupLog "Registering the interactive runner startup task."
    $runnerTaskName = "ExcelMcp-GitHub-Runner"
    Unregister-ScheduledTask -TaskName $runnerTaskName -Confirm:$false -ErrorAction SilentlyContinue
    $runnerAction = New-ScheduledTaskAction `
        -Execute "cmd.exe" `
        -Argument "/c `"`"$runnerDir\run.cmd`"`"" `
        -WorkingDirectory $runnerDir
    $runnerTrigger = New-ScheduledTaskTrigger -AtLogOn -User $qualifiedAccount
    $runnerPrincipal = New-ScheduledTaskPrincipal `
        -UserId $qualifiedAccount `
        -LogonType Interactive `
        -RunLevel Highest
    $runnerSettings = New-ScheduledTaskSettingsSet `
        -StartWhenAvailable `
        -RestartCount 3 `
        -RestartInterval (New-TimeSpan -Minutes 1) `
        -ExecutionTimeLimit ([TimeSpan]::Zero)
    Register-ScheduledTask `
        -TaskName $runnerTaskName `
        -Action $runnerAction `
        -Trigger $runnerTrigger `
        -Principal $runnerPrincipal `
        -Settings $runnerSettings `
        -Force | Out-Null

    Write-SetupLog "Interactive runner configured. Reboot to activate auto-logon."
}
catch {
    Write-SetupLog "Setup failed: $($_.Exception.Message)"
    throw
}
