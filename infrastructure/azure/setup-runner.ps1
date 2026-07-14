# Installs the GitHub Actions runner on the Azure Excel VM.
# Run as Administrator after the VM is provisioned. The registration token is
# short-lived and is never written to the setup log.

param(
    [Parameter(Mandatory = $true)]
    [string]$GithubRepoUrl,

    [Parameter(Mandatory = $true)]
    [string]$GithubRunnerToken,

    [Parameter(Mandatory = $true)]
    [string]$WindowsServiceAccount,

    [Parameter(Mandatory = $true)]
    [string]$WindowsServicePassword,

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

try {
    Write-SetupLog "Starting GitHub Actions runner setup."

    $dotnetExe = Join-Path $env:ProgramFiles "dotnet\dotnet.exe"
    $installedSdk = if (Test-Path $dotnetExe) { & $dotnetExe --list-sdks 2>$null } else { @() }
    if (-not ($installedSdk -match "^10\.")) {
        Write-SetupLog "Installing .NET 10 SDK."
        $dotnetInstaller = Join-Path $env:TEMP "dotnet-sdk.exe"
        Invoke-WebRequest `
            -Uri "https://aka.ms/dotnet/10.0/dotnet-sdk-win-x64.exe" `
            -OutFile $dotnetInstaller `
            -UseBasicParsing
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

    New-Item -Path $runnerDir -ItemType Directory -Force | Out-Null
    Set-Location $runnerDir

    if (Test-Path (Join-Path $runnerDir ".runner")) {
        Write-SetupLog "Runner is already configured; restarting its service."
        $service = Get-Service -Name "actions.runner.*" -ErrorAction Stop
        Set-Service -Name $service.Name -StartupType Automatic
        Restart-Service -Name $service.Name -Force
        exit 0
    }

    Write-SetupLog "Resolving the latest GitHub Actions runner release."
    $release = Invoke-RestMethod `
        -Uri "https://api.github.com/repos/actions/runner/releases/latest" `
        -Headers @{ "User-Agent" = "ExcelMcp-Runner-Setup" }
    $runnerVersion = $release.tag_name.TrimStart("v")
    $runnerArchive = Join-Path $runnerDir "actions-runner.zip"
    $runnerUri = "https://github.com/actions/runner/releases/download/v$runnerVersion/actions-runner-win-x64-$runnerVersion.zip"

    Write-SetupLog "Installing GitHub Actions runner v$runnerVersion."
    Invoke-WebRequest -Uri $runnerUri -OutFile $runnerArchive -UseBasicParsing
    Expand-Archive -Path $runnerArchive -DestinationPath $runnerDir -Force
    Remove-Item $runnerArchive -Force

    & .\config.cmd `
        --url $GithubRepoUrl `
        --token $GithubRunnerToken `
        --name $RunnerName `
        --labels "excel" `
        --runnergroup "Default" `
        --work "_work" `
        --runasservice `
        --windowslogonaccount $WindowsServiceAccount `
        --windowslogonpassword $WindowsServicePassword `
        --unattended `
        --replace
    if ($LASTEXITCODE -ne 0) {
        throw "Runner configuration failed with exit code $LASTEXITCODE."
    }

    $service = Get-Service -Name "actions.runner.*" -ErrorAction Stop
    Set-Service -Name $service.Name -StartupType Automatic
    Write-SetupLog "Runner service $($service.Name) is $($service.Status)."
}
catch {
    Write-SetupLog "Setup failed: $($_.Exception.Message)"
    throw
}
