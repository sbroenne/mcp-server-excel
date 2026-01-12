# GitHub Actions Self-Hosted Runner Setup Script
# This script is executed by Azure VM CustomScriptExtension during VM provisioning
# Parameters are passed from Bicep template

param(
    [Parameter(Mandatory=$true)]
    [string]$GithubRepoUrl,

    [Parameter(Mandatory=$true)]
    [string]$GithubRunnerToken,

    [string]$RunnerName = "azure-excel-runner",
    [string]$RunnerVersion = "2.321.0"
)

# Error handling
$ErrorActionPreference = "Stop"

# Logging function
function Write-Log {
    param([string]$Message, [string]$Level = "INFO")
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logMessage = "[$timestamp] [$Level] $Message"
    Write-Host $logMessage
    Add-Content -Path "C:\runner-setup.log" -Value $logMessage
}

try {
    Write-Log "Starting GitHub Actions Runner setup..."
    Write-Log "Repository URL: $GithubRepoUrl"
    Write-Log "Runner Name: $RunnerName"
    Write-Log "Runner Version: $RunnerVersion"

    # Step 1: Install .NET 10 SDK
    Write-Log "Step 1: Installing .NET 10 SDK..."
    $dotnetInstallerPath = "$env:TEMP\dotnet-sdk.exe"

    Write-Log "Downloading .NET SDK installer..."
    Invoke-WebRequest -Uri "https://aka.ms/dotnet/10.0/dotnet-sdk-win-x64.exe" -OutFile $dotnetInstallerPath -UseBasicParsing

    Write-Log "Installing .NET SDK (silent)..."
    Start-Process -FilePath $dotnetInstallerPath -ArgumentList '/quiet', '/norestart' -Wait -NoNewWindow

    # Refresh PATH to include dotnet
    $env:Path = [System.Environment]::GetEnvironmentVariable("Path", "Machine") + ";" + [System.Environment]::GetEnvironmentVariable("Path", "User")

    Write-Log ".NET SDK installation complete"

    # Verify installation
    $dotnetVersion = & dotnet --version 2>&1
    if ($LASTEXITCODE -eq 0) {
        Write-Log ".NET SDK version: $dotnetVersion"
    } else {
        Write-Log "Warning: Could not verify .NET SDK installation" "WARN"
    }

    # Step 2: Create runner directory
    Write-Log "Step 2: Creating runner directory..."
    $runnerDir = "C:\actions-runner"
    if (-not (Test-Path $runnerDir)) {
        New-Item -Path $runnerDir -ItemType Directory -Force | Out-Null
        Write-Log "Created directory: $runnerDir"
    } else {
        Write-Log "Directory already exists: $runnerDir"
    }

    # Step 3: Download GitHub Actions Runner
    Write-Log "Step 3: Downloading GitHub Actions Runner v$RunnerVersion..."
    Set-Location $runnerDir

    $runnerZip = "actions-runner.zip"
    $runnerUrl = "https://github.com/actions/runner/releases/download/v$RunnerVersion/actions-runner-win-x64-$RunnerVersion.zip"

    Write-Log "Downloading from: $runnerUrl"
    Invoke-WebRequest -Uri $runnerUrl -OutFile $runnerZip -UseBasicParsing

    Write-Log "Extracting runner package..."
    Expand-Archive -Path $runnerZip -DestinationPath . -Force

    Write-Log "Runner package extracted successfully"

    # Step 4: Configure runner
    Write-Log "Step 4: Configuring GitHub Actions Runner..."

    $configArgs = @(
        "--url", $GithubRepoUrl,
        "--token", $GithubRunnerToken,
        "--name", $RunnerName,
        "--labels", "self-hosted,windows,excel",
        "--runnergroup", "Default",
        "--work", "_work",
        "--unattended"
    )

    Write-Log "Running config.cmd with arguments: $($configArgs -join ' ')"
    & .\config.cmd @configArgs

    if ($LASTEXITCODE -ne 0) {
        throw "Runner configuration failed with exit code: $LASTEXITCODE"
    }

    Write-Log "Runner configured successfully"

    # Step 5: Install as Windows Service
    Write-Log "Step 5: Installing runner as Windows service..."

    & .\svc.cmd install
    if ($LASTEXITCODE -ne 0) {
        throw "Service installation failed with exit code: $LASTEXITCODE"
    }

    Write-Log "Service installed successfully"

    # Step 6: Start the service
    Write-Log "Step 6: Starting runner service..."

    & .\svc.cmd start
    if ($LASTEXITCODE -ne 0) {
        throw "Service start failed with exit code: $LASTEXITCODE"
    }

    Write-Log "Service started successfully"

    # Step 7: Verify service status
    Write-Log "Step 7: Verifying service status..."
    Start-Sleep -Seconds 5

    $service = Get-Service -Name "actions.runner.*" -ErrorAction SilentlyContinue
    if ($service -and $service.Status -eq "Running") {
        Write-Log "✅ Runner service is running: $($service.Name)" "SUCCESS"
    } else {
        Write-Log "⚠️ Service status: $($service.Status)" "WARN"
    }

    Write-Log "✅ GitHub Actions Runner setup completed successfully!" "SUCCESS"
    Write-Log ""
    Write-Log "Next steps:"
    Write-Log "1. Verify runner appears in GitHub: $GithubRepoUrl/settings/actions/runners"
    Write-Log "2. RDP to VM and install Office 365 Excel"
    Write-Log "3. Activate Excel with your Office 365 account"
    Write-Log "4. Reboot VM for all changes to take effect"

    exit 0

} catch {
    Write-Log "❌ Error during setup: $_" "ERROR"
    Write-Log "Exception: $($_.Exception.Message)" "ERROR"
    Write-Log "Stack trace: $($_.ScriptStackTrace)" "ERROR"
    exit 1
}
