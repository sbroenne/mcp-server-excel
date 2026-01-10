<#
.SYNOPSIS
    Runs Excel MCP Server LLM integration tests using agent-benchmark.

.DESCRIPTION
    This script runs the agent-benchmark test suite against the Excel MCP Server.
    It builds the MCP server, downloads agent-benchmark if needed, and runs all test scenarios.

    Configuration can be provided via command-line parameters or a config file.
    The script looks for config files in this order:
    1. llm-tests.config.local.json (git-ignored, for personal settings)
    2. llm-tests.config.json (shared defaults)

    Command-line parameters override config file settings.

.PARAMETER Scenario
    Optional. Run only a specific test scenario file. If not specified, runs all scenarios.
    Example: excel-file-worksheet-test.yaml

.PARAMETER Build
    If specified, builds the MCP server before running tests.

.PARAMETER AgentBenchmarkPath
    Path to agent-benchmark executable. Can be absolute or relative to the test directory.
    If not specified, downloads from GitHub releases.

.PARAMETER Verbose
    If specified, shows detailed output during test execution.

.EXAMPLE
    .\Run-LLMTests.ps1 -Build
    Builds the MCP server and runs all tests.

.EXAMPLE
    .\Run-LLMTests.ps1 -Scenario excel-file-worksheet-test.yaml
    Runs only the file/worksheet scenario.

.EXAMPLE
    .\Run-LLMTests.ps1 -AgentBenchmarkPath "..\..\..\..\agent-benchmark\agent-benchmark.exe"
    Uses local agent-benchmark build with relative path.
#>

[CmdletBinding()]
param(
    [string]$Scenario = "",
    [switch]$Build,
    [string]$AgentBenchmarkPath
)

$ErrorActionPreference = "Stop"

# Paths
$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$RepoRoot = (Get-Item $ScriptDir).Parent.Parent.FullName
$SrcDir = Join-Path $RepoRoot "src\ExcelMcp.McpServer"
$ProjectPath = Join-Path $SrcDir "ExcelMcp.McpServer.csproj"
$TestDir = Join-Path $RepoRoot "tests\ExcelMcp.McpServer.LLM.Tests"
$ScenariosDir = Join-Path $TestDir "Scenarios"
$ReportsDir = Join-Path $TestDir "TestResults"

# The server command uses dotnet run with the project
# Use forward slashes for YAML compatibility (Windows handles both)
$ProjectPathForYaml = $ProjectPath -replace '\\', '/'
$ServerCommand = "dotnet run --project $ProjectPathForYaml -c Release --"

# Load configuration from file
$ConfigLocalPath = Join-Path $TestDir "llm-tests.config.local.json"
$ConfigPath = Join-Path $TestDir "llm-tests.config.json"
$Config = $null

if (Test-Path $ConfigLocalPath) {
    Write-Host "Loading config from: llm-tests.config.local.json" -ForegroundColor DarkGray
    $Config = Get-Content $ConfigLocalPath -Raw | ConvertFrom-Json
}
elseif (Test-Path $ConfigPath) {
    Write-Host "Loading config from: llm-tests.config.json" -ForegroundColor DarkGray
    $Config = Get-Content $ConfigPath -Raw | ConvertFrom-Json
}

# Apply config defaults (command-line parameters override config)
if (-not $Build -and $Config -and $Config.build) {
    $Build = $Config.build
}

# Determine agent-benchmark path and mode
$ResolvedAgentBenchmarkPath = $null
$AgentBenchmarkMode = "executable"  # Default mode

if ($AgentBenchmarkPath) {
    # Command-line parameter takes precedence
    if ([System.IO.Path]::IsPathRooted($AgentBenchmarkPath)) {
        $ResolvedAgentBenchmarkPath = $AgentBenchmarkPath
    }
    else {
        $JoinedPath = Join-Path $TestDir $AgentBenchmarkPath
        $ResolvedAgentBenchmarkPath = [System.IO.Path]::GetFullPath($JoinedPath)
    }
}
elseif ($Config -and $Config.agentBenchmarkPath) {
    # Config file path
    $ConfigAgentPath = $Config.agentBenchmarkPath
    if ([System.IO.Path]::IsPathRooted($ConfigAgentPath)) {
        $ResolvedAgentBenchmarkPath = $ConfigAgentPath
    }
    else {
        $JoinedPath = Join-Path $TestDir $ConfigAgentPath
        $ResolvedAgentBenchmarkPath = [System.IO.Path]::GetFullPath($JoinedPath)
    }

    # Check for mode from config
    if ($Config.agentBenchmarkMode) {
        $AgentBenchmarkMode = $Config.agentBenchmarkMode
    }
}
else {
    # Default: download location in test directory
    $ResolvedAgentBenchmarkPath = Join-Path $TestDir "agent-benchmark.exe"
}

# Ensure reports directory exists
if (-not (Test-Path $ReportsDir)) {
    New-Item -ItemType Directory -Path $ReportsDir | Out-Null
}

# Check for required environment variables
if (-not $env:AZURE_OPENAI_ENDPOINT) {
    Write-Error "AZURE_OPENAI_ENDPOINT environment variable is not set."
    exit 1
}
if (-not $env:AZURE_OPENAI_API_KEY) {
    Write-Host "Note: AZURE_OPENAI_API_KEY not set. Using Entra ID authentication." -ForegroundColor DarkGray
}

# Build MCP server if requested (optional, dotnet run will build anyway)
if ($Build) {
    Write-Host "Building Excel MCP Server..." -ForegroundColor Cyan
    Push-Location $SrcDir
    try {
        dotnet build -c Release
        if ($LASTEXITCODE -ne 0) {
            Write-Error "Build failed."
            exit 1
        }
    }
    finally {
        Pop-Location
    }
}

# Verify project exists
if (-not (Test-Path $ProjectPath)) {
    Write-Error "MCP server project not found at: $ProjectPath"
    exit 1
}

# Download agent-benchmark if not present (only for executable mode with default path)
if ($AgentBenchmarkMode -eq "executable") {
    if (-not (Test-Path $ResolvedAgentBenchmarkPath)) {
        # Only attempt download if using default location
        $DefaultDownloadPath = Join-Path $TestDir "agent-benchmark.exe"
        if ($ResolvedAgentBenchmarkPath -eq $DefaultDownloadPath) {
            Write-Host "Downloading agent-benchmark (latest release)..." -ForegroundColor Cyan

            try {
                # Get latest release info from GitHub API
                $ReleaseInfo = Invoke-RestMethod "https://api.github.com/repos/mykhaliev/agent-benchmark/releases/latest"
                $LatestVersion = $ReleaseInfo.tag_name
                Write-Host "Latest version: $LatestVersion" -ForegroundColor DarkGray

                # Find the Windows amd64 zip asset
                $Asset = $ReleaseInfo.assets | Where-Object { $_.name -match "windows_amd64\.zip$" -and $_.name -notmatch "upx" } | Select-Object -First 1
                if (-not $Asset) {
                    throw "Could not find Windows amd64 zip asset in release $LatestVersion"
                }

                $ZipPath = Join-Path $TestDir "agent-benchmark.zip"
                Write-Host "Downloading: $($Asset.name)" -ForegroundColor DarkGray
                Invoke-WebRequest -Uri $Asset.browser_download_url -OutFile $ZipPath

                # Extract the zip
                Write-Host "Extracting..." -ForegroundColor DarkGray
                Expand-Archive -Path $ZipPath -DestinationPath $TestDir -Force
                Remove-Item $ZipPath -Force

                # The zip contains agent-benchmark.exe directly
                if (Test-Path $ResolvedAgentBenchmarkPath) {
                    Write-Host "Downloaded agent-benchmark $LatestVersion to: $ResolvedAgentBenchmarkPath" -ForegroundColor Green
                }
                else {
                    throw "agent-benchmark.exe not found after extraction"
                }
            }
            catch {
                Write-Warning "Could not download agent-benchmark: $_"
                Write-Host "Please download manually from: https://github.com/mykhaliev/agent-benchmark/releases"
                Write-Host "Or clone and build: https://github.com/mykhaliev/agent-benchmark"
                exit 1
            }
        }
        else {
            Write-Error "agent-benchmark not found at: $ResolvedAgentBenchmarkPath`nPlease verify the path in your config file or command-line parameter."
            exit 1
        }
    }
}
elseif ($AgentBenchmarkMode -eq "go-run") {
    # Verify Go project directory exists
    if (-not (Test-Path $ResolvedAgentBenchmarkPath)) {
        Write-Error "agent-benchmark Go project not found at: $ResolvedAgentBenchmarkPath`nPlease verify the path in your config file."
        exit 1
    }
    # Verify it's a Go project (has go.mod or .go files)
    $GoModPath = Join-Path $ResolvedAgentBenchmarkPath "go.mod"
    if (-not (Test-Path $GoModPath)) {
        Write-Error "No go.mod found at: $ResolvedAgentBenchmarkPath`nThis doesn't appear to be a Go project."
        exit 1
    }
}

# Get scenario files
if ($Scenario) {
    $ScenarioFiles = Get-Item (Join-Path $ScenariosDir $Scenario)
}
else {
    $ScenarioFiles = Get-ChildItem -Path $ScenariosDir -Filter "*.yaml"
}

if ($ScenarioFiles.Count -eq 0) {
    Write-Error "No scenario files found in: $ScenariosDir"
    exit 1
}

Write-Host "`n========================================" -ForegroundColor Cyan
Write-Host "Excel MCP Server - LLM Integration Tests" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "Server: dotnet run --project $ProjectPath"
Write-Host "Agent-Benchmark: $ResolvedAgentBenchmarkPath ($AgentBenchmarkMode)"
Write-Host "Scenarios: $($ScenarioFiles.Count) file(s)"
Write-Host ""

# Run each scenario
$TotalPassed = 0
$TotalFailed = 0
$Results = @()

foreach ($ScenarioFile in $ScenarioFiles) {
    Write-Host "`nRunning: $($ScenarioFile.Name)" -ForegroundColor Cyan
    Write-Host ("-" * 50)

    # Set environment variables for template substitution
    # agent-benchmark automatically picks up all env vars as template variables
    $env:SERVER_COMMAND = $ServerCommand
    $env:TEST_RESULTS_PATH = $ReportsDir -replace '\\', '/'

    # Run agent-benchmark directly on the original file
    $ReportFile = Join-Path $ReportsDir "$($ScenarioFile.BaseName)-report"

    $Args = @(
        "-f", $ScenarioFile.FullName,
        "-o", $ReportFile,
        "-reportType", "html,json",
        "-verbose"
    )

    # Execute agent-benchmark based on mode
    if ($AgentBenchmarkMode -eq "go-run") {
        Write-Host "Command: go run . $($Args -join ' ')" -ForegroundColor DarkGray
        Push-Location $ResolvedAgentBenchmarkPath
        try {
            & go run . @Args
            $ExitCode = $LASTEXITCODE
        }
        finally {
            Pop-Location
        }
    }
    else {
        Write-Host "Command: agent-benchmark $($Args -join ' ')" -ForegroundColor DarkGray
        & $ResolvedAgentBenchmarkPath @Args
        $ExitCode = $LASTEXITCODE
    }

    if ($ExitCode -eq 0) {
        Write-Host "PASSED" -ForegroundColor Green
        $TotalPassed++
        $Results += [PSCustomObject]@{
            Scenario = $ScenarioFile.Name
            Status = "PASSED"
            Report = $ReportFile
        }
    }
    else {
        Write-Host "FAILED (exit code: $ExitCode)" -ForegroundColor Red
        $TotalFailed++
        $Results += [PSCustomObject]@{
            Scenario = $ScenarioFile.Name
            Status = "FAILED"
            Report = $ReportFile
        }
    }
}

# Summary
Write-Host "`n========================================" -ForegroundColor Cyan
Write-Host "Test Summary" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "Passed: $TotalPassed" -ForegroundColor Green
Write-Host "Failed: $TotalFailed" -ForegroundColor $(if ($TotalFailed -gt 0) { "Red" } else { "Green" })
Write-Host ""

$Results | Format-Table -AutoSize

Write-Host "`nReports saved to: $ReportsDir"

# Exit with appropriate code
if ($TotalFailed -gt 0) {
    exit 1
}
exit 0
