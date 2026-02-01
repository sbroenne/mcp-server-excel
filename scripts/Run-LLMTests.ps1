<#
.SYNOPSIS
    Runs Excel LLM integration tests using agent-benchmark.

.DESCRIPTION
    Unified test runner for both MCP Server and CLI LLM tests.
    Uses agent-benchmark to verify AI agents can effectively use Excel tools.

    Configuration can be provided via command-line parameters or config files.
    The script looks for config files in the test directory in this order:
    1. llm-tests.config.local.json (git-ignored, for personal settings)
    2. llm-tests.config.json (shared defaults)

.PARAMETER Component
    Which component to test: "mcp" or "cli". Required.

.PARAMETER Scenario
    Optional. Run only a specific test scenario file.
    Example: excel-range-test.yaml

.PARAMETER Build
    If specified, builds the component before running tests.

.PARAMETER AgentBenchmarkPath
    Path to agent-benchmark. Can be absolute path to executable or Go project.
    If not specified, uses config file or downloads from GitHub.

.EXAMPLE
    .\Run-LLMTests.ps1 -Component mcp -Build
    Builds MCP Server and runs all MCP tests.

.EXAMPLE
    .\Run-LLMTests.ps1 -Component cli -Scenario excel-range-cli-test.yaml
    Runs only the CLI range scenario.

.EXAMPLE
    .\Run-LLMTests.ps1 -Component mcp -AgentBenchmarkPath "D:\source\agent-benchmark"
    Uses local agent-benchmark build for MCP tests.
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [ValidateSet("mcp", "cli")]
    [string]$Component,

    [string]$Scenario = "",
    [switch]$Build,
    [string]$AgentBenchmarkPath
)

$ErrorActionPreference = "Stop"

# Determine paths based on component
$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$RepoRoot = (Get-Item $ScriptDir).Parent.FullName

switch ($Component) {
    "mcp" {
        $SrcDir = Join-Path $RepoRoot "src\ExcelMcp.McpServer"
        $ProjectPath = Join-Path $SrcDir "ExcelMcp.McpServer.csproj"
        $TestDir = Join-Path $RepoRoot "tests\ExcelMcp.McpServer.LLM.Tests"
        $ComponentName = "Excel MCP Server"
        $SkillName = "excel-mcp"
    }
    "cli" {
        $SrcDir = Join-Path $RepoRoot "src\ExcelMcp.CLI"
        $ProjectPath = Join-Path $SrcDir "ExcelMcp.CLI.csproj"
        $TestDir = Join-Path $RepoRoot "tests\ExcelMcp.CLI.LLM.Tests"
        $ComponentName = "Excel CLI"
        $SkillName = "excel-cli"
    }
}

$ScenariosDir = Join-Path $TestDir "Scenarios"
$ReportsDir = Join-Path $TestDir "TestResults"

# Build command string (forward slashes for YAML compatibility)
$ProjectPathForYaml = $ProjectPath -replace '\\', '/'

# For both MCP Server and CLI, use pre-built exe for fast startup
# This avoids 60s+ cold start from dotnet run compilation
if ($Component -eq "cli") {
    $ExePath = Join-Path $SrcDir "bin\Release\net10.0-windows\excelcli.exe"
    if (-not $Build -and -not (Test-Path $ExePath)) {
        Write-Host "CLI exe not found. Building first..." -ForegroundColor Cyan
        $Build = $true
    }
    $ServerCommand = ($ExePath -replace '\\', '/')
} else {
    # MCP Server - use pre-built exe to avoid 60s+ dotnet run compilation delay
    $ExePath = Join-Path $SrcDir "bin\Release\net10.0\Sbroenne.ExcelMcp.McpServer.exe"
    if (-not $Build -and -not (Test-Path $ExePath)) {
        Write-Host "MCP Server exe not found. Building first..." -ForegroundColor Cyan
        $Build = $true
    }
    $ServerCommand = ($ExePath -replace '\\', '/')
}

# Load configuration from test directory
$ConfigLocalPath = Join-Path $TestDir "llm-tests.config.local.json"
$ConfigPath = Join-Path $TestDir "llm-tests.config.json"
$Config = $null

if (Test-Path $ConfigLocalPath) {
    Write-Host "Loading config from: $TestDir\llm-tests.config.local.json" -ForegroundColor DarkGray
    $Config = Get-Content $ConfigLocalPath -Raw | ConvertFrom-Json
}
elseif (Test-Path $ConfigPath) {
    Write-Host "Loading config from: $TestDir\llm-tests.config.json" -ForegroundColor DarkGray
    $Config = Get-Content $ConfigPath -Raw | ConvertFrom-Json
}

# Apply config defaults (command-line parameters override config)
if (-not $Build -and $Config -and $Config.build) {
    $Build = $Config.build
}

# Determine agent-benchmark path and mode
$ResolvedAgentBenchmarkPath = $null
$AgentBenchmarkMode = "executable"

if ($AgentBenchmarkPath) {
    if ([System.IO.Path]::IsPathRooted($AgentBenchmarkPath)) {
        $ResolvedAgentBenchmarkPath = $AgentBenchmarkPath
    }
    else {
        $JoinedPath = Join-Path $TestDir $AgentBenchmarkPath
        $ResolvedAgentBenchmarkPath = [System.IO.Path]::GetFullPath($JoinedPath)
    }
}
elseif ($Config -and $Config.agentBenchmarkPath) {
    $ConfigAgentPath = $Config.agentBenchmarkPath
    if ([System.IO.Path]::IsPathRooted($ConfigAgentPath)) {
        $ResolvedAgentBenchmarkPath = $ConfigAgentPath
    }
    else {
        $JoinedPath = Join-Path $TestDir $ConfigAgentPath
        $ResolvedAgentBenchmarkPath = [System.IO.Path]::GetFullPath($JoinedPath)
    }

    if ($Config.agentBenchmarkMode) {
        $AgentBenchmarkMode = $Config.agentBenchmarkMode
    }
}
else {
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

# Build component if requested
if ($Build) {
    Write-Host "Building $ComponentName..." -ForegroundColor Cyan
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

# Verify project/exe exists
if (-not (Test-Path $ProjectPath)) {
    Write-Error "$ComponentName project not found at: $ProjectPath"
    exit 1
}

# Verify exe exists after build (both CLI and MCP use pre-built exe now)
if (-not (Test-Path $ExePath)) {
    Write-Error "$ComponentName exe not found at: $ExePath. Run with -Build flag."
    exit 1
}

# Download agent-benchmark if needed (executable mode only)
if ($AgentBenchmarkMode -eq "executable") {
    if (-not (Test-Path $ResolvedAgentBenchmarkPath)) {
        $DefaultDownloadPath = Join-Path $TestDir "agent-benchmark.exe"
        if ($ResolvedAgentBenchmarkPath -eq $DefaultDownloadPath) {
            Write-Host "Downloading agent-benchmark (latest release)..." -ForegroundColor Cyan

            try {
                $ReleaseInfo = Invoke-RestMethod "https://api.github.com/repos/mykhaliev/agent-benchmark/releases/latest"
                $LatestVersion = $ReleaseInfo.tag_name
                Write-Host "Latest version: $LatestVersion" -ForegroundColor DarkGray

                $Asset = $ReleaseInfo.assets | Where-Object { $_.name -match "windows_amd64\.zip$" -and $_.name -notmatch "upx" } | Select-Object -First 1
                if (-not $Asset) {
                    throw "Could not find Windows amd64 zip asset in release $LatestVersion"
                }

                $ZipPath = Join-Path $TestDir "agent-benchmark.zip"
                Write-Host "Downloading: $($Asset.name)" -ForegroundColor DarkGray
                Invoke-WebRequest -Uri $Asset.browser_download_url -OutFile $ZipPath

                Write-Host "Extracting..." -ForegroundColor DarkGray
                Expand-Archive -Path $ZipPath -DestinationPath $TestDir -Force
                Remove-Item $ZipPath -Force

                if (Test-Path $ResolvedAgentBenchmarkPath) {
                    Write-Host "Downloaded agent-benchmark $LatestVersion" -ForegroundColor Green
                }
                else {
                    throw "agent-benchmark.exe not found after extraction"
                }
            }
            catch {
                Write-Warning "Could not download agent-benchmark: $_"
                Write-Host "Please download from: https://github.com/mykhaliev/agent-benchmark/releases"
                exit 1
            }
        }
        else {
            Write-Error "agent-benchmark not found at: $ResolvedAgentBenchmarkPath"
            exit 1
        }
    }
}
elseif ($AgentBenchmarkMode -eq "go-run") {
    if (-not (Test-Path $ResolvedAgentBenchmarkPath)) {
        Write-Error "agent-benchmark Go project not found at: $ResolvedAgentBenchmarkPath"
        exit 1
    }
    $GoModPath = Join-Path $ResolvedAgentBenchmarkPath "go.mod"
    if (-not (Test-Path $GoModPath)) {
        Write-Error "No go.mod found at: $ResolvedAgentBenchmarkPath"
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
Write-Host "$ComponentName - LLM Integration Tests" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "Command: $ServerCommand"
Write-Host "Agent-Benchmark: $ResolvedAgentBenchmarkPath ($AgentBenchmarkMode)"
Write-Host "Skill: $SkillName"
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
    switch ($Component) {
        "mcp" {
            $env:SERVER_COMMAND = $ServerCommand
        }
        "cli" {
            $env:CLI_COMMAND = $ServerCommand
        }
    }

    $env:TEST_RESULTS_PATH = $ReportsDir -replace '\\', '/'
    $env:TEMP_DIR = "C:/temp"
    $env:TEST_DIR = $ScenariosDir -replace '\\', '/'
    $env:SKILL_PATH_MCP = (Join-Path $RepoRoot "skills\excel-mcp") -replace '\\', '/'
    $env:SKILL_PATH_CLI = (Join-Path $RepoRoot "skills\excel-cli") -replace '\\', '/'

    $ReportFile = Join-Path $ReportsDir "$($ScenarioFile.BaseName)-report"

    $AgentArgs = @(
        "-f", $ScenarioFile.FullName,
        "-o", $ReportFile,
        "-reportType", "html,json",
        "-verbose"
    )

    if ($AgentBenchmarkMode -eq "go-run") {
        Write-Host "Command: go run . $($AgentArgs -join ' ')" -ForegroundColor DarkGray
        Push-Location $ResolvedAgentBenchmarkPath
        try {
            & go run . @AgentArgs
            $ExitCode = $LASTEXITCODE
        }
        finally {
            Pop-Location
        }
    }
    else {
        Write-Host "Command: agent-benchmark $($AgentArgs -join ' ')" -ForegroundColor DarkGray
        & $ResolvedAgentBenchmarkPath @AgentArgs
        $ExitCode = $LASTEXITCODE
    }

    if ($ExitCode -eq 0) {
        Write-Host "PASSED" -ForegroundColor Green
        $TotalPassed++
        $Results += [PSCustomObject]@{
            Scenario = $ScenarioFile.Name
            Status   = "PASSED"
            Report   = $ReportFile
        }
    }
    else {
        Write-Host "FAILED (exit code: $ExitCode)" -ForegroundColor Red
        $TotalFailed++
        $Results += [PSCustomObject]@{
            Scenario = $ScenarioFile.Name
            Status   = "FAILED"
            Report   = $ReportFile
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

if ($TotalFailed -gt 0) {
    exit 1
}
exit 0
