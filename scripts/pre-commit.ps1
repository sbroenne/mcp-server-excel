#!/usr/bin/env pwsh
<#
.SYNOPSIS
    Git pre-commit hook to check for COM object leaks, Core Commands coverage, naming consistency, Success flag violations, CLI workflow, MCP Server functionality, and release deliverables

.DESCRIPTION
    Runs checks before allowing commits:
    0. Process cleanup - kills stale Excel, excelcli, and MCP server processes to prevent file locks
    1. COM leak checker - ensures no Excel COM objects are leaked
    2. Coverage and naming audit - ensures 100% Core Commands are exposed via MCP Server with aligned action names
    3. MCP-Core implementation audit - ensures every MCP action still has a Core implementation
    4. Success flag validation - ensures Success=true never paired with ErrorMessage (Rule 0)
    5. Release solution build - generates Release binaries and skill outputs used by downstream packaging
    6. CLI workflow smoke test - validates end-to-end CLI functionality
    7. MCP Server smoke test - validates all MCP tools work correctly
    8. CLI release packaging - validates NuGet + standalone ZIP artifacts
    9. MCP Server release packaging - validates NuGet + standalone ZIP artifacts
    10. VS Code extension packaging - validates the VSIX release packaging path
    11. MCPB bundle packaging - validates the Claude Desktop bundle artifact
    12. Agent skills packaging - validates the ZIP deliverable
    13. Plugin README validation - ensures overlays are complete and not stub content
    14. Dynamic cast audit - ensures ((dynamic)) casts are documented

    Ensures code quality and prevents regression.

.EXAMPLE
    .\pre-commit.ps1

.NOTES
    This script is called by the Git pre-commit hook.
    To install: Copy .git/hooks/pre-commit (bash) or configure Git to use this PowerShell version.
#>

$ErrorActionPreference = "Stop"
$rootDir = Split-Path -Parent $PSScriptRoot
$preCommitArtifactsDir = Join-Path $rootDir "artifacts\pre-commit"
$version = $null

function Invoke-ValidationStep {
    param(
        [string]$Heading,
        [scriptblock]$Action,
        [string]$FailureSummary,
        [string]$SuccessSummary
    )

    Write-Host ""
    Write-Host $Heading -ForegroundColor Cyan

    try {
        $output = & $Action 2>&1 | Out-String
        $exitCode = $LASTEXITCODE

        if ($exitCode -ne 0) {
            Write-Host ""
            Write-Host $FailureSummary -ForegroundColor Red
            if (-not [string]::IsNullOrWhiteSpace($output)) {
                Write-Host ""
                Write-Host $output -ForegroundColor Gray
            }
            exit 1
        }

        Write-Host $SuccessSummary -ForegroundColor Green
    }
    catch {
        Write-Host ""
        Write-Host "$FailureSummary $($_.Exception.Message)" -ForegroundColor Red
        exit 1
    }
}

function Reset-Directory {
    param([string]$Path)

    if (Test-Path $Path) {
        Remove-Item $Path -Recurse -Force
    }

    New-Item -ItemType Directory -Path $Path -Force | Out-Null
}

# CRITICAL: Check branch FIRST - never commit directly to main (Rule 6)
Write-Host "Checking current branch..." -ForegroundColor Cyan
$currentBranch = git branch --show-current

if ($currentBranch -eq "main") {
    Write-Host ""
    Write-Host "BLOCKED: Cannot commit directly to 'main' branch!" -ForegroundColor Red
    Write-Host ""
    Write-Host "   Rule 6: All Changes Via Pull Requests" -ForegroundColor Yellow
    Write-Host "   'Never commit to main. Create feature branch -> PR -> CI/CD + review -> merge.'" -ForegroundColor Yellow
    Write-Host ""
    Write-Host "   To fix:" -ForegroundColor Cyan
    Write-Host "   1. git stash                                    # Save your changes" -ForegroundColor White
    Write-Host "   2. git checkout -b feature/your-feature-name    # Create feature branch" -ForegroundColor White
    Write-Host "   3. git stash pop                                # Restore changes" -ForegroundColor White
    Write-Host "   4. git add <files>                              # Stage changes" -ForegroundColor White
    Write-Host "   5. git commit -m 'your message'                 # Commit to feature branch" -ForegroundColor White
    Write-Host ""
    exit 1
}

Write-Host "Branch check passed - on '$currentBranch' (not main)" -ForegroundColor Green
Write-Host ""

# Kill stale Excel and MCP server processes to avoid file locks on Release binaries
Write-Host "Killing stale Excel and server processes..." -ForegroundColor Cyan

$killedProcesses = @()
foreach ($procName in @("EXCEL", "excelcli", "Sbroenne.ExcelMcp.McpServer", "Sbroenne.ExcelMcp.Service")) {
    $procs = Get-Process -Name $procName -ErrorAction SilentlyContinue
    if ($procs) {
        $procs | Stop-Process -Force -ErrorAction SilentlyContinue
        $killedProcesses += "$procName ($($procs.Count))"
    }
}

if ($killedProcesses.Count -gt 0) {
    Write-Host "   Killed: $($killedProcesses -join ', ')" -ForegroundColor Yellow
    # Brief pause to let file handles release
    Start-Sleep -Milliseconds 500
}
else {
    Write-Host "   No stale processes found" -ForegroundColor Gray
}

Write-Host "Process cleanup done" -ForegroundColor Green
Write-Host ""

Reset-Directory -Path $preCommitArtifactsDir

try {
    $propsPath = Join-Path $rootDir "Directory.Build.props"
    $propsXml = [xml](Get-Content $propsPath)
    $version = $propsXml.Project.PropertyGroup.Version | Where-Object { $_ } | Select-Object -First 1
}
catch {
    Write-Host "Warning: Could not read version from Directory.Build.props ($($_.Exception.Message))" -ForegroundColor Yellow
    $version = "local"
}

Write-Host "Checking for COM object leaks..." -ForegroundColor Cyan

try {
    $leakCheckScript = Join-Path $rootDir "scripts\check-com-leaks.ps1"
    & $leakCheckScript

    if ($LASTEXITCODE -ne 0) {
        Write-Host ""
        Write-Host "COM object leaks detected! Fix them before committing." -ForegroundColor Red
        exit 1
    }

    Write-Host "COM leak check passed" -ForegroundColor Green
}
catch {
    Write-Host "Error running COM leak check: $($_.Exception.Message)" -ForegroundColor Yellow
    Write-Host "   Continuing with coverage audit..." -ForegroundColor Gray
}

Write-Host ""
Write-Host "Checking Core Commands coverage and naming..." -ForegroundColor Cyan

try {
    $auditScript = Join-Path $rootDir "scripts\audit-core-coverage.ps1"
    & $auditScript -CheckNaming -FailOnGaps

    if ($LASTEXITCODE -ne 0) {
        Write-Host ""
        Write-Host "Coverage or naming issues detected!" -ForegroundColor Red
        Write-Host "   All Core methods must be exposed via MCP Server with matching names." -ForegroundColor Red
        Write-Host "   Fix the issues before committing (add/rename enum values and mappings)." -ForegroundColor Red
        exit 1
    }

    Write-Host "Coverage and naming checks passed - 100% coverage with consistent names" -ForegroundColor Green
}
catch {
    Write-Host ""
    Write-Host "Error running coverage audit: $($_.Exception.Message)" -ForegroundColor Red
    exit 1
}

Write-Host ""
Write-Host "Checking MCP actions have Core implementations..." -ForegroundColor Cyan

try {
    $mcpCoreScript = Join-Path $rootDir "scripts\check-mcp-core-implementations.ps1"
    & $mcpCoreScript

    if ($LASTEXITCODE -ne 0) {
        Write-Host ""
        Write-Host "MCP actions without Core implementations detected!" -ForegroundColor Red
        Write-Host "   All enum actions must have matching Core Command methods." -ForegroundColor Red
        Write-Host "   Fix the issues before committing (remove enum or implement method)." -ForegroundColor Red
        exit 1
    }

    Write-Host "MCP-Core implementation check passed" -ForegroundColor Green
}
catch {
    Write-Host ""
    Write-Host "Error running MCP-Core implementation check: $($_.Exception.Message)" -ForegroundColor Red
    exit 1
}

Write-Host ""
Write-Host "Checking Success flag violations (Rule 0)..." -ForegroundColor Cyan

try {
    $successFlagScript = Join-Path $rootDir "scripts\check-success-flag.ps1"
    & $successFlagScript

    if ($LASTEXITCODE -ne 0) {
        Write-Host ""
        Write-Host "Success flag violations detected!" -ForegroundColor Red
        Write-Host "   CRITICAL: Success=true with ErrorMessage confuses LLMs and causes data corruption." -ForegroundColor Red
        Write-Host "   Fix the violations before committing (add Success=false in catch blocks)." -ForegroundColor Red
        exit 1
    }

    Write-Host "Success flag check passed - all flags match reality" -ForegroundColor Green
}
catch {
    Write-Host ""
    Write-Host "Error running success flag check: $($_.Exception.Message)" -ForegroundColor Red
    exit 1
}

# NOTE: CLI coverage checks removed - commands are now auto-generated by Roslyn source generators
# The CLI generator produces all command classes and registration from Core interfaces
# Validation is handled by:
# - Build-time generator errors if interfaces are malformed
# - CLI workflow smoke test below (end-to-end validation)

Invoke-ValidationStep `
    -Heading "Building Release solution..." `
    -FailureSummary "Release solution build failed!" `
    -SuccessSummary "Release solution build passed - Release binaries and generated skill docs are up to date" `
    -Action {
        Push-Location $rootDir
        try {
            dotnet build Sbroenne.ExcelMcp.sln --configuration Release -p:NuGetAudit=false --verbosity minimal
        }
        finally {
            Pop-Location
        }
    }

Write-Host ""
Write-Host "Auto-staging generated SKILL.md files..." -ForegroundColor Cyan

try {
    # SKILL.md + references are generated during the Release solution build above.
    # Auto-stage all of them so developers never have to think about it.
    $skillPaths = @(
        "skills/excel-mcp/SKILL.md",
        "skills/excel-cli/SKILL.md",
        "skills/excel-mcp/references/",
        "skills/excel-cli/references/"
    )
    $skillDiff = git diff --name-only -- @skillPaths 2>&1
    $untrackedSkills = git ls-files --others --exclude-standard -- @skillPaths 2>&1

    $allChanges = @()
    if ($skillDiff) { $allChanges += $skillDiff }
    if ($untrackedSkills) { $allChanges += $untrackedSkills }

    if ($allChanges.Count -gt 0) {
        git add -- @skillPaths
        Write-Host "Skill files were regenerated and auto-staged ($($allChanges.Count) files)" -ForegroundColor Green
        $allChanges | ForEach-Object { Write-Host "   + $_" -ForegroundColor DarkGray }
    } else {
        Write-Host "Skill files are already up to date" -ForegroundColor Green
    }
}
catch {
    Write-Host "Error auto-staging SKILL.md files: $($_.Exception.Message)" -ForegroundColor Yellow
    Write-Host "   Continuing with remaining checks..." -ForegroundColor Gray
}

Invoke-ValidationStep `
    -Heading "Running CLI workflow smoke test..." `
    -FailureSummary "CLI workflow smoke test failed! This blocks the release deliverable gates because the CLI artifact itself is not healthy." `
    -SuccessSummary "CLI workflow smoke test passed" `
    -Action {
        $cliWorkflowScript = Join-Path $rootDir "scripts\Test-CliWorkflow.ps1"
        & $cliWorkflowScript
    }

Write-Host ""
Write-Host "Running MCP Server smoke test..." -ForegroundColor Cyan

# Stop ExcelMCP Service before smoke test to prevent DLL locking
& "$PSScriptRoot\Stop-ExcelMcpProcesses.ps1"

try {
    # Run the smoke test - validates all MCP tools work correctly
    $smokeTestFilter = "FullyQualifiedName~McpServerSmokeTests.SmokeTest_AllTools_E2EWorkflow"

    Write-Host "   dotnet test --filter `"$smokeTestFilter`"" -ForegroundColor Gray

    # Capture output to verify tests actually ran (dotnet test returns 0 even if no tests match!)
    $testOutput = dotnet test --filter $smokeTestFilter --verbosity minimal 2>&1 | Out-String
    $testExitCode = $LASTEXITCODE

    # Check if any tests actually passed (critical - filter typos cause silent failures!)
    # Note: "No test matches" appears for projects without the test, so we check for "Passed"
    if (-not ($testOutput -match "Passed!.*Passed:\s*[1-9]")) {
        Write-Host ""
        Write-Host "CRITICAL: No smoke tests passed! Filter may have matched zero tests." -ForegroundColor Red
        Write-Host "   Filter: $smokeTestFilter" -ForegroundColor Yellow
        Write-Host "   This likely means the test was renamed or deleted." -ForegroundColor Yellow
        Write-Host "   Verify the test exists: McpServerSmokeTests.SmokeTest_AllTools_E2EWorkflow" -ForegroundColor Yellow
        Write-Host ""
        Write-Host $testOutput -ForegroundColor Gray
        exit 1
    }

    if ($testExitCode -ne 0) {
        Write-Host ""
        Write-Host "MCP Server smoke test failed! Core functionality is broken." -ForegroundColor Red
        Write-Host "   This test validates all MCP tools work correctly." -ForegroundColor Red
        Write-Host "   Fix the issues before committing." -ForegroundColor Red
        Write-Host ""
        Write-Host $testOutput -ForegroundColor Gray
        exit 1
    }

    Write-Host "MCP Server smoke test passed - all tools functional" -ForegroundColor Green
}
catch {
    Write-Host ""
    Write-Host "Error running smoke test: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "   Ensure Excel is installed and accessible." -ForegroundColor Yellow
    exit 1
}

Invoke-ValidationStep `
    -Heading "Building CLI release deliverables..." `
    -FailureSummary "CLI release deliverable validation failed!" `
    -SuccessSummary "CLI release deliverables passed - NuGet package and standalone ZIP were built locally" `
    -Action {
        $cliNupkgDir = Join-Path $preCommitArtifactsDir "cli-nupkg"
        $cliPublishDir = Join-Path $preCommitArtifactsDir "cli-publish"
        $cliReleaseDir = Join-Path $preCommitArtifactsDir "cli-release"
        $cliZipPath = Join-Path $preCommitArtifactsDir "ExcelMcp-CLI-$version-windows.zip"

        Reset-Directory -Path $cliNupkgDir
        Reset-Directory -Path $cliPublishDir
        Reset-Directory -Path $cliReleaseDir

        Push-Location $rootDir
        try {
            dotnet pack src\ExcelMcp.CLI\ExcelMcp.CLI.csproj --configuration Release --no-build --output $cliNupkgDir -p:Version=$version -p:NuGetAudit=false
            dotnet publish src\ExcelMcp.CLI\ExcelMcp.CLI.csproj --configuration Release --runtime win-x64 --self-contained true -p:PublishSingleFile=true -p:IncludeNativeLibrariesForSelfExtract=true -p:PublishTrimmed=false -p:PublishReadyToRun=false -p:Version=$version -p:NuGetAudit=false --output $cliPublishDir

            Copy-Item (Join-Path $cliPublishDir "excelcli.exe") $cliReleaseDir
            Copy-Item "README.md" $cliReleaseDir
            Copy-Item "LICENSE" $cliReleaseDir

            if (Test-Path $cliZipPath) {
                Remove-Item $cliZipPath -Force
            }

            Compress-Archive -Path (Join-Path $cliReleaseDir "*") -DestinationPath $cliZipPath

            if (-not (Get-ChildItem $cliNupkgDir -Filter "*.nupkg" -ErrorAction Stop)) {
                throw "CLI NuGet package was not created."
            }

            if (-not (Test-Path $cliZipPath)) {
                throw "CLI ZIP artifact was not created."
            }
        }
        finally {
            Pop-Location
        }
    }

Invoke-ValidationStep `
    -Heading "Building MCP Server release deliverables..." `
    -FailureSummary "MCP Server release deliverable validation failed!" `
    -SuccessSummary "MCP Server release deliverables passed - NuGet package and standalone ZIP were built locally" `
    -Action {
        $mcpNupkgDir = Join-Path $preCommitArtifactsDir "mcp-server-nupkg"
        $mcpPublishDir = Join-Path $preCommitArtifactsDir "mcp-server-publish"
        $mcpReleaseDir = Join-Path $preCommitArtifactsDir "mcp-server-release"
        $mcpZipPath = Join-Path $preCommitArtifactsDir "ExcelMcp-MCP-Server-$version-windows.zip"

        Reset-Directory -Path $mcpNupkgDir
        Reset-Directory -Path $mcpPublishDir
        Reset-Directory -Path $mcpReleaseDir

        Push-Location $rootDir
        try {
            dotnet pack src\ExcelMcp.McpServer\ExcelMcp.McpServer.csproj --configuration Release --no-build --output $mcpNupkgDir -p:Version=$version -p:NuGetAudit=false
            dotnet publish src\ExcelMcp.McpServer\ExcelMcp.McpServer.csproj --configuration Release --runtime win-x64 --self-contained true -p:PublishSingleFile=true -p:IncludeNativeLibrariesForSelfExtract=true -p:PublishTrimmed=false -p:PublishReadyToRun=false -p:Version=$version -p:NuGetAudit=false --output $mcpPublishDir

            $publishedExe = Join-Path $mcpPublishDir "Sbroenne.ExcelMcp.McpServer.exe"
            $renamedExe = Join-Path $mcpPublishDir "mcp-excel.exe"
            if (-not (Test-Path $publishedExe)) {
                throw "Published MCP Server executable was not created."
            }

            if (Test-Path $renamedExe) {
                Remove-Item $renamedExe -Force
            }

            Rename-Item $publishedExe "mcp-excel.exe"

            Copy-Item (Join-Path $mcpPublishDir "mcp-excel.exe") $mcpReleaseDir
            Copy-Item "README.md" $mcpReleaseDir
            Copy-Item "LICENSE" $mcpReleaseDir

            if (Test-Path $mcpZipPath) {
                Remove-Item $mcpZipPath -Force
            }

            Compress-Archive -Path (Join-Path $mcpReleaseDir "*") -DestinationPath $mcpZipPath

            if (-not (Get-ChildItem $mcpNupkgDir -Filter "*.nupkg" -ErrorAction Stop)) {
                throw "MCP Server NuGet package was not created."
            }

            if (-not (Test-Path $mcpZipPath)) {
                throw "MCP Server ZIP artifact was not created."
            }
        }
        finally {
            Pop-Location
        }
    }

Invoke-ValidationStep `
    -Heading "Running VS Code extension package validation..." `
    -FailureSummary "VS Code extension package validation failed! Fix the extension build or manifest mismatch before committing." `
    -SuccessSummary "VS Code extension package validation passed" `
    -Action {
        $extensionDir = Join-Path $rootDir "vscode-extension"
        Push-Location $extensionDir
        try {
            npm run package
        }
        finally {
            Pop-Location
        }
    }

Invoke-ValidationStep `
    -Heading "Building MCPB bundle deliverable..." `
    -FailureSummary "MCPB bundle validation failed!" `
    -SuccessSummary "MCPB bundle validation passed - Claude Desktop bundle was built locally" `
    -Action {
        $mcpbOutputRelative = "..\artifacts\pre-commit\mcpb"
        $mcpbOutputDir = Join-Path $preCommitArtifactsDir "mcpb"
        Reset-Directory -Path $mcpbOutputDir

        $mcpbDir = Join-Path $rootDir "mcpb"
        Push-Location $mcpbDir
        try {
            .\Build-McpBundle.ps1 -Version $version -OutputDir $mcpbOutputRelative

            if (-not (Get-ChildItem $mcpbOutputDir -Filter "*.mcpb" -ErrorAction Stop)) {
                throw "MCPB artifact was not created."
            }
        }
        finally {
            Pop-Location
        }
    }

Invoke-ValidationStep `
    -Heading "Building agent skills deliverables..." `
    -FailureSummary "Agent skills deliverable validation failed!" `
    -SuccessSummary "Agent skills deliverables passed - ZIP package was built locally" `
    -Action {
        $skillsOutputDir = Join-Path $preCommitArtifactsDir "skills"
        Reset-Directory -Path $skillsOutputDir

        Push-Location $rootDir
        try {
            .\scripts\Build-AgentSkills.ps1 -OutputDir "artifacts/pre-commit/skills" -Version $version

            if (-not (Get-ChildItem $skillsOutputDir -Filter "excel-skills-v*.zip" -ErrorAction Stop)) {
                throw "Agent skills ZIP artifact was not created."
            }
        }
        finally {
            Pop-Location
        }
    }

Write-Host ""
Write-Host "Validating plugin README overlays..." -ForegroundColor Cyan

try {
    $pluginReadmeScript = Join-Path $rootDir "scripts\check-plugin-readmes.ps1"
    & $pluginReadmeScript

    if ($LASTEXITCODE -ne 0) {
        Write-Host ""
        Write-Host "Plugin README validation failed!" -ForegroundColor Red
        Write-Host "   Thin/stub README overlays would overwrite richer published templates." -ForegroundColor Red
        Write-Host "   Enrich the overlay or remove it to use the published template." -ForegroundColor Red
        exit 1
    }

    Write-Host "Plugin README validation passed - overlays are complete" -ForegroundColor Green
}
catch {
    Write-Host "Error running plugin README check: $($_.Exception.Message)" -ForegroundColor Yellow
    Write-Host "   Continuing..." -ForegroundColor Gray
}

Write-Host ""
Write-Host "Checking for undocumented ((dynamic)) casts..." -ForegroundColor Cyan

try {
    $dynamicCastScript = Join-Path $rootDir "scripts\check-dynamic-casts.ps1"
    & $dynamicCastScript

    if ($LASTEXITCODE -ne 0) {
        Write-Host ""
        Write-Host "Undocumented ((dynamic)) casts detected!" -ForegroundColor Red
        Write-Host "   Add a justification comment (// PIA gap:, // TODO:, or // Reason:) before each cast." -ForegroundColor Red
        Write-Host "   See docs/PIA-COVERAGE.md for guidance." -ForegroundColor Red
        exit 1
    }

    Write-Host "Dynamic cast check passed - all casts are documented" -ForegroundColor Green
}
catch {
    Write-Host "Error running dynamic cast check: $($_.Exception.Message)" -ForegroundColor Yellow
    Write-Host "   Continuing..." -ForegroundColor Gray
}

Write-Host ""
Write-Host "All pre-commit checks passed!" -ForegroundColor Green
exit 0
