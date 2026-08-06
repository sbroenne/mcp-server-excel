#!/usr/bin/env pwsh
<#
.SYNOPSIS
    Validates that the tool/operation counts advertised in user-facing docs match
    the authoritative counts derived from the code (generated skill manifest + handwritten action enums).

.DESCRIPTION
    This guard exists to make the "count discrepancy" class of bug impossible to reintroduce.

    THE PROBLEM IT PREVENTS
    -----------------------
    The MCP server and the CLI expose DIFFERENT internal surfaces, and several docs used to
    hard-code counts from memory. That drifted (docs said 232, the generated SKILL.md said 229,
    FEATURES.md section headers summed to 231). This script computes the ONE canonical answer
    from code on every commit and fails if any doc disagrees.

    HOW THE CANONICAL NUMBERS ARE DERIVED
    -------------------------------------
    Authoritative source = the generated `_SkillManifest.g.cs` (produced by
    ServiceRegistryGenerator from the Core [ServiceCategory] interfaces). It reports the
    CLI/service surface: every Core command category, which INCLUDES the CLI-only `diag`
    self-test category but EXCLUDES the hand-written `file`/session tool (FileAction is not a
    Core [ServiceCategory]).

    The CLI user-facing surface is:

        CLI operations = manifest.TotalOperations
                       - diag operations        (CLI-only self-test, not user-facing)
                       + FileAction operations  (the file/session tool)

    The MCP surface adds the handwritten WorkflowAction tool on top of the CLI-equivalent
    domain surface. MCP and CLI counts are intentionally tracked separately.

    These MUST stay in lock-step with the ExcludeCommands/ExtraOperationCount/ExtraToolCount
    values passed to GenerateSkillFile in the CLI and MCP .csproj files, and with the ground
    truth (the actual [McpServerTool(Name=...)] surface). All of that is cross-checked below,
    so if anyone adds an action, adds/removes a tool, or renames diag/file/workflow, this fails until the
    docs are updated.

.NOTES
    Run after a Release build so the generated manifest and SKILL.md files are current.
    Exit code 0 = all counts consistent. Exit code 1 = a mismatch was found.
#>

$ErrorActionPreference = "Stop"
$rootDir = Split-Path -Parent $PSScriptRoot

$errors = [System.Collections.Generic.List[string]]::new()
function Add-Failure([string]$message) { $script:errors.Add($message) }

# ---------------------------------------------------------------------------
# 1. Parse the authoritative generated skill manifest
# ---------------------------------------------------------------------------
$manifestFile = Get-ChildItem -Path (Join-Path $rootDir "src\ExcelMcp.Core\obj") -Recurse -Filter "_SkillManifest.g.cs" -ErrorAction SilentlyContinue |
    Sort-Object { $_.FullName -notmatch "GeneratedFiles" } |
    Select-Object -First 1

if (-not $manifestFile) {
    Write-Host "ERROR: Could not find generated _SkillManifest.g.cs. Run a Release build first." -ForegroundColor Red
    exit 1
}

$manifestContent = Get-Content $manifestFile.FullName -Raw
$startMarker = 'public const string Json = @"'
$startIdx = $manifestContent.IndexOf($startMarker)
$endIdx = $manifestContent.LastIndexOf('";')
if ($startIdx -lt 0 -or $endIdx -le $startIdx) {
    Write-Host "ERROR: Could not extract JSON from $($manifestFile.FullName)" -ForegroundColor Red
    exit 1
}
$startIdx += $startMarker.Length
$json = $manifestContent.Substring($startIdx, $endIdx - $startIdx).Replace('""', '"')
$manifest = $json | ConvertFrom-Json

$manifestTools = [int]$manifest.TotalCommands
$manifestOps = [int]$manifest.TotalOperations

# ---------------------------------------------------------------------------
# 2. Compute the two adjustments from ground truth
# ---------------------------------------------------------------------------
# diag: the CLI-only self-test category (must NOT be counted in the user-facing surface).
$diagCommand = $manifest.Commands | Where-Object { $_.Name -eq 'diag' }
if (-not $diagCommand) {
    Add-Failure "Expected a 'diag' command in the manifest (used to compute the user-facing count). It is gone - update this script and the csproj ExcludeCommands."
    $diagOps = 0
} else {
    $diagOps = @($diagCommand.Actions).Count
}

# file/workflow: hand-written tool action enums absent from the generated manifest.
$toolActionsPath = Join-Path $rootDir "src\ExcelMcp.Core\Models\Actions\ToolActions.cs"
$toolActionsContent = Get-Content $toolActionsPath -Raw
function Get-ActionEnumCount([string]$enumName) {
    $enumMatch = [regex]::Match($toolActionsContent, "enum\s+$enumName\s*\{(?<body>[^}]*)\}")
    if (-not $enumMatch.Success) {
        Write-Host "ERROR: Could not locate the $enumName enum in ToolActions.cs" -ForegroundColor Red
        exit 1
    }

    $count = ([regex]::Matches($enumMatch.Groups['body'].Value, 'JsonStringEnumMemberName')).Count
    if ($count -eq 0) {
        Write-Host "ERROR: $enumName parsed to 0 operations - parsing bug." -ForegroundColor Red
        exit 1
    }

    return $count
}

$fileOps = Get-ActionEnumCount "FileAction"
$workflowOps = Get-ActionEnumCount "WorkflowAction"
# The CLI documentation groups related generated commands into 18 product-facing
# feature categories (for example, several range commands are documented as one
# Ranges category). That taxonomy is intentionally not the generator command count.
$canonicalCliTools = 18
$canonicalCliOps = $manifestOps - $diagOps + $fileOps
$canonicalMcpTools = $manifestTools - 1 + 1 + 1     # - diag + file + workflow
$canonicalMcpOps = $canonicalCliOps + $workflowOps

# ---------------------------------------------------------------------------
# 3. Cross-check against the REAL MCP tool surface ([McpServerTool(Name=...)])
# ---------------------------------------------------------------------------
$mcpToolNames = [System.Collections.Generic.HashSet[string]]::new()
$mcpSearchDirs = @(
    (Join-Path $rootDir "src\ExcelMcp.McpServer")
)
foreach ($dir in $mcpSearchDirs) {
    if (-not (Test-Path $dir)) { continue }
    Get-ChildItem -Path $dir -Recurse -Filter "*.cs" -ErrorAction SilentlyContinue | ForEach-Object {
        $c = Get-Content $_.FullName -Raw
        foreach ($m in [regex]::Matches($c, 'McpServerTool\s*\(\s*Name\s*=\s*"([^"]+)"')) {
            [void]$mcpToolNames.Add($m.Groups[1].Value)
        }
    }
}

# `layout` is a compact-profile-only facade and is intentionally excluded from
# the full-profile headline count. Keep the physical registration cross-check
# explicit so adding/removing a compact facade cannot silently drift the full docs.
$fullMcpToolNames = @($mcpToolNames | Where-Object { $_ -ne 'layout' })
if ($fullMcpToolNames.Count -ne $canonicalMcpTools) {
    Add-Failure ("Full MCP tool surface has {0} tools ([McpServerTool(Name=...)], excluding compact-only layout) but the code-derived MCP tool count is {1}." -f $fullMcpToolNames.Count, $canonicalMcpTools)
}
if (-not $mcpToolNames.Contains('layout')) {
    Add-Failure "No 'layout' MCP tool found - compact profile documentation expects the layout facade."
}
if ($mcpToolNames.Contains('diag')) {
    Add-Failure "A 'diag' MCP tool now exists - the user-facing count assumption (diag is CLI-only) is broken. Update this script and the csproj ExcludeCommands."
}
if (-not $mcpToolNames.Contains('file')) {
    Add-Failure "No 'file' MCP tool found - the user-facing count assumption (file adds $fileOps ops) is broken. Update this script and the csproj ExtraOperationCount."
}
if (-not $mcpToolNames.Contains('workflow')) {
    Add-Failure "No 'workflow' MCP tool found - the MCP count assumption (workflow adds $workflowOps ops) is broken."
}

# ---------------------------------------------------------------------------
# 4. Cross-check the csproj GenerateSkillFile parameters stay in sync
# ---------------------------------------------------------------------------
foreach ($projectExpectation in @(
    @{ Project = "src\ExcelMcp.McpServer\ExcelMcp.McpServer.csproj"; ExtraOps = $fileOps + $workflowOps; ExtraTools = 2 },
    @{ Project = "src\ExcelMcp.CLI\ExcelMcp.CLI.csproj"; ExtraOps = $fileOps; ExtraTools = 1 }
)) {
    $proj = $projectExpectation.Project
    $projPath = Join-Path $rootDir $proj
    if (-not (Test-Path $projPath)) { continue }
    $projContent = Get-Content $projPath -Raw
    $extraOpsMatch = [regex]::Match($projContent, 'ExtraOperationCount\s*=\s*"(\d+)"')
    if ($extraOpsMatch.Success -and [int]$extraOpsMatch.Groups[1].Value -ne $projectExpectation.ExtraOps) {
        Add-Failure ("$proj sets ExtraOperationCount={0} but its handwritten tools require {1}." -f $extraOpsMatch.Groups[1].Value, $projectExpectation.ExtraOps)
    }
    $extraToolsMatch = [regex]::Match($projContent, 'ExtraToolCount\s*=\s*"(\d+)"')
    if ($extraToolsMatch.Success -and [int]$extraToolsMatch.Groups[1].Value -ne $projectExpectation.ExtraTools) {
        Add-Failure ("$proj sets ExtraToolCount={0} but its handwritten tools require {1}." -f $extraToolsMatch.Groups[1].Value, $projectExpectation.ExtraTools)
    }
}

Write-Host "Canonical MCP (from code): $canonicalMcpTools tools, $canonicalMcpOps operations" -ForegroundColor Cyan
Write-Host "Canonical CLI (from code): $canonicalCliTools command groups, $canonicalCliOps operations" -ForegroundColor Cyan
Write-Host "  manifest: $manifestTools tools / $manifestOps ops; - diag($diagOps) + file($fileOps) + MCP workflow($workflowOps); full MCP tool surface: $($fullMcpToolNames.Count) tools; compact-only: layout" -ForegroundColor DarkGray

# ---------------------------------------------------------------------------
# 5. Validate headline claims across user-facing docs
# ---------------------------------------------------------------------------
# Each check: file + surface + regex. Capture group 't' (optional) must equal that
# surface's tool/group count; capture group 'o' (optional) must equal its operation
# count. A check that matches nothing fails
# (so a headline can't silently disappear or be reworded past the guard).
$checks = @(
    @{ Surface = "MCP"; File = "README.md";                              Pattern = '(?<t>\d+) tools with (?<o>\d+) operations' }
    @{ Surface = "MCP"; File = "README.md";                              Pattern = '(?<t>\d+) specialized tools with (?<o>\d+) operations' }
    @{ Surface = "MCP"; File = "README.md";                              Pattern = 'all (?<o>\d+) operations' }
    @{ Surface = "MCP"; File = "FEATURES.md";                            Pattern = '(?<t>\d+) specialized tools with (?<o>\d+) operations' }
    @{ Surface = "MCP"; File = "src\ExcelMcp.McpServer\README.md";       Pattern = '(?<t>\d+) specialized tools with (?<o>\d+) operations' }
    @{ Surface = "MCP"; File = "src\ExcelMcp.McpServer\README.md";       Pattern = 'all (?<o>\d+) operations' }
    @{ Surface = "CLI"; File = "src\ExcelMcp.CLI\README.md";             Pattern = '(?<t>\d+) command categories with (?<o>\d+) operations' }
    @{ Surface = "CLI"; File = "src\ExcelMcp.CLI\README.md";             Pattern = '\*\*(?<o>\d+) operations\*\* across' }
    @{ Surface = "MCP"; File = "vscode-extension\README.md";             Pattern = '(?<t>\d+) specialized tools with (?<o>\d+) operations' }
    @{ Surface = "MCP"; File = "vscode-extension\README.md";             Pattern = 'all (?<t>\d+) tools and (?<o>\d+) operations' }
    @{ Surface = "MCP"; File = "mcpb\README.md";                         Pattern = '(?<t>\d+) tools with (?<o>\d+) operations' }
    @{ Surface = "MCP"; File = "mcpb\manifest.json";                     Pattern = '(?<t>\d+) specialized tools with (?<o>\d+) operations' }
    @{ Surface = "MCP"; File = "gh-pages\docs\index.md";                 Pattern = '(?<t>\d+) tools and (?<o>\d+) operations' }
    @{ Surface = "MCP"; File = "gh-pages\docs\features.md";              Pattern = '(?<t>\d+) specialized tools with (?<o>\d+) operations' }
    @{ Surface = "MCP"; File = ".github\plugins\excel-mcp\README.md";    Pattern = '(?<t>\d+) specialized tools with (?<o>\d+) operations' }
    @{ Surface = "CLI"; File = ".github\plugins\excel-cli\README.md";    Pattern = '(?<t>\d+) command categories with (?<o>\d+) operations' }
    @{ Surface = "MCP"; File = "docs\COPILOT-PLUGIN-DISTRIBUTION.md";      Pattern = 'MCP Server with (?<t>\d+) tools \((?<o>\d+) operations\)' }
    @{ Surface = "MCP"; File = "docs\COPILOT-PLUGIN-DISTRIBUTION.md";      Pattern = 'full MCP Server with (?<t>\d+) tools \((?<o>\d+) operations\)' }
    @{ Surface = "MCP"; File = "skills\excel-mcp\SKILL.md";              Pattern = 'Provides (?<o>\d+) Excel operations' }
)

foreach ($check in $checks) {
    $path = Join-Path $rootDir $check.File
    if (-not (Test-Path $path)) {
        Add-Failure "Expected doc not found: $($check.File)"
        continue
    }
    $content = Get-Content $path -Raw
    if ($check.Surface -eq "MCP") {
        $expectedTools = $canonicalMcpTools
        $expectedOps = $canonicalMcpOps
    } else {
        $expectedTools = $canonicalCliTools
        $expectedOps = $canonicalCliOps
    }
    $matches = [regex]::Matches($content, $check.Pattern)
    if ($matches.Count -eq 0) {
        Add-Failure "$($check.File): expected headline pattern not found (was it reworded or removed?): /$($check.Pattern)/"
        continue
    }
    foreach ($m in $matches) {
        if ($m.Groups['t'].Success -and [int]$m.Groups['t'].Value -ne $expectedTools) {
            Add-Failure ("$($check.File) [$($check.Surface)]: tool/group count is {0} but should be {1} -> `"{2}`"" -f $m.Groups['t'].Value, $expectedTools, $m.Value.Trim())
        }
        if ($m.Groups['o'].Success -and [int]$m.Groups['o'].Value -ne $expectedOps) {
            Add-Failure ("$($check.File) [$($check.Surface)]: operation count is {0} but should be {1} -> `"{2}`"" -f $m.Groups['o'].Value, $expectedOps, $m.Value.Trim())
        }
    }
}

# ---------------------------------------------------------------------------
# 6. FEATURES.md describes the MCP surface, so its section sum must equal MCP.
# ---------------------------------------------------------------------------
$featuresPath = Join-Path $rootDir "FEATURES.md"
if (Test-Path $featuresPath) {
    $featuresContent = Get-Content $featuresPath -Raw
    $sectionSum = 0
    foreach ($m in [regex]::Matches($featuresContent, '(?m)^##\s+.*\((?<n>\d+) operations\)')) {
        $sectionSum += [int]$m.Groups['n'].Value
    }
    if ($sectionSum -ne $canonicalMcpOps) {
        Add-Failure ("FEATURES.md section headers sum to {0} operations but the MCP total is {1}. Fix the section header(s) that drifted." -f $sectionSum, $canonicalMcpOps)
    }
}

# ---------------------------------------------------------------------------
# Result
# ---------------------------------------------------------------------------
if ($errors.Count -gt 0) {
    Write-Host ""
    Write-Host "Documentation count validation FAILED ($($errors.Count) issue(s)):" -ForegroundColor Red
    foreach ($e in $errors) { Write-Host "  - $e" -ForegroundColor Red }
    Write-Host ""
    Write-Host "Canonical MCP: $canonicalMcpTools tools / $canonicalMcpOps operations; CLI: $canonicalCliTools groups / $canonicalCliOps operations." -ForegroundColor Yellow
    Write-Host "Update the docs above to match, or if the surface genuinely changed, update the counts everywhere." -ForegroundColor Yellow
    exit 1
}

Write-Host "Documentation count validation passed - MCP $canonicalMcpTools/$canonicalMcpOps; CLI $canonicalCliTools/$canonicalCliOps" -ForegroundColor Green
exit 0
