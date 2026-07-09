#!/usr/bin/env pwsh
<#
.SYNOPSIS
    Validates that the tool/operation counts advertised in user-facing docs match
    the authoritative counts derived from the code (generated skill manifest + FileAction enum).

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

    The user-facing surface (what README/FEATURES/SKILL.md advertise) is:

        canonical operations = manifest.TotalOperations
                               - diag operations        (CLI-only self-test, not user-facing)
                               + FileAction operations   (the file/session tool)

        canonical tools      = manifest.TotalCommands
                               - 1 (diag)
                               + 1 (file)

    These MUST stay in lock-step with the ExcludeCommands/ExtraOperationCount/ExtraToolCount
    values passed to GenerateSkillFile in the CLI and MCP .csproj files, and with the ground
    truth (the actual [McpServerTool(Name=...)] surface). All of that is cross-checked below,
    so if anyone adds an action, adds/removes a tool, or renames diag/file, this fails until the
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

# file: the hand-written session/file tool (FileAction enum), absent from the manifest.
$toolActionsPath = Join-Path $rootDir "src\ExcelMcp.Core\Models\Actions\ToolActions.cs"
$toolActionsContent = Get-Content $toolActionsPath -Raw
$fileEnumMatch = [regex]::Match($toolActionsContent, 'enum\s+FileAction\s*\{(?<body>[^}]*)\}')
if (-not $fileEnumMatch.Success) {
    Write-Host "ERROR: Could not locate the FileAction enum in ToolActions.cs" -ForegroundColor Red
    exit 1
}
$fileOps = ([regex]::Matches($fileEnumMatch.Groups['body'].Value, 'JsonStringEnumMemberName')).Count
if ($fileOps -eq 0) {
    Write-Host "ERROR: FileAction enum parsed to 0 operations - parsing bug." -ForegroundColor Red
    exit 1
}

$canonicalTools = $manifestTools - 1 + 1        # - diag + file
$canonicalOps = $manifestOps - $diagOps + $fileOps

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

if ($mcpToolNames.Count -ne $canonicalTools) {
    Add-Failure ("MCP tool surface has {0} tools ([McpServerTool(Name=...)]) but the manifest-derived canonical tool count is {1}. If you added/removed a tool, update the docs; if diag/file assumptions changed, update this script and the csproj GenerateSkillFile parameters." -f $mcpToolNames.Count, $canonicalTools)
}
if ($mcpToolNames.Contains('diag')) {
    Add-Failure "A 'diag' MCP tool now exists - the user-facing count assumption (diag is CLI-only) is broken. Update this script and the csproj ExcludeCommands."
}
if (-not $mcpToolNames.Contains('file')) {
    Add-Failure "No 'file' MCP tool found - the user-facing count assumption (file adds $fileOps ops) is broken. Update this script and the csproj ExtraOperationCount."
}

# ---------------------------------------------------------------------------
# 4. Cross-check the csproj GenerateSkillFile parameters stay in sync
# ---------------------------------------------------------------------------
foreach ($proj in @("src\ExcelMcp.McpServer\ExcelMcp.McpServer.csproj", "src\ExcelMcp.CLI\ExcelMcp.CLI.csproj")) {
    $projPath = Join-Path $rootDir $proj
    if (-not (Test-Path $projPath)) { continue }
    $projContent = Get-Content $projPath -Raw
    $extraOpsMatch = [regex]::Match($projContent, 'ExtraOperationCount\s*=\s*"(\d+)"')
    if ($extraOpsMatch.Success -and [int]$extraOpsMatch.Groups[1].Value -ne $fileOps) {
        Add-Failure ("$proj sets ExtraOperationCount={0} but FileAction has {1} operations. They must match so the generated SKILL.md count is correct." -f $extraOpsMatch.Groups[1].Value, $fileOps)
    }
}

Write-Host "Canonical (from code): $canonicalTools tools, $canonicalOps operations" -ForegroundColor Cyan
Write-Host "  manifest: $manifestTools tools / $manifestOps ops; - diag($diagOps) + file($fileOps); MCP tool surface: $($mcpToolNames.Count) tools" -ForegroundColor DarkGray

# ---------------------------------------------------------------------------
# 5. Validate headline claims across user-facing docs
# ---------------------------------------------------------------------------
# Each check: file + regex. Capture group 't' (optional) must equal canonicalTools,
# capture group 'o' (optional) must equal canonicalOps. A check that matches nothing fails
# (so a headline can't silently disappear or be reworded past the guard).
$checks = @(
    @{ File = "README.md";                              Pattern = '(?<t>\d+) tools with (?<o>\d+) operations' }
    @{ File = "README.md";                              Pattern = '(?<t>\d+) specialized tools with (?<o>\d+) operations' }
    @{ File = "README.md";                              Pattern = 'all (?<o>\d+) operations' }
    @{ File = "FEATURES.md";                            Pattern = '(?<t>\d+) specialized tools with (?<o>\d+) operations' }
    @{ File = "src\ExcelMcp.McpServer\README.md";       Pattern = '(?<t>\d+) specialized tools with (?<o>\d+) operations' }
    @{ File = "src\ExcelMcp.McpServer\README.md";       Pattern = 'all (?<o>\d+) operations' }
    @{ File = "src\ExcelMcp.CLI\README.md";             Pattern = 'with (?<o>\d+) operations matching' }
    @{ File = "src\ExcelMcp.CLI\README.md";             Pattern = '\*\*(?<o>\d+) operations\*\* across' }
    @{ File = "vscode-extension\README.md";             Pattern = '(?<t>\d+) specialized tools with (?<o>\d+) operations' }
    @{ File = "mcpb\README.md";                         Pattern = '(?<t>\d+) tools with (?<o>\d+) operations' }
    @{ File = "gh-pages\docs\index.md";                 Pattern = '(?<t>\d+) tools and (?<o>\d+) operations' }
    @{ File = ".github\plugins\excel-mcp\README.md";    Pattern = '(?<t>\d+) specialized tools with (?<o>\d+) operations' }
    @{ File = ".github\plugins\excel-cli\README.md";    Pattern = 'command categories with (?<o>\d+) operations' }
    @{ File = "skills\excel-mcp\SKILL.md";              Pattern = 'Provides (?<o>\d+) Excel operations' }
)

foreach ($check in $checks) {
    $path = Join-Path $rootDir $check.File
    if (-not (Test-Path $path)) {
        Add-Failure "Expected doc not found: $($check.File)"
        continue
    }
    $content = Get-Content $path -Raw
    $matches = [regex]::Matches($content, $check.Pattern)
    if ($matches.Count -eq 0) {
        Add-Failure "$($check.File): expected headline pattern not found (was it reworded or removed?): /$($check.Pattern)/"
        continue
    }
    foreach ($m in $matches) {
        if ($m.Groups['t'].Success -and [int]$m.Groups['t'].Value -ne $canonicalTools) {
            Add-Failure ("$($check.File): tool count is {0} but should be {1} -> `"{2}`"" -f $m.Groups['t'].Value, $canonicalTools, $m.Value.Trim())
        }
        if ($m.Groups['o'].Success -and [int]$m.Groups['o'].Value -ne $canonicalOps) {
            Add-Failure ("$($check.File): operation count is {0} but should be {1} -> `"{2}`"" -f $m.Groups['o'].Value, $canonicalOps, $m.Value.Trim())
        }
    }
}

# ---------------------------------------------------------------------------
# 6. FEATURES.md: the sum of per-section "(N operations)" headers must equal canonical
# ---------------------------------------------------------------------------
$featuresPath = Join-Path $rootDir "FEATURES.md"
if (Test-Path $featuresPath) {
    $featuresContent = Get-Content $featuresPath -Raw
    $sectionSum = 0
    foreach ($m in [regex]::Matches($featuresContent, '(?m)^##\s+.*\((?<n>\d+) operations\)')) {
        $sectionSum += [int]$m.Groups['n'].Value
    }
    if ($sectionSum -ne $canonicalOps) {
        Add-Failure ("FEATURES.md section headers sum to {0} operations but the canonical total is {1}. Fix the section header(s) that drifted." -f $sectionSum, $canonicalOps)
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
    Write-Host "Canonical counts are derived from code: $canonicalTools tools / $canonicalOps operations." -ForegroundColor Yellow
    Write-Host "Update the docs above to match, or if the surface genuinely changed, update the counts everywhere." -ForegroundColor Yellow
    exit 1
}

Write-Host "Documentation count validation passed - all docs report $canonicalTools tools / $canonicalOps operations" -ForegroundColor Green
exit 0
