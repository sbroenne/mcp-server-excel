#!/usr/bin/env pwsh
<#
.SYNOPSIS
    Updates or validates tool/operation counts from the authoritative code-derived counts.

.DESCRIPTION
    A workflow that runs on every push to `main` uses -Update to compute the canonical
    counts once and write them to the single generated include file `doc-counts.json`
    (repo root) and to every managed headline claim. Nothing else derives or restates
    these numbers: gh-pages, release notes, and every managed doc read `doc-counts.json`
    or the headline text it wrote. Development CI (pull requests) uses
    -AllowStaleAdvertisedCounts so feature-section totals and the count derivation
    remain guarded without requiring `main`'s already-current headlines to be
    re-updated on every feature branch.

    THE PROBLEM IT PREVENTS
    -----------------------
    The MCP server and the CLI expose DIFFERENT internal surfaces, and several docs used to
    hard-code counts from memory. That drifted (docs said 232, the generated SKILL.md said 229,
    the canonical feature references summed to 231). This script computes the ONE canonical answer
    from code. Release automation writes that answer into every managed claim.

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

    By default, the script refreshes the Release compiler-generated files used by this check so
    persisted files under obj cannot report stale counts. Use -SkipBuild only when the caller
    has already completed a Release solution build in the current working tree.

.PARAMETER SkipBuild
    Skip the Release solution build. Intended for callers that build the solution
    immediately before invoking this script.

.PARAMETER Update
    Replace managed advertised totals with the canonical code-derived values.

.PARAMETER AllowStaleAdvertisedCounts
    Do not fail when managed advertised totals differ from the canonical values.
    Structural checks, count derivation, and per-feature section totals still fail.

.NOTES
    Exit code 0 = all counts consistent. Exit code 1 = a mismatch was found.
#>

[CmdletBinding()]
param(
    [switch]$SkipBuild,
    [switch]$Update,
    [switch]$AllowStaleAdvertisedCounts
)

$ErrorActionPreference = "Stop"
$rootDir = Split-Path -Parent $PSScriptRoot

if ($Update -and $AllowStaleAdvertisedCounts) {
    throw "-Update and -AllowStaleAdvertisedCounts cannot be used together."
}

if (-not $SkipBuild) {
    Write-Host "Refreshing generated Release count metadata..." -ForegroundColor Cyan
    & dotnet build (Join-Path $rootDir "src\ExcelMcp.Core\ExcelMcp.Core.csproj") --configuration Release --no-restore -p:NuGetAudit=false --verbosity minimal
    if ($LASTEXITCODE -ne 0) {
        Write-Host "ERROR: Core Release build failed. Run dotnet restore, then retry this check." -ForegroundColor Red
        exit 1
    }

    & dotnet build (Join-Path $rootDir "src\ExcelMcp.McpServer\ExcelMcp.McpServer.csproj") --configuration Release --no-restore --no-dependencies -p:NuGetAudit=false --verbosity minimal
    if ($LASTEXITCODE -ne 0) {
        Write-Host "ERROR: MCP Server Release build failed. Run dotnet restore, then retry this check." -ForegroundColor Red
        exit 1
    }
}

$errors = [System.Collections.Generic.List[string]]::new()
function Add-Failure([string]$message) { $script:errors.Add($message) }

# ---------------------------------------------------------------------------
# 1. Parse the authoritative generated skill manifest
# ---------------------------------------------------------------------------
# Emitted sources can survive dotnet clean, so the default refresh above is required before
# trusting this file. Automated callers may skip it only after their own Release solution build.
$manifestPath = Join-Path $rootDir "src\ExcelMcp.Core\obj\GeneratedFiles\ExcelMcp.Generators\Sbroenne.ExcelMcp.Generators.ServiceRegistryGenerator\_SkillManifest.g.cs"

if (-not (Test-Path -LiteralPath $manifestPath -PathType Leaf)) {
    Write-Host "ERROR: Could not find generated _SkillManifest.g.cs. Run without -SkipBuild or complete a Release build first." -ForegroundColor Red
    exit 1
}

$manifestFile = Get-Item -LiteralPath $manifestPath
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
# 5. Update or validate headline claims across user-facing docs
# ---------------------------------------------------------------------------
# Each check: file + regex. Capture group 't' (optional) must equal canonicalTools,
# capture group 'o' (optional) must equal canonicalOps. A check that matches nothing fails
# (so a headline can't silently disappear or be reworded past the guard).
$checks = @(
    @{ File = "README.md";                              Pattern = '(?<t>\d+) tools with (?<o>\d+) operations' }
    @{ File = "README.md";                              Pattern = 'all (?<o>\d+) operations' }
    @{ File = "FEATURES.md";                            Pattern = '(?<t>\d+) specialized tools with (?<o>\d+) operations' }
    @{ File = "src\ExcelMcp.McpServer\README.md";       Pattern = '(?<t>\d+) specialized tools with (?<o>\d+) operations' }
    @{ File = "src\ExcelMcp.McpServer\README.md";       Pattern = 'all (?<o>\d+) operations' }
    @{ File = "src\ExcelMcp.CLI\README.md";             Pattern = 'provides (?<t>\d+) feature command categories with (?<o>\d+) operations matching' }
    @{ File = "src\ExcelMcp.CLI\README.md";             Pattern = 'without loading (?<t>\d+) tool schemas' }
    @{ File = "src\ExcelMcp.CLI\README.md";             Pattern = '\*\*(?<o>\d+) operations\*\* across' }
    @{ File = "vscode-extension\README.md";             Pattern = '(?<t>\d+) specialized tools with (?<o>\d+) operations' }
    @{ File = "vscode-extension\README.md";             Pattern = 'all (?<t>\d+) tools and (?<o>\d+) operations' }
    @{ File = "mcpb\README.md";                         Pattern = '(?<t>\d+) tools with (?<o>\d+) operations' }
    @{ File = "mcpb\manifest.json";                     Pattern = '(?<t>\d+) specialized tools with (?<o>\d+) operations' }
    @{ File = "mcpb\BUILD.md";                           Pattern = 'generates its (?<t>\d+) tool schemas' }
    @{ File = "src\ExcelMcp.CLI\ExcelMcp.CLI.csproj";   Pattern = '(?<o>\d+) operations across' }
    @{ File = "gh-pages\docs\index.md";                 Pattern = '(?<t>\d+) tools and (?<o>\d+) operations' }
    @{ File = ".github\plugins\excel-mcp\README.md";    Pattern = '(?<t>\d+) specialized tools with (?<o>\d+) operations' }
    @{ File = ".github\plugins\excel-cli\README.md";    Pattern = 'command categories with (?<o>\d+) operations' }
    @{ File = ".github\plugins\excel-cli\README.md";    Pattern = '\| (?<t>\d+) tool schemas loaded into context \|' }
    @{ File = "skills\excel-mcp\SKILL.md";              Pattern = 'Provides (?<o>\d+) Excel operations' }
    @{ File = "gh-pages\docs\faq.md";                   Pattern = 'same (?<o>\d+) operations' }
    @{ File = "docs\INSTALLATION-CLI.md";               Pattern = 'all (?<t>\d+) feature command categories' }
    @{ File = "docs\guides\EXCEL-COM-VS-FILE-PARSERS.md"; Pattern = '(?<o>\d+)\s+operations across (?<t>\d+) tools' }
    @{ File = "docs\COPILOT-PLUGIN-DISTRIBUTION.md";    Pattern = 'with (?<t>\d+) tools \((?<o>\d+) operations\)' }
)

# The website feature overview includes FEATURES.md; audit_site.py enforces
# that wrapper contract instead of requiring a second handwritten headline.
$documentContent = @{}
$updatedFiles = [System.Collections.Generic.HashSet[string]]::new()
foreach ($check in $checks) {
    $path = Join-Path $rootDir $check.File
    if (-not (Test-Path $path)) {
        Add-Failure "Expected doc not found: $($check.File)"
        continue
    }
    $content = if ($documentContent.ContainsKey($check.File)) {
        $documentContent[$check.File]
    } else {
        Get-Content $path -Raw
    }
    $matches = [regex]::Matches($content, $check.Pattern)
    if ($matches.Count -eq 0) {
        Add-Failure "$($check.File): expected headline pattern not found (was it reworded or removed?): /$($check.Pattern)/"
        continue
    }

    foreach ($m in @($matches) | Sort-Object Index -Descending) {
        $replacements = [System.Collections.Generic.List[object]]::new()
        if ($m.Groups['t'].Success -and [int]$m.Groups['t'].Value -ne $canonicalTools) {
            $replacements.Add([pscustomobject]@{
                Index = $m.Groups['t'].Index
                Length = $m.Groups['t'].Length
                Value = [string]$canonicalTools
                Kind = "tool"
                Previous = $m.Groups['t'].Value
            })
        }
        if ($m.Groups['o'].Success -and [int]$m.Groups['o'].Value -ne $canonicalOps) {
            $replacements.Add([pscustomobject]@{
                Index = $m.Groups['o'].Index
                Length = $m.Groups['o'].Length
                Value = [string]$canonicalOps
                Kind = "operation"
                Previous = $m.Groups['o'].Value
            })
        }

        foreach ($replacement in $replacements | Sort-Object Index -Descending) {
            if ($Update) {
                $content = $content.Remove($replacement.Index, $replacement.Length).Insert($replacement.Index, $replacement.Value)
                [void]$updatedFiles.Add($check.File)
            } elseif (-not $AllowStaleAdvertisedCounts) {
                Add-Failure ("$($check.File): {0} count is {1} but should be {2} -> `"{3}`"" -f $replacement.Kind, $replacement.Previous, $replacement.Value, $m.Value.Trim())
            }
        }
    }
    $documentContent[$check.File] = $content
}

if ($Update) {
    foreach ($file in $updatedFiles) {
        Set-Content -LiteralPath (Join-Path $rootDir $file) -Value $documentContent[$file] -NoNewline -Encoding utf8
        Write-Host "Updated $file" -ForegroundColor Green
    }
}

# ---------------------------------------------------------------------------
# 5a. Canonical generated include file - the ONE machine-readable place other
# automation (gh-pages, release notes, third-party tooling) reads counts from,
# instead of re-deriving them or parsing markdown headline text.
# ---------------------------------------------------------------------------
$docCountsFile = "doc-counts.json"
$docCountsPath = Join-Path $rootDir $docCountsFile
$docCountsJson = ([ordered]@{ tools = $canonicalTools; operations = $canonicalOps } | ConvertTo-Json) + "`n"

if ($Update) {
    $currentDocCounts = if (Test-Path -LiteralPath $docCountsPath) { Get-Content -LiteralPath $docCountsPath -Raw } else { $null }
    if ($currentDocCounts -ne $docCountsJson) {
        Set-Content -LiteralPath $docCountsPath -Value $docCountsJson -NoNewline -Encoding utf8
        Write-Host "Updated $docCountsFile" -ForegroundColor Green
    }
} elseif (-not (Test-Path -LiteralPath $docCountsPath)) {
    Add-Failure "$docCountsFile not found. It is the single generated include file every other count consumer reads; run this script with -Update to create it."
} else {
    $parsedDocCounts = $null
    try { $parsedDocCounts = Get-Content -LiteralPath $docCountsPath -Raw | ConvertFrom-Json } catch { $parsedDocCounts = $null }
    if (-not $parsedDocCounts -or $null -eq $parsedDocCounts.tools -or $null -eq $parsedDocCounts.operations) {
        Add-Failure "$docCountsFile is malformed - expected a JSON object with 'tools' and 'operations' properties."
    } elseif (([int]$parsedDocCounts.tools -ne $canonicalTools -or [int]$parsedDocCounts.operations -ne $canonicalOps) -and -not $AllowStaleAdvertisedCounts) {
        Add-Failure ("{0}: reports {1} tools / {2} operations but the canonical count is {3} tools / {4} operations." -f $docCountsFile, $parsedDocCounts.tools, $parsedDocCounts.operations, $canonicalTools, $canonicalOps)
    }
}

# ---------------------------------------------------------------------------
# 5b. The --help banner must stay derived, not restated
# ---------------------------------------------------------------------------
# The banner is the first thing a user sees when the server will not start, and it silently
# drifted to "22 tools with 195+ operations" while the real surface was 31/326 -- because nothing
# checked it. It is now interpolated from McpToolSurface, which reflects over the live
# [McpServerTool] registration. Rather than validating a literal here (there no longer is one),
# this fails if anyone reintroduces one.
$helpBannerFile = "src\ExcelMcp.McpServer\Program.cs"
$helpBannerPath = Join-Path $rootDir $helpBannerFile
if (-not (Test-Path $helpBannerPath)) {
    Add-Failure "Expected source file not found: $helpBannerFile"
} else {
    $helpBannerContent = Get-Content $helpBannerPath -Raw

    $hardCoded = [regex]::Matches($helpBannerContent, 'Provides\s+\d+\s+tools\s+with\s+\d+\+?\s+operations')
    foreach ($m in $hardCoded) {
        Add-Failure ("{0}: the --help banner restates the counts as a literal -> `"{1}`". Interpolate McpToolSurface.ToolCount / McpToolSurface.OperationCount instead so it cannot drift." -f $helpBannerFile, $m.Value.Trim())
    }

    if ($helpBannerContent -notmatch 'Provides \{McpToolSurface\.ToolCount\} tools with \{McpToolSurface\.OperationCount\} operations') {
        Add-Failure "${helpBannerFile}: the --help banner no longer derives its counts from McpToolSurface (was it reworded or removed?)."
    }
}

# ---------------------------------------------------------------------------
# 6. Distribution metadata must link to the documentation site
# ---------------------------------------------------------------------------
# Package/registry listing pages are the highest-authority inbound links we control.
# They must point at the canonical docs site, not straight at the GitHub repo.
$siteUrl = 'https://excelmcpserver.dev/'
$siteLinkChecks = @(
    @{ File = "Directory.Build.props";                  Pattern = '<PackageProjectUrl>https://excelmcpserver\.dev/</PackageProjectUrl>';  What = "PackageProjectUrl (NuGet package pages)" }
    @{ File = "src\ExcelMcp.McpServer\.mcp\server.json"; Pattern = '"websiteUrl"\s*:\s*"https://excelmcpserver\.dev/"';                    What = "websiteUrl (MCP registry listing)" }
    @{ File = "mcpb\manifest.json";                     Pattern = '"homepage"\s*:\s*"https://excelmcpserver\.dev/"';                      What = "homepage (Claude Desktop bundle)" }
    @{ File = "vscode-extension\package.json";          Pattern = '"homepage"\s*:\s*"https://excelmcpserver\.dev/"';                      What = "homepage (VS Code Marketplace)" }
)

foreach ($check in $siteLinkChecks) {
    $path = Join-Path $rootDir $check.File
    if (-not (Test-Path $path)) {
        Add-Failure "Expected metadata file not found: $($check.File)"
        continue
    }
    if ((Get-Content $path -Raw) -notmatch $check.Pattern) {
        Add-Failure "$($check.File): $($check.What) must point at $siteUrl - this is a primary inbound link and must not regress to the GitHub URL."
    }
}

# ---------------------------------------------------------------------------
# 7. Canonical feature docs: per-section "(N operations)" headers must sum to canonical
# ---------------------------------------------------------------------------
$featureFiles = Get-ChildItem (Join-Path $rootDir "docs\features") -Filter "*.md"
$sectionSum = 0
foreach ($featureFile in $featureFiles) {
    $featureContent = Get-Content $featureFile.FullName -Raw
    foreach ($m in [regex]::Matches($featureContent, '(?m)^##\s+.*\((?<n>\d+) operations\)')) {
        $sectionSum += [int]$m.Groups['n'].Value
    }
}
if ($sectionSum -ne $canonicalOps) {
    Add-Failure ("Canonical feature section headers sum to {0} operations but the canonical total is {1}. Fix the section header(s) that drifted." -f $sectionSum, $canonicalOps)
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
    Write-Host "Fix the structural mismatch above before the next -Update pass can proceed." -ForegroundColor Yellow
    exit 1
}

if ($Update) {
    Write-Host "Documentation counts generated - $canonicalTools tools / $canonicalOps operations ($($updatedFiles.Count) file(s) changed)" -ForegroundColor Green
} elseif ($AllowStaleAdvertisedCounts) {
    Write-Host "Documentation count structure passed - advertised totals may be behind the latest main (regenerated on every push to main)" -ForegroundColor Green
} else {
    Write-Host "Documentation count validation passed - all docs report $canonicalTools tools / $canonicalOps operations" -ForegroundColor Green
}
exit 0
