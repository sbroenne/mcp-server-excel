<#
.SYNOPSIS
    Generates and packages the Excel MCP and CLI agent skills.

.DESCRIPTION
    GenerateOnly prepares CLI discovery and the two report-formatting skills.
    Packaging consumes that prepared output without regenerating it.
    CLI syntax and action lists are discovered through native --help, not copied
    into the skill package.

.EXAMPLE
    .\Build-AgentSkills.ps1 -GenerateOnly

.EXAMPLE
    .\Build-AgentSkills.ps1 -Version 1.2.0
#>
[CmdletBinding()]
param(
    [string]$OutputDir,
    [string]$Version,
    [switch]$GenerateOnly,
    [string]$SkillsDirectory
)

$ErrorActionPreference = "Stop"
$RepoRoot = Split-Path -Parent $PSScriptRoot
$SkillsDir = Join-Path $RepoRoot "skills"
$SharedDir = Join-Path $RepoRoot "docs\reference"
$SkillNames = @('excel-cli', 'excel-cli-report-formatting', 'excel-mcp-report-formatting')
. (Join-Path $PSScriptRoot 'PackageHelpers.ps1')

function Copy-SharedReferences {
    param(
        [string]$SkillPath,
        [ValidateSet('cli', 'mcp')][string]$Surface
    )

    $RefsDir = Join-Path $SkillPath "references"
    New-Item -ItemType Directory -Path $RefsDir -Force | Out-Null
    if (-not (Test-Path $SharedDir)) { throw "Required shared references not found: $SharedDir" }
    $FilesToCopy = @((Get-Item -LiteralPath (Join-Path $SharedDir 'report-formatting.md')))
    if ($FilesToCopy.Count -eq 0) { throw "No shared reference files found: $SharedDir" }
    foreach ($sourceFile in $FilesToCopy) {
        $destination = Join-Path $RefsDir $sourceFile.Name
        $surface = $Surface
        $sourceContent = (Get-Content -LiteralPath $sourceFile.FullName -Raw) -replace "`r`n?", "`n"
        $rendered = [regex]::Replace($sourceContent, '(?ms)^```(?<surface>cli|mcp)\n(?<body>.*?)^```[ \t]*(?:\n|$)', {
            param($match)
            if ($match.Groups['surface'].Value -ne $surface) { return '' }
            $language = if ($surface -eq 'cli') { 'powershell' } else { 'text' }
            return '```' + $language + "`n" + $match.Groups['body'].Value + '```' + "`n"
        })
        $rendered = [regex]::Replace($rendered, '\[([^\]]+)\]\(([a-z-]+)\.md(#[^)]*)?\)', {
            param($match)
            return '[' + $match.Groups[1].Value + '](https://excelmcpserver.dev/reference/' +
                $match.Groups[2].Value + '/' + $match.Groups[3].Value + ')'
        })
        Set-Content -LiteralPath $destination -Value $rendered -Encoding UTF8 -NoNewline
    }
    Write-Host "  Prepared formatting reference for $Surface" -ForegroundColor Green
}

if ($GenerateOnly) {
    if (-not $OutputDir) { $OutputDir = 'artifacts\generated-skills' }
    if (-not $Version) { $Version = (Get-Content (Join-Path $RepoRoot 'package.json') -Raw | ConvertFrom-Json).version }
}
elseif (-not $OutputDir) {
    $OutputDir = 'artifacts\skills'
}
if ([string]::IsNullOrWhiteSpace($Version)) {
    throw "Version is required. Pass -Version <version>."
}
$Version = $Version.Trim()
$OutputPath = [IO.Path]::GetFullPath($OutputDir, $RepoRoot)
Assert-PackageOutputPath -Path $OutputPath -RepoRoot $RepoRoot
if ($OutputPath -eq [IO.Path]::GetPathRoot($OutputPath) -or
    $OutputPath -eq $RepoRoot -or
    $OutputPath.StartsWith("$SkillsDir$([IO.Path]::DirectorySeparatorChar)", [StringComparison]::OrdinalIgnoreCase) -or
    $OutputPath -eq $SkillsDir) {
    throw "Skill output must not overlap source files: $OutputPath"
}
if (-not $SkillsDirectory) { $SkillsDirectory = Join-Path $RepoRoot 'artifacts\generated-skills' }
$StagingDir = Join-Path ([IO.Path]::GetTempPath()) "excel-skills-$([guid]::NewGuid().ToString('N'))"
New-Item -ItemType Directory -Path $StagingDir -Force | Out-Null
try {
    $SkillsStagingDir = Join-Path $StagingDir "skills"
    New-Item -ItemType Directory -Path $SkillsStagingDir -Force | Out-Null
    foreach ($name in $SkillNames) {
        $destination = Join-Path $SkillsStagingDir $name
        if ($GenerateOnly) {
            Copy-Item -LiteralPath (Join-Path $SkillsDir $name) -Destination $destination -Recurse
            if ($name.EndsWith('-report-formatting')) {
                $component = if ($name.StartsWith('excel-cli-')) { 'cli' } else { 'mcp' }
                Copy-SharedReferences -SkillPath $destination -Surface $component
            }
        }
        else {
            Copy-Item -LiteralPath (Join-Path $SkillsDirectory $name) -Destination $destination -Recurse
        }
        $requiredFiles = @('SKILL.md')
        if ($name.EndsWith('-report-formatting')) { $requiredFiles += 'references\report-formatting.md' }
        foreach ($required in $requiredFiles) {
            if (-not (Test-Path -LiteralPath (Join-Path $destination $required))) { throw "$name is missing $required." }
        }
        Set-Content -LiteralPath (Join-Path $destination 'VERSION') -Value $Version -Encoding utf8 -NoNewline
    }
    New-Item -ItemType Directory -Path $OutputPath -Force | Out-Null
    if ($GenerateOnly) {
        foreach ($name in $SkillNames) {
            $destination = Join-Path $OutputPath $name
            Install-PackageOutput -Source (Join-Path $SkillsStagingDir $name) -Destination $destination
        }
        foreach ($retired in @('excel-mcp')) {
            $legacy = Join-Path $OutputPath $retired
            if (Test-Path -LiteralPath $legacy) {
                Assert-PackageOutputPath -Path $legacy -RepoRoot $RepoRoot
                if (-not (Test-Path -LiteralPath (Join-Path $legacy 'SKILL.md')) -or
                    -not (Test-Path -LiteralPath (Join-Path $legacy 'VERSION'))) {
                    throw "Refusing to remove unrecognized retired skill output: $legacy"
                }
                Remove-Item -LiteralPath $legacy -Recurse -Force
            }
        }
        Write-Host "Generated complete skills at $OutputPath"
    }
    else {
        Copy-Item -LiteralPath (Join-Path $RepoRoot 'docs\AGENT-SKILLS.md') (Join-Path $StagingDir 'README.md')
        $zip = Join-Path $StagingDir "excel-skills-v$Version.zip"
        Compress-Archive -LiteralPath $SkillsStagingDir,(Join-Path $StagingDir 'README.md') -DestinationPath $zip
        Install-PackageOutput -Source $zip -Destination (Join-Path $OutputPath (Split-Path $zip -Leaf))
        Write-Host "Created $OutputPath\excel-skills-v$Version.zip"
    }
}
finally { Remove-Item -LiteralPath $StagingDir -Recurse -Force }
$global:LASTEXITCODE = 0
