<#
.SYNOPSIS
    Generates and packages the Excel MCP and CLI agent skills.

.DESCRIPTION
    GenerateOnly renders templates and authored references after a Release build.
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
    [string]$SkillsDirectory,
    [string]$ManifestPath
)

$ErrorActionPreference = "Stop"
$RepoRoot = Split-Path -Parent $PSScriptRoot
$SkillsDir = Join-Path $RepoRoot "skills"
$SharedDir = Join-Path $SkillsDir "shared"
. (Join-Path $PSScriptRoot 'PackageHelpers.ps1')

function Copy-SharedReferences {
    param(
        [string]$SkillPath,
        [string]$SkillName
    )

    $RefsDir = Join-Path $SkillPath "references"
    New-Item -ItemType Directory -Path $RefsDir -Force | Out-Null
    if (-not (Test-Path $SharedDir)) { throw "Required shared references not found: $SharedDir" }
    $FilesToCopy = @(Get-ChildItem -Path $SharedDir -File -Filter "*.md")
    if ($FilesToCopy.Count -eq 0) { throw "No shared reference files found: $SharedDir" }
    foreach ($sourceFile in $FilesToCopy) {
        $destination = Join-Path $RefsDir $sourceFile.Name
        $surface = $SkillName -replace '^excel-', ''
        $sourceContent = (Get-Content -LiteralPath $sourceFile.FullName -Raw) -replace "`r`n?", "`n"
        $rendered = [regex]::Replace($sourceContent, '(?ms)^```(?<surface>cli|mcp)\n(?<body>.*?)^```[ \t]*(?:\n|$)', {
            param($match)
            if ($match.Groups['surface'].Value -ne $surface) { return '' }
            $language = if ($surface -eq 'cli') { 'powershell' } else { 'text' }
            return '```' + $language + "`n" + $match.Groups['body'].Value + '```' + "`n"
        })
        Set-Content -LiteralPath $destination -Value $rendered -Encoding UTF8 -NoNewline
    }
    Write-Host "  Copied $($FilesToCopy.Count) shared references to $SkillName/references/" -ForegroundColor Green
}

if ($GenerateOnly) {
    if (-not $ManifestPath) {
        $ManifestPath = Join-Path $RepoRoot 'src\ExcelMcp.Core\obj\GeneratedFiles\ExcelMcp.Generators\Sbroenne.ExcelMcp.Generators.ServiceRegistryGenerator\_SkillManifest.g.cs'
    }
    if (-not (Test-Path -LiteralPath $ManifestPath -PathType Leaf)) { throw "Required manifest not found: $ManifestPath" }
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
    foreach ($component in @('cli', 'mcp')) {
        $name = "excel-$component"
        $destination = Join-Path $SkillsStagingDir $name
        if ($GenerateOnly) {
            $assets = Join-Path $SkillsDir "assets\$name"
            Copy-Item -LiteralPath $assets -Destination $destination -Recurse
            $projectName = if ($component -eq 'cli') { 'CLI' } else { 'McpServer' }
            $target = if ($component -eq 'cli') { 'GenerateCliSkill' } else { 'GenerateMcpSkill' }
            dotnet msbuild (Join-Path $RepoRoot "src\ExcelMcp.$projectName\ExcelMcp.$projectName.csproj") `
                "-target:$target" -p:Configuration=Release "-p:SkillOutputRoot=$SkillsStagingDir" `
                "-p:SkillManifestPath=$ManifestPath" -nodeReuse:false -verbosity:minimal
            if ($LASTEXITCODE -ne 0) { throw "$name rendering failed with exit code $LASTEXITCODE." }
            Copy-SharedReferences -SkillPath $destination -SkillName $name
            $references = Join-Path $destination 'references'
            $index = @('# Reference index', '', 'Read only the guide needed for the task. Discover command syntax, actions, and defaults from native help or MCP tool schemas.', '')
            foreach ($reference in (Get-ChildItem -LiteralPath $references -Filter '*.md' -File | Sort-Object Name)) {
                $title = (Get-Content -LiteralPath $reference.FullName -TotalCount 1) -replace '^#+\s*', ''
                $index += "- [$title]($($reference.Name))"
            }
            Set-Content -LiteralPath (Join-Path $references 'index.md') -Value ($index -join "`n") -Encoding UTF8
        }
        else {
            Copy-Item -LiteralPath (Join-Path $SkillsDirectory $name) -Destination $destination -Recurse
        }
        foreach ($required in @('SKILL.md', 'README.md', 'references\range.md')) {
            if (-not (Test-Path -LiteralPath (Join-Path $destination $required))) { throw "$name is missing $required." }
        }
        Set-Content -LiteralPath (Join-Path $destination 'VERSION') -Value $Version -Encoding utf8 -NoNewline
    }
    New-Item -ItemType Directory -Path $OutputPath -Force | Out-Null
    if ($GenerateOnly) {
        foreach ($name in @('excel-cli', 'excel-mcp')) {
            $destination = Join-Path $OutputPath $name
            Install-PackageOutput -Source (Join-Path $SkillsStagingDir $name) -Destination $destination
        }
        Write-Host "Generated complete skills at $OutputPath"
    }
    else {
        Copy-Item -LiteralPath (Join-Path $SkillsDir 'README.md') $StagingDir
        $zip = Join-Path $StagingDir "excel-skills-v$Version.zip"
        Compress-Archive -LiteralPath $SkillsStagingDir,(Join-Path $StagingDir 'README.md') -DestinationPath $zip
        Install-PackageOutput -Source $zip -Destination (Join-Path $OutputPath (Split-Path $zip -Leaf))
        Write-Host "Created $OutputPath\excel-skills-v$Version.zip"
    }
}
finally { Remove-Item -LiteralPath $StagingDir -Recurse -Force }
$global:LASTEXITCODE = 0
