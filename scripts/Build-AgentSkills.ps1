#Requires -Version 7.4
<#
.SYNOPSIS
    Generates CLI discovery and report-formatting skills.
#>
[CmdletBinding()]
param(
    [string]$OutputDir,
    [string]$Version,
    [switch]$GenerateOnly
)
$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'PackageHelpers.ps1')
$root = Split-Path $PSScriptRoot -Parent
if (-not $GenerateOnly) { throw 'This script only prepares generated skills. Use -GenerateOnly.' }
if (-not $Version) { $Version = (Read-PackageJson (Join-Path $root 'package.json')).version }
$Version = Get-PackageVersion $Version
if (-not $OutputDir) { $OutputDir = 'artifacts\generated-skills' }
$output = [IO.Path]::GetFullPath($OutputDir, $root)
Assert-PackageOutputPath $output $root
$names = @('excel-cli', 'excel-cli-report-formatting', 'excel-mcp-report-formatting')
Invoke-PackageStage 'skills' {
    param($stage)
    foreach ($name in $names) {
        $destination = Join-Path $stage $name
        Copy-PackageItem (Join-Path $root "skills\$name") $destination
        if ($name.EndsWith('-report-formatting', [StringComparison]::Ordinal)) {
            $surface = if ($name.StartsWith('excel-cli-', [StringComparison]::Ordinal)) { 'cli' } else { 'mcp' }
            $content = [IO.File]::ReadAllText((Join-Path $root 'docs\reference\report-formatting.md')).Replace("`r`n", "`n").Replace("`r", "`n")
            $content = [regex]::Replace($content, '(?ms)^```(?<surface>cli|mcp)\n(?<body>.*?)^```[ \t]*(?:\n|$)', {
                param($match)
                if ($match.Groups['surface'].Value -ceq $surface) {
                    $language = if ($surface -eq 'cli') { 'powershell' } else { 'text' }
                    return '```' + $language + "`n" + $match.Groups['body'].Value + '```' + "`n"
                }
                return ''
            })
            $content = [regex]::Replace($content, '\[([^\]]+)\]\(([a-z-]+)\.md(#[^)]*)?\)', {
                param($match)
                "[$($match.Groups[1].Value)](https://excelmcpserver.dev/reference/$($match.Groups[2].Value)/$($match.Groups[3].Value))"
            })
            [IO.Directory]::CreateDirectory((Join-Path $destination 'references')) | Out-Null
            [IO.File]::WriteAllText((Join-Path $destination 'references\report-formatting.md'), $content)
        }
        if (-not [IO.File]::Exists((Join-Path $destination 'SKILL.md'))) { throw "$name is missing SKILL.md." }
        [IO.File]::WriteAllText((Join-Path $destination 'VERSION'), $Version)
    }
    [IO.Directory]::CreateDirectory($output) | Out-Null
    foreach ($name in $names) { Install-PackageOutput (Join-Path $stage $name) (Join-Path $output $name) }
    $legacy = Join-Path $output 'excel-mcp'
    if ([IO.Directory]::Exists($legacy)) {
        Assert-PackageOutputPath $legacy $root
        if (-not [IO.File]::Exists((Join-Path $legacy 'SKILL.md')) -or -not [IO.File]::Exists((Join-Path $legacy 'VERSION'))) {
            throw "Refusing to remove unrecognized retired skill output: $legacy"
        }
        Remove-PackageItem $legacy
    }
    $output
}
