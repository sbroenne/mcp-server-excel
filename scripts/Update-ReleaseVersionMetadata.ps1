<#
.SYNOPSIS
    Synchronizes persistent release version metadata across the source tree.

.DESCRIPTION
    Updates the version sources that must remain current after a release. Build-time
    placeholders such as the Agent Plugin manifests are intentionally not changed.

.PARAMETER Version
    The plain semantic version being released, without a leading "v".

.PARAMETER RepoRoot
    The repository root. Defaults to the parent directory of this script.
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [ValidatePattern('^\d+\.\d+\.\d+$')]
    [string]$Version,

    [string]$RepoRoot = (Split-Path $PSScriptRoot -Parent)
)

$ErrorActionPreference = 'Stop'
$RepoRoot = (Resolve-Path -LiteralPath $RepoRoot).Path

function Get-RequiredJson {
    param([Parameter(Mandatory)][string]$Path)

    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) {
        throw "Release metadata file was not found: $Path"
    }

    return Get-Content -LiteralPath $Path -Raw | ConvertFrom-Json
}

function Set-RootJsonVersion {
    param([Parameter(Mandatory)][string]$Path)

    $json = Get-RequiredJson -Path $Path
    if ($null -eq $json.PSObject.Properties['version']) {
        throw "Release metadata must contain a top-level 'version' property: $Path"
    }

    $content = Get-Content -LiteralPath $Path -Raw
    $pattern = '(?m)^(\s*"version"\s*:\s*")[^"]+(")'
    if ([regex]::Matches($content, $pattern).Count -lt 1) {
        throw "Release metadata version could not be located in: $Path"
    }

    $content = [regex]::new($pattern).Replace($content, "`${1}$Version`${2}", 1)
    [System.IO.File]::WriteAllText(
        $Path,
        $content,
        [System.Text.UTF8Encoding]::new($false))

    if ((Get-RequiredJson -Path $Path).version -ne $Version) {
        throw "Release metadata validation failed after stamping version '$Version': $Path"
    }
}

function Set-PackageLockVersion {
    param([Parameter(Mandatory)][string]$Path)

    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) {
        throw "Release metadata file was not found: $Path"
    }

    $json = Get-Content -LiteralPath $Path -Raw | ConvertFrom-Json -AsHashtable
    $rootPackage = $json.packages['']
    if (-not $json.ContainsKey('version') -or
        $null -eq $rootPackage -or
        -not $rootPackage.ContainsKey('version')) {
        throw "Package lock must contain top-level and root-package versions: $Path"
    }

    $content = Get-Content -LiteralPath $Path -Raw
    $pattern = '(?m)^(\s*"version"\s*:\s*")[^"]+(")'
    if ([regex]::Matches($content, $pattern).Count -lt 2) {
        throw "Package lock versions could not be located in: $Path"
    }

    $content = [regex]::new($pattern).Replace($content, "`${1}$Version`${2}", 2)
    [System.IO.File]::WriteAllText(
        $Path,
        $content,
        [System.Text.UTF8Encoding]::new($false))

    $updatedJson = Get-Content -LiteralPath $Path -Raw | ConvertFrom-Json -AsHashtable
    if ($updatedJson.version -ne $Version -or
        $updatedJson.packages[''].version -ne $Version) {
        throw "Package lock validation failed after stamping version '$Version': $Path"
    }
}

function Set-ProjectVersions {
    param([Parameter(Mandatory)][string]$Path)

    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) {
        throw "Directory.Build.props was not found: $Path"
    }

    $content = Get-Content -LiteralPath $Path -Raw
    $replacements = @{
        '<Version>[^<]+</Version>' = "<Version>$Version</Version>"
        '<AssemblyVersion>[^<]+</AssemblyVersion>' = "<AssemblyVersion>$Version.0</AssemblyVersion>"
        '<FileVersion>[^<]+</FileVersion>' = "<FileVersion>$Version.0</FileVersion>"
    }

    foreach ($replacement in $replacements.GetEnumerator()) {
        $versionMatches = [regex]::Matches($content, $replacement.Key)
        if ($versionMatches.Count -ne 1) {
            throw "Expected exactly one '$($replacement.Key)' value in $Path, found $($versionMatches.Count)."
        }

        $content = [regex]::Replace($content, $replacement.Key, $replacement.Value)
    }

    [System.IO.File]::WriteAllText(
        $Path,
        $content,
        [System.Text.UTF8Encoding]::new($false))
}

Set-RootJsonVersion -Path (Join-Path $RepoRoot 'package.json')
Set-PackageLockVersion -Path (Join-Path $RepoRoot 'package-lock.json')
Set-ProjectVersions -Path (Join-Path $RepoRoot 'Directory.Build.props')
Set-RootJsonVersion -Path (Join-Path $RepoRoot 'mcpb' 'manifest.json')
Set-RootJsonVersion -Path (Join-Path $RepoRoot 'vscode-extension' 'package.json')
Set-PackageLockVersion -Path (Join-Path $RepoRoot 'vscode-extension' 'package-lock.json')

& (Join-Path $PSScriptRoot 'Update-McpRegistryMetadata.ps1') `
    -ServerJsonPath (Join-Path $RepoRoot 'src' 'ExcelMcp.McpServer' '.mcp' 'server.json') `
    -Version $Version

Write-Output "Synchronized persistent release metadata to version $Version."
