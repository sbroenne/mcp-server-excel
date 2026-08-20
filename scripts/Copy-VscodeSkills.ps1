<#
.SYNOPSIS
    Copies the canonical MCP skill into the VS Code extension and stamps its package version.
#>
$ErrorActionPreference = "Stop"
$RepoRoot = Split-Path -Parent $PSScriptRoot
$SourceDir = Join-Path $RepoRoot "skills\excel-mcp"
$OutputDir = Join-Path $RepoRoot "vscode-extension\skills\excel-mcp"
$PackageJsonPath = Join-Path $RepoRoot "vscode-extension\package.json"

if (-not (Test-Path $SourceDir -PathType Container)) {
    throw "Canonical Agent Skill directory not found: $SourceDir"
}

$packageJson = Get-Content $PackageJsonPath -Raw | ConvertFrom-Json
$version = $packageJson.version
if ([string]::IsNullOrWhiteSpace($version)) {
    throw "VS Code extension version is missing from $PackageJsonPath."
}

if (Test-Path $OutputDir) {
    Remove-Item -Path $OutputDir -Recurse -Force
}

New-Item -ItemType Directory -Path (Split-Path -Parent $OutputDir) -Force | Out-Null
Copy-Item -Path $SourceDir -Destination $OutputDir -Recurse -Force
Set-Content -Path (Join-Path $OutputDir "VERSION") -Value $version.Trim() -Encoding UTF8 -NoNewline

Write-Host "Copied excel-mcp skill and stamped version $($version.Trim()) at $OutputDir" -ForegroundColor Green
