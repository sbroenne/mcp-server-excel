#!/usr/bin/env pwsh
[CmdletBinding()]
param(
    [switch]$Check,
    [string]$OutputRoot = $null
)

$ErrorActionPreference = "Stop"
Set-StrictMode -Version Latest

$RepoRoot = Split-Path -Parent $PSScriptRoot
$SharedDir = Join-Path $RepoRoot ".github\plugins\_shared"
$TemplatePath = Join-Path $SharedDir "download.ps1.template"
$TokenMapPath = Join-Path $SharedDir "bootstrap-tokens.json"

if ([string]::IsNullOrWhiteSpace($OutputRoot)) {
    $OutputRoot = Join-Path $RepoRoot ".github\plugins"
}

$PluginNames = @("excel-cli", "excel-mcp")

function Get-TokenMap {
    if (-not (Test-Path $TokenMapPath)) {
        throw "Missing bootstrap token map: $TokenMapPath"
    }

    $raw = Get-Content $TokenMapPath -Raw
    $map = $raw | ConvertFrom-Json
    $result = @{}

    foreach ($pluginName in $PluginNames) {
        $pluginConfig = $map.$pluginName
        if ($null -eq $pluginConfig) {
            throw "No bootstrap token config for $pluginName in $TokenMapPath"
        }

        $result[$pluginName] = @{
            "{{PLUGIN_NAME}}" = [string]$pluginConfig.plugin_name
            "{{RUNTIME_EXE}}" = [string]$pluginConfig.runtime_exe
            "{{ASSET_PREFIX}}" = [string]$pluginConfig.asset_prefix
            "{{RELEASE_DESCRIPTION}}" = [string]$pluginConfig.release_description
            "{{READY_MESSAGE}}" = [string]$pluginConfig.ready_message
        }
    }

    return $result
}

function Get-RenderedScriptText {
    param(
        [Parameter(Mandatory = $true)][string]$PluginName,
        [Parameter(Mandatory = $true)][hashtable]$TokenMap
    )

    if (-not (Test-Path $TemplatePath)) {
        throw "Missing bootstrap template: $TemplatePath"
    }

    $scriptText = Get-Content $TemplatePath -Raw
    foreach ($token in $TokenMap[$PluginName].Keys) {
        $scriptText = $scriptText.Replace($token, $TokenMap[$PluginName][$token])
    }

    return $scriptText
}

function Write-Utf8NoBomText {
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [Parameter(Mandatory = $true)][string]$Content
    )

    $parent = Split-Path -Parent $Path
    if (-not [string]::IsNullOrWhiteSpace($parent) -and -not (Test-Path $parent)) {
        New-Item -ItemType Directory -Path $parent -Force | Out-Null
    }

    [System.IO.File]::WriteAllText($Path, $Content, [System.Text.UTF8Encoding]::new($false))
}

$tokenMap = Get-TokenMap
$failures = @()

foreach ($pluginName in $PluginNames) {
    $rendered = Get-RenderedScriptText -PluginName $pluginName -TokenMap $tokenMap
    $targetDir = Join-Path (Join-Path $OutputRoot $pluginName) "bin"
    $targetPath = Join-Path $targetDir "download.ps1"

    if ($Check) {
        if (-not (Test-Path $targetPath)) {
            $failures += "Missing bootstrap script for ${pluginName}: ${targetPath}"
            continue
        }

        $actualBytes = [System.IO.File]::ReadAllBytes($targetPath)
        $expectedBytes = [System.Text.UTF8Encoding]::new($false).GetBytes($rendered)

        $matches = $actualBytes.Length -eq $expectedBytes.Length
        if ($matches) {
            for ($index = 0; $index -lt $actualBytes.Length; $index++) {
                if ($actualBytes[$index] -ne $expectedBytes[$index]) {
                    $matches = $false
                    break
                }
            }
        }

        if (-not $matches) {
            $failures += "Bootstrap drift detected for $pluginName in $targetPath. Run scripts/Build-BootstrapScripts.ps1 to regenerate."
        }

        continue
    }

    Write-Utf8NoBomText -Path $targetPath -Content $rendered
    Write-Host "Generated $pluginName bootstrap: $targetPath" -ForegroundColor Green
}

if ($Check) {
    if ($failures.Count -gt 0) {
        foreach ($failure in $failures) {
            Write-Host $failure -ForegroundColor Red
        }

        exit 1
    }

    Write-Host "Bootstrap scripts are matching the canonical template." -ForegroundColor Green
    exit 0
}

Write-Host "Bootstrap templates rendered under $OutputRoot" -ForegroundColor Cyan
exit 0
