[CmdletBinding()]
param(
    [ValidateSet('McpServer', 'Cli')]
    [string]$Component = 'McpServer',

    [Parameter(Mandatory)]
    [ValidateNotNullOrEmpty()]
    [string]$LauncherPackage,

    [Parameter(Mandatory)]
    [ValidateNotNullOrEmpty()]
    [string]$RuntimePackage
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

$repoRoot = Split-Path $PSScriptRoot -Parent
$packageName = if ($Component -eq 'Cli') { 'excelcli' } else { 'mcp-server-excel' }
$commandName = if ($Component -eq 'Cli') { 'excelcli' } else { 'mcp-excel' }
$smokeScript = Join-Path $repoRoot "npm-packages\$packageName\scripts\verify-runtime.mjs"
$resolvedLauncher = (Resolve-Path -LiteralPath $LauncherPackage).Path
$resolvedRuntime = (Resolve-Path -LiteralPath $RuntimePackage).Path
if (-not $IsWindows) {
    throw 'npm runtime smoke tests require Windows.'
}
$sandbox = Join-Path ([IO.Path]::GetTempPath()) "ExcelMcpNpmTest-$([Guid]::NewGuid().ToString('N'))"

function Remove-Sandbox {
    for ($attempt = 1; $attempt -le 20; $attempt++) {
        try {
            Remove-Item -LiteralPath $sandbox -Recurse -Force
            return
        }
        catch {
            if ($attempt -eq 20) {
                throw
            }

            Start-Sleep -Milliseconds 250
        }
    }
}

New-Item -ItemType Directory -Path $sandbox -Force | Out-Null

try {
    & npm.cmd install `
        --prefix $sandbox `
        --ignore-scripts `
        --no-audit `
        --no-fund `
        $resolvedRuntime `
        $resolvedLauncher
    if ($LASTEXITCODE -ne 0) {
        throw "npm package installation failed with exit code $LASTEXITCODE."
    }

    $launcherScript = Join-Path $sandbox "node_modules\@sbroenne\$packageName\bin\$commandName.js"
    $versionOutput = & node.exe $launcherScript --version 2>&1 | Out-String
    if ($LASTEXITCODE -ne 0) {
        throw "npm launcher --version failed with exit code $LASTEXITCODE. $versionOutput"
    }
    Write-Output ($versionOutput.Trim())

    $smokeOutput = & node.exe $smokeScript $launcherScript 2>&1 | Out-String
    if ($LASTEXITCODE -ne 0) {
        throw "$Component npm runtime smoke test failed with exit code $LASTEXITCODE. $smokeOutput"
    }
    Write-Output ($smokeOutput.Trim())
}
finally {
    if (Test-Path -LiteralPath $sandbox) {
        Remove-Sandbox
    }
}
