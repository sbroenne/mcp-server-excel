[CmdletBinding()]
param(
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
$smokeScript = Join-Path $repoRoot 'npm-packages\mcp-server-excel\scripts\verify-runtime.mjs'
$resolvedLauncher = (Resolve-Path -LiteralPath $LauncherPackage).Path
$resolvedRuntime = (Resolve-Path -LiteralPath $RuntimePackage).Path
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

    $launcherScript = Join-Path $sandbox 'node_modules\@sbroenne\mcp-server-excel\bin\mcp-excel.js'
    $versionOutput = & node.exe $launcherScript --version 2>&1 | Out-String
    if ($LASTEXITCODE -ne 0) {
        throw "npm launcher --version failed with exit code $LASTEXITCODE. $versionOutput"
    }
    Write-Output ($versionOutput.Trim())

    $smokeOutput = & node.exe $smokeScript $launcherScript 2>&1 | Out-String
    if ($LASTEXITCODE -ne 0) {
        throw "npm launcher MCP handshake failed with exit code $LASTEXITCODE. $smokeOutput"
    }
    Write-Output ($smokeOutput.Trim())
}
finally {
    if (Test-Path -LiteralPath $sandbox) {
        Remove-Sandbox
    }
}
