[CmdletBinding()]
param(
    [ValidateSet('McpServer', 'Cli')]
    [string]$Component = 'McpServer',

    [ValidateSet('x64', 'arm64')]
    [string]$Architecture = 'x64',

    [switch]$ArchiveOnly,

    [Parameter(Mandatory)]
    [ValidateNotNullOrEmpty()]
    [string]$LauncherPackage,

    [Parameter(Mandatory)]
    [ValidateNotNullOrEmpty()]
    [string]$RuntimePackage
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest
. (Join-Path $PSScriptRoot 'PackageHelpers.ps1')

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
    foreach ($entry in @(
        @{ Archive = $resolvedRuntime; Name = "$packageName-win32-$Architecture"; Kind = 'runtime' },
        @{ Archive = $resolvedLauncher; Name = $packageName; Kind = 'launcher' }
    )) {
        $inspection = Join-Path $sandbox $entry.Kind
        New-Item -ItemType Directory -Path $inspection | Out-Null
        & tar -xf $entry.Archive -C $inspection
        if ($LASTEXITCODE -ne 0) { throw "Could not inspect $($entry.Kind) npm archive." }
        $packageRoot = Join-Path $inspection 'package'
        $manifest = Get-Content (Join-Path $packageRoot 'package.json') -Raw | ConvertFrom-Json
        if ($manifest.name -ne "@sbroenne/$($entry.Name)") {
            throw "Unexpected npm package name: $($manifest.name)"
        }
        if (-not (Test-Path -LiteralPath (Join-Path $packageRoot 'LICENSE') -PathType Leaf)) {
            throw "Missing license in $($entry.Kind) npm archive."
        }
        if ($entry.Kind -eq 'runtime') {
            $runtimeVersion = $manifest.version
            if ($manifest.main -ne "$commandName.exe" -or
                @($manifest.os).Count -ne 1 -or $manifest.os[0] -ne 'win32' -or
                @($manifest.cpu).Count -ne 1 -or $manifest.cpu[0] -ne $Architecture) {
                throw 'npm runtime metadata does not match the requested architecture.'
            }
            Assert-PackageRuntimeArchitecture -Path (Join-Path $packageRoot $manifest.main) -Architecture $Architecture
        } else {
            if ($manifest.version -ne $runtimeVersion -or $manifest.bin.$commandName -ne "bin/$commandName.js") {
                throw 'npm launcher metadata does not match the runtime.'
            }
            foreach ($arch in @('x64', 'arm64')) {
                if ($manifest.optionalDependencies."@sbroenne/$packageName-win32-$arch" -ne $runtimeVersion) {
                    throw "npm launcher must depend on the matching $arch release version."
                }
            }
            foreach ($file in @("bin/$commandName.js", 'lib/launcher.js')) {
                if (-not (Test-Path -LiteralPath (Join-Path $packageRoot $file) -PathType Leaf)) {
                    throw "Missing launcher file: $file"
                }
            }
        }
    }
    Write-Output "$Component $Architecture npm archives validated."
    if ($ArchiveOnly) {
        Write-Output "$Component $Architecture archive-only validation requested; native execution is a separate check."
        return
    }
    $nodeArchitecture = (& node.exe -p 'process.arch' | Out-String).Trim()
    if ($LASTEXITCODE -ne 0) { throw 'Could not determine Node.js architecture.' }
    if ($nodeArchitecture -ne $Architecture) {
        Write-Warning "$Component $Architecture execution NOT RUN: Node.js is $nodeArchitecture. Archive validation passed."
        return
    }

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
