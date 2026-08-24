[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [ValidateNotNullOrEmpty()]
    [string]$Version,

    [Parameter(Mandatory)]
    [ValidateNotNullOrEmpty()]
    [string]$RuntimeExecutable,

    [Parameter(Mandatory)]
    [ValidateNotNullOrEmpty()]
    [string]$OutputDirectory
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

$repoRoot = Split-Path $PSScriptRoot -Parent
$launcherSource = Join-Path $repoRoot 'npm-packages\mcp-server-excel'
$runtimeSource = Join-Path $repoRoot 'npm-packages\mcp-server-excel-win32-x64'
$licensePath = Join-Path $repoRoot 'LICENSE'
$resolvedRuntime = (Resolve-Path -LiteralPath $RuntimeExecutable).Path
$resolvedOutput = [IO.Path]::GetFullPath($OutputDirectory)
$stagingRoot = Join-Path ([IO.Path]::GetTempPath()) "ExcelMcpNpm-$([Guid]::NewGuid().ToString('N'))"
$launcherStage = Join-Path $stagingRoot 'mcp-server-excel'
$runtimeStage = Join-Path $stagingRoot 'mcp-server-excel-win32-x64'

function Copy-PackageSource {
    param(
        [Parameter(Mandatory)]
        [string]$Source,

        [Parameter(Mandatory)]
        [string]$Destination,

        [Parameter(Mandatory)]
        [string[]]$Entries
    )

    New-Item -ItemType Directory -Path $Destination -Force | Out-Null
    foreach ($entry in $Entries) {
        Copy-Item -LiteralPath (Join-Path $Source $entry) -Destination $Destination -Recurse
    }
    Copy-Item -LiteralPath $licensePath -Destination $Destination
}

function Write-PackageManifest {
    param(
        [Parameter(Mandatory)]
        [string]$Path,

        [Parameter(Mandatory)]
        [scriptblock]$Update
    )

    $manifest = Get-Content -LiteralPath $Path -Raw | ConvertFrom-Json
    & $Update $manifest
    $json = ($manifest | ConvertTo-Json -Depth 20) -replace "`r?`n", "`n"
    Set-Content -LiteralPath $Path -Value $json -NoNewline
}

function New-NpmTarball {
    param(
        [Parameter(Mandatory)]
        [string]$PackageDirectory,

        [Parameter(Mandatory)]
        [string[]]$RequiredFiles
    )

    $manifest = Get-Content -LiteralPath (Join-Path $PackageDirectory 'package.json') -Raw | ConvertFrom-Json
    $archiveName = "$(($manifest.name.TrimStart('@')) -replace '/', '-')-$($manifest.version).tgz"
    $archivePath = Join-Path $resolvedOutput $archiveName
    $inspectionDirectory = Join-Path $stagingRoot "inspect-$([Guid]::NewGuid().ToString('N'))"

    if (Test-Path -LiteralPath $archivePath) {
        Remove-Item -LiteralPath $archivePath -Force
    }

    & npm.cmd pack $PackageDirectory --pack-destination $resolvedOutput --silent | Out-Null
    if ($LASTEXITCODE -ne 0) {
        throw "npm pack failed for '$PackageDirectory' with exit code $LASTEXITCODE."
    }

    if (-not (Test-Path -LiteralPath $archivePath -PathType Leaf)) {
        throw "npm pack did not create the expected archive '$archivePath'."
    }

    New-Item -ItemType Directory -Path $inspectionDirectory -Force | Out-Null
    try {
        & tar -xf $archivePath -C $inspectionDirectory
        if ($LASTEXITCODE -ne 0) {
            throw "Could not inspect packed npm package '$archiveName'."
        }

        foreach ($requiredFile in $RequiredFiles) {
            $packedPath = Join-Path (Join-Path $inspectionDirectory 'package') $requiredFile
            if (-not (Test-Path -LiteralPath $packedPath -PathType Leaf)) {
                throw "Packed npm package '$archiveName' is missing required file '$requiredFile'."
            }
        }
    }
    finally {
        Remove-Item -LiteralPath $inspectionDirectory -Recurse -Force
    }

    return $archivePath
}

if (-not (Test-Path -LiteralPath $launcherSource -PathType Container) -or
    -not (Test-Path -LiteralPath $runtimeSource -PathType Container)) {
    throw 'npm package source directories are missing.'
}

if ([IO.Path]::GetExtension($resolvedRuntime) -ne '.exe') {
    throw "Runtime executable must be an .exe file: $resolvedRuntime"
}

New-Item -ItemType Directory -Path $resolvedOutput -Force | Out-Null
New-Item -ItemType Directory -Path $stagingRoot -Force | Out-Null

try {
    Copy-PackageSource `
        -Source $launcherSource `
        -Destination $launcherStage `
        -Entries @('package.json', 'README.md', 'bin', 'lib')
    Copy-PackageSource `
        -Source $runtimeSource `
        -Destination $runtimeStage `
        -Entries @('package.json', 'README.md')
    Copy-Item -LiteralPath $resolvedRuntime -Destination (Join-Path $runtimeStage 'mcp-excel.exe')

    Write-PackageManifest -Path (Join-Path $runtimeStage 'package.json') -Update {
        param($manifest)
        $manifest.version = $Version
    }
    Write-PackageManifest -Path (Join-Path $launcherStage 'package.json') -Update {
        param($manifest)
        $manifest.version = $Version
        $manifest.optionalDependencies.'@sbroenne/mcp-server-excel-win32-x64' = $Version
    }

    $runtimeTarball = New-NpmTarball `
        -PackageDirectory $runtimeStage `
        -RequiredFiles @('mcp-excel.exe', 'package.json')
    $launcherTarball = New-NpmTarball `
        -PackageDirectory $launcherStage `
        -RequiredFiles @('bin/mcp-excel.js', 'lib/launcher.js', 'package.json')

    [pscustomobject]@{
        LauncherPackage = $launcherTarball
        RuntimePackage = $runtimeTarball
    } | ConvertTo-Json -Compress
}
finally {
    if (Test-Path -LiteralPath $stagingRoot) {
        Remove-Item -LiteralPath $stagingRoot -Recurse -Force
    }
}
