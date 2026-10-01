[CmdletBinding()]
param(
    [ValidateSet('McpServer', 'Cli')]
    [string]$Component = 'McpServer',

    [ValidateSet('x64', 'arm64')]
    [string]$Architecture = 'x64',

    [Parameter(Mandatory)]
    [ValidateNotNullOrEmpty()]
    [string]$Version,

    [ValidateSet('win-x64', 'win-arm64', 'osx-arm64')]
    [string]$RuntimeIdentifier = 'win-x64',

    [Parameter(Mandatory)]
    [ValidateNotNullOrEmpty()]
    [string]$RuntimeExecutable,

    [Parameter(Mandatory)]
    [ValidateNotNullOrEmpty()]
    [string]$OutputDirectory
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest
. (Join-Path $PSScriptRoot 'PackageHelpers.ps1')

$repoRoot = Split-Path $PSScriptRoot -Parent
if (-not $PSBoundParameters.ContainsKey('RuntimeIdentifier')) {
    $RuntimeIdentifier = "win-$Architecture"
}
$packageName = if ($Component -eq 'Cli') { 'excelcli' } else { 'mcp-server-excel' }
$commandName = if ($Component -eq 'Cli') { 'excelcli' } else { 'mcp-excel' }
$runtimePackageSuffix = switch ($RuntimeIdentifier) {
    'osx-arm64' { 'darwin-arm64' }
    'win-arm64' { 'win32-arm64' }
    default { 'win32-x64' }
}
$runtimeFileName = if ($RuntimeIdentifier.StartsWith('osx-', [StringComparison]::Ordinal)) { $commandName } else { "$commandName.exe" }
$launcherSource = Join-Path $repoRoot "npm-packages\$packageName"
$runtimeSource = Join-Path $repoRoot "npm-packages\$packageName-$runtimePackageSuffix"
$sharedLauncher = Join-Path $repoRoot 'npm-packages\shared\launcher.js'
$npmCommand = if ($IsWindows) { 'npm.cmd' } else { 'npm' }
$licensePath = Join-Path $repoRoot 'LICENSE'
$resolvedRuntime = (Resolve-Path -LiteralPath $RuntimeExecutable).Path
$runtimeDirectory = Split-Path -Parent $resolvedRuntime
$helperSource = Join-Path $runtimeDirectory 'helpers/excelmcp-screencapture'
$resolvedOutput = [IO.Path]::GetFullPath($OutputDirectory)
$stagingRoot = Join-Path ([IO.Path]::GetTempPath()) "ExcelMcpNpm-$([Guid]::NewGuid().ToString('N'))"
$launcherStage = Join-Path $stagingRoot $packageName
$runtimeStage = Join-Path $stagingRoot "$packageName-$runtimePackageSuffix"

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
    [System.IO.File]::WriteAllText(
        $Path,
        $json,
        [System.Text.UTF8Encoding]::new($false))
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

    & $npmCommand pack $PackageDirectory --pack-destination $resolvedOutput --silent | Out-Null
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

if ($RuntimeIdentifier.StartsWith('win-', [StringComparison]::Ordinal) -and [IO.Path]::GetExtension($resolvedRuntime) -ne '.exe') {
    throw "Runtime executable must be an .exe file: $resolvedRuntime"
}
if ($RuntimeIdentifier.StartsWith('osx-', [StringComparison]::Ordinal) -and
    -not (Test-Path -LiteralPath $helperSource -PathType Leaf)) {
    throw "macOS runtime package requires the ScreenCaptureKit helper: $helperSource"
}
if ($RuntimeIdentifier.StartsWith('win-', [StringComparison]::Ordinal)) {
    Assert-PackageRuntimeArchitecture -Path $resolvedRuntime -Architecture $RuntimeIdentifier.Substring(4)
}

New-Item -ItemType Directory -Path $resolvedOutput -Force | Out-Null
New-Item -ItemType Directory -Path $stagingRoot -Force | Out-Null

try {
    Copy-PackageSource `
        -Source $launcherSource `
        -Destination $launcherStage `
        -Entries @('package.json', 'README.md', 'bin')
    $launcherLib = New-Item -ItemType Directory -Path (Join-Path $launcherStage 'lib')
    Copy-Item -LiteralPath $sharedLauncher -Destination $launcherLib.FullName
    Copy-PackageSource `
        -Source $runtimeSource `
        -Destination $runtimeStage `
        -Entries @('package.json', 'README.md')
    Copy-Item -LiteralPath $resolvedRuntime -Destination (Join-Path $runtimeStage $runtimeFileName)
    $runtimeRequiredFiles = @($runtimeFileName, 'package.json', 'LICENSE')
    if ($RuntimeIdentifier.StartsWith('osx-', [StringComparison]::Ordinal)) {
        $helperStage = Join-Path $runtimeStage 'helpers'
        New-Item -ItemType Directory -Path $helperStage -Force | Out-Null
        Copy-Item -LiteralPath $helperSource -Destination $helperStage
        $runtimeRequiredFiles += 'helpers/excelmcp-screencapture'
    }

    Write-PackageManifest -Path (Join-Path $runtimeStage 'package.json') -Update {
        param($manifest)
        $manifest.version = $Version
    }
    Write-PackageManifest -Path (Join-Path $launcherStage 'package.json') -Update {
        param($manifest)
        $manifest.version = $Version
        foreach ($property in $manifest.optionalDependencies.PSObject.Properties) {
            $property.Value = $Version
        }
    }

    $runtimeTarball = New-NpmTarball `
        -PackageDirectory $runtimeStage `
        -RequiredFiles $runtimeRequiredFiles
    $launcherTarball = New-NpmTarball `
        -PackageDirectory $launcherStage `
        -RequiredFiles @("bin/$commandName.js", 'lib/launcher.js', 'package.json', 'LICENSE')

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
