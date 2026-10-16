#Requires -Version 7.4
[CmdletBinding()]
param(
    [ValidateSet('McpServer', 'Cli')][string]$Component = 'McpServer',
    [ValidateSet('x64', 'arm64')][string]$Architecture = 'x64',
    [Parameter(Mandatory)][ValidateNotNullOrEmpty()][string]$Version,
    [Parameter(Mandatory)][ValidateNotNullOrEmpty()][string]$RuntimeExecutable,
    [Parameter(Mandatory)][ValidateNotNullOrEmpty()][string]$OutputDirectory
)
$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'PackageHelpers.ps1')
$root = Split-Path $PSScriptRoot -Parent
$Version = Get-PackageVersion $Version
$RuntimeExecutable = [IO.Path]::GetFullPath($RuntimeExecutable, $root)
$OutputDirectory = [IO.Path]::GetFullPath($OutputDirectory, $root)
Assert-PackageOutputPath $OutputDirectory $root @($RuntimeExecutable)
if ([IO.Path]::GetExtension($RuntimeExecutable) -ine '.exe') { throw "Runtime executable must be an .exe file: $RuntimeExecutable" }
Assert-PackageRuntimeArchitecture $RuntimeExecutable $Architecture
$name, $command = if ($Component -eq 'Cli') { 'excelcli'; 'excelcli' } else { 'mcp-server-excel'; 'mcp-excel' }

function New-NpmArchive {
    param([string]$Directory, [string]$Stage, [string[]]$Required)
    $manifest = Read-PackageJson (Join-Path $Directory 'package.json')
    $archiveName = $manifest.name.TrimStart('@').Replace('/', '-') + "-$($manifest.version).tgz"
    $archive = Join-Path $OutputDirectory $archiveName
    Remove-PackageItem $archive
    Invoke-PackageCommand $root 'npm' @('pack', $Directory, '--pack-destination', $OutputDirectory, '--silent') | Out-Null
    if (-not [IO.File]::Exists($archive)) { throw "npm pack did not create the expected archive '$archive'." }
    $inspection = Join-Path $Stage "inspect-$([Guid]::NewGuid().ToString('N'))"
    Expand-PackageTar $archive $inspection
    foreach ($file in $Required) {
        if (-not [IO.File]::Exists((Join-Path $inspection "package\$file"))) { throw "Packed npm package '$archiveName' is missing required file '$file'." }
    }
    $archive
}

Invoke-PackageStage 'npm' {
    param($stage)
    $launcher = Join-Path $stage $name
    $runtime = Join-Path $stage "$name-win32-$Architecture"
    foreach ($directory in $launcher, $runtime) {
        $source = Join-Path $root "npm-packages\$([IO.Path]::GetFileName($directory))"
        foreach ($entry in 'package.json', 'README.md') { Copy-PackageItem (Join-Path $source $entry) (Join-Path $directory $entry) }
        Copy-PackageItem (Join-Path $root 'LICENSE') (Join-Path $directory 'LICENSE')
    }
    Copy-PackageItem (Join-Path $root "npm-packages\$name\bin") (Join-Path $launcher 'bin')
    Copy-PackageItem (Join-Path $root 'npm-packages\shared\launcher.js') (Join-Path $launcher 'lib\launcher.js')
    Copy-PackageItem $RuntimeExecutable (Join-Path $runtime "$command.exe")
    foreach ($directory in $runtime, $launcher) {
        $path = Join-Path $directory 'package.json'
        $manifest = Read-PackageJson $path
        $manifest.version = $Version
        if ($directory -eq $launcher) {
            if ($manifest.optionalDependencies -isnot [Collections.IDictionary]) { throw "Missing object optionalDependencies: $path" }
            foreach ($arch in 'x64', 'arm64') { $manifest.optionalDependencies["@sbroenne/$name-win32-$arch"] = $Version }
        }
        Write-PackageJson $path $manifest
    }
    [IO.Directory]::CreateDirectory($OutputDirectory) | Out-Null
    $runtimeArchive = New-NpmArchive $runtime $stage @("$command.exe", 'package.json', 'LICENSE')
    $launcherArchive = New-NpmArchive $launcher $stage @("bin\$command.js", 'lib\launcher.js', 'package.json', 'LICENSE')
    @{ LauncherPackage = $launcherArchive; RuntimePackage = $runtimeArchive } | ConvertTo-Json -Compress
}
