#Requires -Version 7.4
[CmdletBinding()]
param(
    [ValidateSet('McpServer', 'Cli')][string]$Component = 'McpServer',
    [ValidateSet('x64', 'arm64')][string]$Architecture = 'x64',
    [switch]$ArchiveOnly,
    [Parameter(Mandatory)][ValidateNotNullOrEmpty()][string]$LauncherPackage,
    [Parameter(Mandatory)][ValidateNotNullOrEmpty()][string]$RuntimePackage
)
$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'PackageHelpers.ps1')
$root = Split-Path $PSScriptRoot -Parent
$LauncherPackage = [IO.Path]::GetFullPath($LauncherPackage, $root)
$RuntimePackage = [IO.Path]::GetFullPath($RuntimePackage, $root)
$name, $command = if ($Component -eq 'Cli') { 'excelcli'; 'excelcli' } else { 'mcp-server-excel'; 'mcp-excel' }
Invoke-PackageStage 'npmtest' {
    param($stage)
    $version = $null
    foreach ($kind in 'runtime', 'launcher') {
        $archive = if ($kind -eq 'runtime') { $RuntimePackage } else { $LauncherPackage }
        $directory = Join-Path $stage $kind
        Expand-PackageTar $archive $directory
        $package = Join-Path $directory 'package'
        $manifest = Read-PackageJson (Join-Path $package 'package.json')
        $packageName = if ($kind -eq 'runtime') { "$name-win32-$Architecture" } else { $name }
        if ($manifest.name -cne "@sbroenne/$packageName") { throw "Unexpected npm package name: $($manifest.name)" }
        if (-not [IO.File]::Exists((Join-Path $package 'LICENSE'))) { throw "Missing license in $kind npm archive." }
        if ($kind -eq 'runtime') {
            $version = Get-PackageVersion $manifest.version
            if ($manifest.main -cne "$command.exe" -or $manifest.os -isnot [array] -or $manifest.cpu -isnot [array] -or
                $manifest.os.Count -ne 1 -or $manifest.os[0] -cne 'win32' -or $manifest.cpu.Count -ne 1 -or $manifest.cpu[0] -cne $Architecture) {
                throw 'npm runtime metadata does not match the requested architecture.'
            }
            Assert-PackageRuntimeArchitecture (Join-Path $package "$command.exe") $Architecture
        } else {
            if ($manifest.version -cne $version -or $manifest.bin -isnot [Collections.IDictionary] -or $manifest.bin[$command] -cne "bin/$command.js") {
                throw 'npm launcher metadata does not match the runtime.'
            }
            foreach ($arch in 'x64', 'arm64') {
                if ($manifest.optionalDependencies -isnot [Collections.IDictionary] -or $manifest.optionalDependencies["@sbroenne/$name-win32-$arch"] -cne $version) {
                    throw "npm launcher must depend on the matching $arch release version."
                }
            }
            foreach ($file in "bin\$command.js", 'lib\launcher.js') {
                if (-not [IO.File]::Exists((Join-Path $package $file))) { throw "Missing launcher file: $file" }
            }
        }
    }
    Write-Host "$Component $Architecture npm archives validated."
    if ($ArchiveOnly) { Write-Host 'Archive-only validation requested; native execution is a separate check.'; return }
    if (-not $IsWindows) { throw 'npm runtime smoke tests require Windows.' }
    $nodeArchitecture = (Invoke-PackageCommand $stage 'node' @('-p', 'process.arch')).Trim()
    if ($nodeArchitecture -cne $Architecture) {
        [Console]::Error.WriteLine("$Component $Architecture execution NOT RUN: Node.js is $nodeArchitecture. Archive validation passed.")
        return
    }
    Invoke-PackageCommand $stage 'npm' @('install', '--prefix', $stage, '--ignore-scripts', '--no-audit', '--no-fund', $RuntimePackage, $LauncherPackage) | Out-Host
    $installed = Join-Path $stage "node_modules\@sbroenne\$name\bin\$command.js"
    Invoke-PackageCommand $stage 'node' @($installed, '--version') | Out-Host
    Invoke-PackageCommand $stage 'node' @((Join-Path $root "npm-packages\$name\scripts\verify-runtime.mjs"), $installed) | Out-Host
}
