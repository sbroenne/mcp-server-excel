#Requires -Version 7.4
<#
.SYNOPSIS
    Creates and verifies distributable packages. Never publishes them.
.DESCRIPTION
    Checks all components by default. Explicit runs may name components.
#>
[CmdletBinding()]
param(
    [ValidateSet('Cli', 'Mcp', 'Extension', 'Mcpb', 'Plugins')]
    [string[]]$Components = @('Cli', 'Mcp', 'Extension', 'Mcpb', 'Plugins'),
    [string]$Version,
    [string]$SkillsDirectory,
    [string]$McpRuntimeExecutable,
    [string]$CliRuntimeExecutable,
    [string]$OutputDirectory,
    [switch]$SkipExtensionTests
)
$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'PackageHelpers.ps1')
. (Join-Path $PSScriptRoot 'ExtensionPackaging.ps1')
$root = Split-Path $PSScriptRoot -Parent
$Components = @($Components | Select-Object -Unique)
if (-not $Components.Count) { throw 'At least one release package component is required.' }
if (-not $IsWindows) { throw 'Package installation checks require Windows (not Excel).' }
if (-not $Version) { $Version = (Read-PackageJson (Join-Path $root 'package.json')).version }
$Version = Get-PackageVersion $Version
$needsSkills = 'Plugins' -in $Components -or 'Extension' -in $Components
if ($SkillsDirectory) {
    $SkillsDirectory = [IO.Path]::GetFullPath($SkillsDirectory, $root)
    if ($needsSkills) {
        $names = if ('Plugins' -in $Components) { @('excel-cli', 'excel-cli-report-formatting', 'excel-mcp-report-formatting') } else { @('excel-mcp-report-formatting') }
        foreach ($name in $names) {
            $stamp = Join-Path $SkillsDirectory "$name\VERSION"
            if (-not [IO.File]::Exists($stamp) -or [IO.File]::ReadAllText($stamp).Trim() -cne $Version) { throw "Prepared $name skill must match package version $Version." }
        }
    }
}
if (-not $OutputDirectory) { $OutputDirectory = "artifacts\packages\$([Guid]::NewGuid().ToString('N'))" }
$output = [IO.Path]::GetFullPath($OutputDirectory, $root)
Assert-PackageOutputPath $output $root @($SkillsDirectory, $McpRuntimeExecutable, $CliRuntimeExecutable)
if ([IO.Path]::Exists($output)) { throw "Use a new package output directory: $output" }
[IO.Directory]::CreateDirectory($output) | Out-Null

function Publish-ReleaseRuntime {
    param([string]$Component, [string]$Architecture)
    $projectName = if ($Component -eq 'Cli') { 'CLI' } else { 'McpServer' }
    $directoryName = if ($Architecture -eq 'x64') { $Component } else { "$Component-$Architecture" }
    $destination = Join-Path $output "runtimes\$directoryName"
    Invoke-PackageCommand $root 'dotnet' @('publish', (Join-Path $root "src\ExcelMcp.$projectName\ExcelMcp.$projectName.csproj"),
        '-c', 'Release', '-r', "win-$Architecture", '--self-contained', 'true', '-p:PublishSingleFile=true',
        '-p:IncludeNativeLibrariesForSelfExtract=true', '-p:PublishTrimmed=false', '-p:PublishReadyToRun=false',
        '-p:NuGetAudit=false', "-p:Version=$Version", '-o', $destination, '--verbosity', 'minimal') | Out-Host
    Join-Path $destination $(if ($Component -eq 'Cli') { 'excelcli.exe' } else { 'Sbroenne.ExcelMcp.McpServer.exe' })
}

function Assert-ReleaseRuntime {
    param([string]$Path, [string]$Architecture)
    Assert-PackageRuntimeArchitecture $Path $Architecture
    $product = [Diagnostics.FileVersionInfo]::GetVersionInfo($Path).ProductVersion
    if (-not $product -or $product.Split('+')[0] -cne $Version) { throw "Runtime version $product does not match $Version." }
}

function New-VerifiedNuGetPackage {
    param([string]$Component)
    $projectName = if ($Component -eq 'Cli') { 'CLI' } else { 'McpServer' }
    $nuget = Join-Path $output 'nuget'
    Invoke-PackageCommand $root 'dotnet' @('pack', (Join-Path $root "src\ExcelMcp.$projectName\ExcelMcp.$projectName.csproj"),
        '-c', 'Release', "-p:Version=$Version", '-p:NuGetAudit=false', '-o', $nuget) | Out-Host
    $config = Join-Path $output 'nuget.config'
    $xml = [xml]'<configuration><packageSources><clear/><add key="built" value=""/></packageSources></configuration>'
    $xml.configuration.packageSources.add.SetAttribute('value', $nuget)
    $xml.Save($config)
    $installed = Join-Path $output "installed-$Component"
    Invoke-PackageCommand $root 'dotnet' @('tool', 'install', "Sbroenne.ExcelMcp.$projectName", '--version', $Version,
        '--tool-path', $installed, '--configfile', $config, '--no-cache') | Out-Host
    $executable = if ($Component -eq 'Cli') { 'excelcli.exe' } else { 'mcp-excel.exe' }
    Invoke-PackageCommand $root (Join-Path $installed $executable) @('--version') | Out-Host
}

$extensionStage = $null
$primary = $null
try {
    if ($needsSkills -and -not $SkillsDirectory) {
        $SkillsDirectory = Join-Path $output 'generated-skills'
        & (Join-Path $PSScriptRoot 'Build-AgentSkills.ps1') -Version $Version -OutputDir $SkillsDirectory -GenerateOnly | Out-Host
    }
    if ('Extension' -in $Components) {
        $extensionStage = Initialize-ExtensionPackage -RepoRoot $root -Version $Version -SkillsDirectory $SkillsDirectory -SkipExtensionTests:$SkipExtensionTests
    }
    $prepared = @{}
    $runtimes = @()
    if ('Cli' -in $Components) { $runtimes += 'Cli' }
    if ('Mcp' -in $Components -or 'Extension' -in $Components) { $runtimes += 'Mcp' }
    foreach ($component in $runtimes) {
        if ($component -in $Components) { New-VerifiedNuGetPackage $component }
        $supplied = if ($component -eq 'Cli') { $CliRuntimeExecutable } else { $McpRuntimeExecutable }
        $executable = if ($supplied) { [IO.Path]::GetFullPath($supplied, $root) } else { Publish-ReleaseRuntime $component 'x64' }
        Assert-ReleaseRuntime $executable 'x64'
        $prepared[$component] = $executable
        Invoke-PackageCommand $root $executable @('--version') | Out-Host
        if ($component -notin $Components) { continue }
        foreach ($architecture in 'x64', 'arm64') {
            $runtime = $executable
            if ($architecture -eq 'arm64') {
                $runtime = Publish-ReleaseRuntime $component $architecture
                Assert-ReleaseRuntime $runtime $architecture
                $prepared["$component-arm64"] = $runtime
            }
            $npmComponent = if ($component -eq 'Cli') { 'Cli' } else { 'McpServer' }
            $archives = & (Join-Path $PSScriptRoot 'Build-NpmPackages.ps1') -Component $npmComponent -Architecture $architecture -Version $Version -RuntimeExecutable $runtime -OutputDirectory (Join-Path $output 'npm') | ConvertFrom-Json
            & (Join-Path $PSScriptRoot 'Test-NpmPackages.ps1') -Component $npmComponent -Architecture $architecture -LauncherPackage $archives.LauncherPackage -RuntimePackage $archives.RuntimePackage -ArchiveOnly:($architecture -eq 'arm64')
            New-StandalonePackage $component $Version $runtime $architecture $output $root | Out-Host
        }
    }
    if ('Mcpb' -in $Components) { & (Join-Path $root 'mcpb\Build-McpBundle.ps1') -Version $Version -OutputDir (Join-Path $output 'mcpb') | Out-Host }
    if ('Plugins' -in $Components) {
        $plugins = Join-Path $output 'plugins'
        & (Join-Path $PSScriptRoot 'Build-Plugins.ps1') -Version $Version -OutputDir $plugins -SkillsDirectory $SkillsDirectory | Out-Host
        [IO.Compression.ZipFile]::CreateFromDirectory($plugins, (Join-Path $output "excel-plugins-v$Version.zip"))
    }
    if ($extensionStage) {
        if (-not $prepared.ContainsKey('Mcp-arm64')) {
            $prepared['Mcp-arm64'] = Publish-ReleaseRuntime 'Mcp' 'arm64'
            Assert-ReleaseRuntime $prepared['Mcp-arm64'] 'arm64'
        }
        Complete-ExtensionPackages $extensionStage $Version $output $prepared
    }
    Write-Host "Verified packages: $($Components -join ', '). Output: $output"
} catch { $primary = $_; throw }
finally {
    if ($extensionStage) {
        try { Remove-PackageStaging $extensionStage }
        catch {
            if ($primary) { throw [AggregateException]::new('Packaging and extension cleanup failed.', [Exception[]]@($primary.Exception, $_.Exception)) }
            throw
        }
    }
}
