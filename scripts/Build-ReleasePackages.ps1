<#
.SYNOPSIS
    Creates and verifies selected distributable packages. Never publishes them.
.DESCRIPTION
    PR CI passes BaseRef and HeadRef to select affected packages. Explicit manual
    runs and releases default to all components. A selected failure stops the run.
#>
[CmdletBinding()]
param(
    [ValidateSet('Cli', 'Mcp', 'Extension', 'Mcpb', 'Skills', 'Plugins')]
    [string[]]$Components = @('Cli', 'Mcp', 'Extension', 'Mcpb', 'Skills', 'Plugins'),
    [string]$BaseRef,
    [string]$HeadRef = 'HEAD',
    [string]$Version,
    [string]$SkillsDirectory,
    [string]$McpRuntimeExecutable,
    [string]$CliRuntimeExecutable,
    [string]$OutputDirectory
)
$ErrorActionPreference = 'Stop'
$root = Split-Path $PSScriptRoot -Parent
. (Join-Path $PSScriptRoot 'Get-ValidationPlan.ps1')
. (Join-Path $PSScriptRoot 'PackageHelpers.ps1')
function Invoke-PackageStep {
    param([string]$Name, [scriptblock]$Action)
    Write-Host $Name -ForegroundColor Cyan
    $global:LASTEXITCODE = 0
    & $Action
    if ($LASTEXITCODE -ne 0) { throw "$Name failed with exit code $LASTEXITCODE." }
}
function Read-VsixEntry {
    param([IO.Compression.ZipArchive]$Archive, [string]$Name)
    $entry = $Archive.GetEntry($Name)
    if (-not $entry) { throw "VSIX is missing $Name." }
    $reader = [IO.StreamReader]::new($entry.Open())
    try { $reader.ReadToEnd() }
    finally { $reader.Dispose() }
}
Push-Location $root
$extensionStage = $null
try {
    if ($BaseRef) {
        $paths = @(git -c core.quotepath=false diff --name-only --no-renames "$BaseRef...$HeadRef")
        if ($LASTEXITCODE -ne 0) { throw 'Cannot determine package inputs from the requested Git revisions.' }
        $plan = Get-ValidationPlan -Paths $paths
        $plan.Reasons | ForEach-Object { Write-Host $_ }
        $Components = @($Components | Where-Object { $plan.$_ })
    }
    if ($Components.Count -eq 0) {
        Write-Host 'No distributable package inputs changed.'
        return
    }
    if (-not $IsWindows) { throw 'Package installation checks require Windows (not Excel).' }
    if (-not $Version) { $Version = (Get-Content package.json -Raw | ConvertFrom-Json).version }
    if ($Version -notmatch '^\d+\.\d+\.\d+(?:-[A-Za-z0-9.-]+)?$') { throw 'A valid package version is required.' }
    if ($SkillsDirectory -and @($Components | Where-Object { $_ -in @('Skills', 'Extension', 'Plugins') }).Count) {
        foreach ($name in @('excel-cli', 'excel-mcp')) {
            $stamp = Join-Path $SkillsDirectory "$name\VERSION"
            if (-not (Test-Path -LiteralPath $stamp -PathType Leaf) -or (Get-Content -LiteralPath $stamp -Raw).Trim() -ne $Version) {
                throw "Prepared $name skill must match package version $Version."
            }
        }
    }
    if (-not $OutputDirectory) {
        $OutputDirectory = Join-Path $root "artifacts\packages\$([Guid]::NewGuid().ToString('N'))"
    }
    $OutputDirectory = [IO.Path]::GetFullPath($OutputDirectory, $root)
    Assert-PackageOutputPath -Path $OutputDirectory -RepoRoot $root -Inputs @($SkillsDirectory, $McpRuntimeExecutable, $CliRuntimeExecutable)
    if (Test-Path -LiteralPath $OutputDirectory) { throw "Use a new package output directory: $OutputDirectory" }
    New-Item -ItemType Directory -Path $OutputDirectory -Force | Out-Null
    $runtimeRoot = Join-Path $OutputDirectory 'runtimes'
    $prepared = @{}
    $neededRuntimes = @()
    if ($Components -contains 'Cli') { $neededRuntimes += 'Cli' }
    if (@($Components | Where-Object { $_ -in @('Mcp', 'Extension', 'Mcpb') }).Count) { $neededRuntimes += 'Mcp' }
    foreach ($component in $neededRuntimes) {
        $projectName = if ($component -eq 'Cli') { 'CLI' } else { 'McpServer' }
        $project = Join-Path $root "src\ExcelMcp.$projectName\ExcelMcp.$projectName.csproj"
        if ($Components -contains $component) {
            Invoke-PackageStep "$component NuGet package" {
                dotnet pack $project -c Release "-p:Version=$Version" -p:NuGetAudit=false -o (Join-Path $OutputDirectory 'nuget')
            }
            $toolDir = Join-Path $OutputDirectory "installed-$component"
            $packageId = "Sbroenne.ExcelMcp.$projectName"
            $nugetConfig = Join-Path $OutputDirectory 'nuget.config'
            $localFeed = [Security.SecurityElement]::Escape((Join-Path $OutputDirectory 'nuget'))
            Set-Content $nugetConfig "<configuration><packageSources><clear/><add key=`"built`" value=`"$localFeed`"/></packageSources></configuration>"
            Invoke-PackageStep "$component installed NuGet tool" {
                dotnet tool install $packageId --version $Version --tool-path $toolDir `
                    --configfile $nugetConfig --no-cache
            }
            $command = if ($component -eq 'Cli') { 'excelcli.exe' } else { 'mcp-excel.exe' }
            Invoke-PackageStep "$component NuGet version" { & (Join-Path $toolDir $command) --version }
        }
        $supplied = if ($component -eq 'Cli') { $CliRuntimeExecutable } else { $McpRuntimeExecutable }
        if ($supplied) {
            $prepared[$component] = (Resolve-Path -LiteralPath $supplied).Path
        } else {
            $runtimeDir = Join-Path $runtimeRoot $component
            Invoke-PackageStep "$component standalone runtime" {
                Publish-PackageRuntime -Component $component -RepoRoot $root -Version $Version -OutputDirectory $runtimeDir
            }
            $exeName = if ($component -eq 'Cli') { 'excelcli.exe' } else { 'Sbroenne.ExcelMcp.McpServer.exe' }
            $prepared[$component] = Join-Path $runtimeDir $exeName
        }
        Assert-PackageRuntimeArchitecture -Path $prepared[$component] -Architecture x64
        $productVersion = [Diagnostics.FileVersionInfo]::GetVersionInfo($prepared[$component]).ProductVersion
        if (($productVersion -split '\+')[0] -ne $Version) { throw "$component runtime version $productVersion does not match $Version." }
        Invoke-PackageStep "$component runtime version" { & $prepared[$component] --version }
        if ($Components -notcontains $component) { continue }
        $npmComponent = if ($component -eq 'Cli') { 'Cli' } else { 'McpServer' }
        $packageName = if ($component -eq 'Cli') { 'excelcli' } else { 'mcp-server-excel' }
        $npmDir = Join-Path $OutputDirectory 'npm'
        foreach ($architecture in @('x64', 'arm64')) {
            $npmRuntime = $prepared[$component]
            if ($architecture -eq 'arm64') {
                $armRuntimeDir = Join-Path $runtimeRoot "$component-arm64"
                Invoke-PackageStep "$component ARM64 npm runtime" {
                    Publish-PackageRuntime -Component $component -RepoRoot $root -Version $Version `
                        -Architecture arm64 -OutputDirectory $armRuntimeDir
                }
                $exeName = if ($component -eq 'Cli') { 'excelcli.exe' } else { 'Sbroenne.ExcelMcp.McpServer.exe' }
                $npmRuntime = Join-Path $armRuntimeDir $exeName
                $armVersion = [Diagnostics.FileVersionInfo]::GetVersionInfo($npmRuntime).ProductVersion
                if (($armVersion -split '\+')[0] -ne $Version) { throw "$component ARM64 runtime version $armVersion does not match $Version." }
            }
            Invoke-PackageStep "$component $architecture npm packages" {
                & (Join-Path $PSScriptRoot 'Build-NpmPackages.ps1') -Component $npmComponent -Version $Version `
                    -Architecture $architecture -RuntimeExecutable $npmRuntime -OutputDirectory $npmDir
            }
            Invoke-PackageStep "$component $architecture installed npm package" {
                & (Join-Path $PSScriptRoot 'Test-NpmPackages.ps1') -Component $npmComponent -Architecture $architecture `
                    -LauncherPackage (Join-Path $npmDir "sbroenne-$packageName-$Version.tgz") `
                    -RuntimePackage (Join-Path $npmDir "sbroenne-$packageName-win32-$architecture-$Version.tgz")
            }
        }
        $zipStage = Join-Path $OutputDirectory "zip-$component"
        New-Item -ItemType Directory -Path $zipStage | Out-Null
        $zipExecutable = if ($component -eq 'Cli') { 'excelcli.exe' } else { 'mcp-excel.exe' }
        Copy-Item -LiteralPath $prepared[$component] -Destination (Join-Path $zipStage $zipExecutable)
        foreach ($file in @('README.md', 'LICENSE', 'CHANGELOG.md')) { Copy-Item $file $zipStage }
        $zipName = if ($component -eq 'Cli') { "ExcelMcp-CLI-$Version-windows.zip" } else { "ExcelMcp-MCP-Server-$Version-windows.zip" }
        Compress-Archive -Path (Join-Path $zipStage '*') -DestinationPath (Join-Path $OutputDirectory $zipName)
    }
    if ($Components -contains 'Mcpb') {
        Invoke-PackageStep 'Claude Desktop bundle' {
            & (Join-Path $root 'mcpb\Build-McpBundle.ps1') -Version $Version `
                -RuntimeExecutable $prepared.Mcp -OutputDir (Join-Path $OutputDirectory 'mcpb')
        }
    }
    if (-not $SkillsDirectory -and @($Components | Where-Object { $_ -in @('Skills', 'Extension', 'Plugins') }).Count) {
        $SkillsDirectory = Join-Path $OutputDirectory 'generated-skills'
        Invoke-PackageStep 'Complete generated skills' {
            & (Join-Path $PSScriptRoot 'Build-AgentSkills.ps1') -GenerateOnly -Version $Version -OutputDir $SkillsDirectory
        }
    }
    if ($Components -contains 'Skills') {
        Invoke-PackageStep 'Skill package' {
            & (Join-Path $PSScriptRoot 'Build-AgentSkills.ps1') -Version $Version -SkillsDirectory $SkillsDirectory -OutputDir (Join-Path $OutputDirectory 'skills')
        }
    }
    if ($Components -contains 'Plugins') {
        Invoke-PackageStep 'Plugin packages' {
            & (Join-Path $PSScriptRoot 'Build-Plugins.ps1') -Version $Version -SkillsDirectory $SkillsDirectory -OutputDir (Join-Path $OutputDirectory 'plugins')
        }
        Compress-Archive -Path (Join-Path $OutputDirectory 'plugins\*') `
            -DestinationPath (Join-Path $OutputDirectory "excel-plugins-v$Version.zip")
    }
    if ($Components -contains 'Extension') {
        $extensionStage = Join-Path ([IO.Path]::GetTempPath()) "ExcelMcpExtension-$([Guid]::NewGuid().ToString('N'))"
        $extension = $extensionStage
        New-Item -ItemType Directory -Path $extension | Out-Null
        Get-ChildItem (Join-Path $root 'vscode-extension') -Force |
            Where-Object { $_.Name -notin @('node_modules', 'bin', 'out', 'skills') -and $_.Extension -ne '.vsix' } |
            Copy-Item -Destination $extension -Recurse
        $bin = New-Item -ItemType Directory -Path (Join-Path $extension 'bin')
        Copy-Item -LiteralPath $prepared.Mcp -Destination $bin.FullName
        $skills = New-Item -ItemType Directory -Path (Join-Path $extension 'skills')
        Copy-Item (Join-Path $SkillsDirectory 'excel-mcp') $skills.FullName -Recurse
        Set-Content (Join-Path $skills.FullName 'excel-mcp\VERSION') $Version -NoNewline
        Copy-Item (Join-Path $root 'CHANGELOG.md') $extension -Force
        $manifestPath = Join-Path $extension 'package.json'
        $manifest = Get-Content $manifestPath -Raw | ConvertFrom-Json
        $manifest.version = $Version
        $manifest.scripts.'vscode:prepublish' = 'npm run compile'
        $manifest | ConvertTo-Json -Depth 20 | Set-Content $manifestPath -Encoding utf8
        Push-Location $extension
        $extensionPackages = @(
            @{ Target = 'win32-x64'; FileName = "excel-mcp-$Version.vsix" },
            @{ Target = 'win32-arm64'; FileName = "excel-mcp-$Version-win32-arm64.vsix" }
        )
        try {
            Invoke-PackageStep 'Extension dependencies' { npm.cmd ci --ignore-scripts }
            Invoke-PackageStep 'Extension compile and metadata' { npm.cmd run compile }
            Invoke-PackageStep 'Extension lint' { npm.cmd run lint }
            Invoke-PackageStep 'Extension test types' { npm.cmd run typecheck:tests }
            Invoke-PackageStep 'Extension tests' { npm.cmd test }
            foreach ($package in $extensionPackages) {
                Invoke-PackageStep "Extension package ($($package.Target))" {
                    npm.cmd exec -- vsce package --no-dependencies --target $package.Target --out (Join-Path $OutputDirectory $package.FileName)
                }
            }
        }
        finally { Pop-Location }
        foreach ($package in $extensionPackages) {
            $vsix = [IO.Compression.ZipFile]::OpenRead((Join-Path $OutputDirectory $package.FileName))
            try {
                foreach ($required in @('extension/bin/Sbroenne.ExcelMcp.McpServer.exe', 'extension/out/extension.js', 'extension/out/prerequisites.js')) {
                    if (-not $vsix.GetEntry($required)) { throw "VSIX is missing $required." }
                }
                foreach ($skillFile in Get-ChildItem (Join-Path $skills.FullName 'excel-mcp') -File -Recurse) {
                    $relative = [IO.Path]::GetRelativePath($extension, $skillFile.FullName).Replace('\', '/')
                    if (-not $vsix.GetEntry("extension/$relative")) { throw "VSIX is missing $relative." }
                }
                if ((Read-VsixEntry $vsix 'extension/skills/excel-mcp/VERSION').Trim() -ne $Version) {
                    throw 'VSIX skill version does not match the package.'
                }
                $packagedManifest = Read-VsixEntry $vsix 'extension/package.json' | ConvertFrom-Json
                if ($packagedManifest.version -ne $Version -or
                    ($packagedManifest.extensionKind -join ',') -ne 'ui' -or
                    ($packagedManifest.os -join ',') -ne 'win32') {
                    throw 'VSIX version or local Windows host metadata is incorrect.'
                }
                [xml]$metadata = Read-VsixEntry $vsix 'extension.vsixmanifest'
                if ($metadata.PackageManifest.Metadata.Identity.TargetPlatform -ne $package.Target) {
                    throw "VSIX target does not match $($package.Target)."
                }
                foreach ($entry in $vsix.Entries) {
                    if ($entry.FullName -match '^extension/(node_modules|tests|scripts|\.vitest|coverage|out/tests)/' -or
                        $entry.FullName -match '^extension/(vitest\.config\.|tsconfig(?:\.test)?\.json$|bin/excelcli)') {
                        throw "VSIX contains development files or the CLI: $($entry.FullName)"
                    }
                }
                Write-Host "Verified VSIX target: $($package.Target)"
            }
            finally { $vsix.Dispose() }
        }
        $debugDirectory = New-Item -ItemType Directory -Path (Join-Path $OutputDirectory 'extension')
        Get-ChildItem $extension -Force | Where-Object Name -ne 'node_modules' |
            Copy-Item -Destination $debugDirectory.FullName -Recurse
    }
    Write-Host "Verified packages: $($Components -join ', '). Output: $OutputDirectory"
}
finally {
    Pop-Location
    if ($extensionStage -and (Test-Path -LiteralPath $extensionStage)) {
        Remove-Item -LiteralPath $extensionStage -Recurse -Force
    }
}
