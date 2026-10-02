[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [ValidatePattern('^\d+\.\d+\.\d+(?:-[0-9A-Za-z.-]+)?$')]
    [string]$Version,

    [string]$OutputDirectory,

    [string]$SkillsDirectory,

    [switch]$Notarize
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

$root = Split-Path $PSScriptRoot -Parent
. (Join-Path $PSScriptRoot 'PackageHelpers.ps1')

if ($IsWindows -or [Runtime.InteropServices.RuntimeInformation]::OSArchitecture -ne [Runtime.InteropServices.Architecture]::Arm64) {
    throw 'Apple Silicon macOS is required to build Mac release packages.'
}

if (-not $OutputDirectory) {
    $OutputDirectory = Join-Path $root "artifacts/packages-macos/$([Guid]::NewGuid().ToString('N'))"
}
$OutputDirectory = [IO.Path]::GetFullPath($OutputDirectory, $root)
Assert-PackageOutputPath -Path $OutputDirectory -RepoRoot $root -Inputs @($SkillsDirectory)
if (Test-Path -LiteralPath $OutputDirectory) {
    throw "Use a new package output directory: $OutputDirectory"
}
New-Item -ItemType Directory -Path $OutputDirectory -Force | Out-Null

$runtimeRoot = Join-Path $OutputDirectory 'runtimes'
$npmOutput = Join-Path $OutputDirectory 'npm'
$runtimeIdentifier = 'osx-arm64'
$targets = @{
    Cli = @{
        Component = 'Cli'
        SourceExecutable = 'excelcli'
        PackageExecutable = 'excelcli'
        NpmComponent = 'Cli'
        NpmName = 'excelcli'
        Readme = 'src/ExcelMcp.CLI/README.md'
        ZipName = "ExcelMcp-CLI-$Version-macos-arm64.zip"
    }
    Mcp = @{
        Component = 'Mcp'
        SourceExecutable = 'Sbroenne.ExcelMcp.McpServer'
        PackageExecutable = 'mcp-excel'
        NpmComponent = 'McpServer'
        NpmName = 'mcp-server-excel'
        Readme = 'src/ExcelMcp.McpServer/README.md'
        ZipName = "ExcelMcp-MCP-Server-$Version-macos-arm64.zip"
    }
}

foreach ($name in @('Cli', 'Mcp')) {
    $target = $targets[$name]
    $runtimeDirectory = Join-Path $runtimeRoot $name
    Publish-PackageRuntime `
        -Component $target.Component `
        -RepoRoot $root `
        -Version $Version `
        -OutputDirectory $runtimeDirectory `
        -RuntimeIdentifier $runtimeIdentifier

    $runtimeExecutable = Join-Path $runtimeDirectory $target.SourceExecutable
    $helper = & (Join-Path $PSScriptRoot 'Build-MacScreenCaptureHelper.ps1') `
        -RuntimeIdentifier $runtimeIdentifier `
        -OutputRoot (Join-Path $runtimeDirectory 'native')
    $helperDirectory = Join-Path $runtimeDirectory 'helpers'
    New-Item -ItemType Directory -Path $helperDirectory -Force | Out-Null
    Copy-Item -LiteralPath $helper -Destination (Join-Path $helperDirectory 'excelmcp-screencapture')
    Remove-Item -LiteralPath (Join-Path $runtimeDirectory 'native') -Recurse -Force

    & /bin/chmod +x $runtimeExecutable (Join-Path $helperDirectory 'excelmcp-screencapture')
    & (Join-Path $PSScriptRoot 'Sign-MacBinary.ps1') -Path (Join-Path $helperDirectory 'excelmcp-screencapture')
    & (Join-Path $PSScriptRoot 'Sign-MacBinary.ps1') -Path $runtimeExecutable -AutomationClient
    & $runtimeExecutable --version
    if ($LASTEXITCODE -ne 0) {
        throw "$name runtime validation failed with exit code $LASTEXITCODE."
    }

    & (Join-Path $PSScriptRoot 'Build-NpmPackages.ps1') `
        -Component $target.NpmComponent `
        -Version $Version `
        -RuntimeIdentifier $runtimeIdentifier `
        -RuntimeExecutable $runtimeExecutable `
        -OutputDirectory $npmOutput
    & (Join-Path $PSScriptRoot 'Test-NpmPackages.ps1') `
        -Component $target.NpmComponent `
        -LauncherPackage (Join-Path $npmOutput "sbroenne-$($target.NpmName)-$Version.tgz") `
        -RuntimePackage (Join-Path $npmOutput "sbroenne-$($target.NpmName)-darwin-arm64-$Version.tgz")

    $zipStage = Join-Path $OutputDirectory "zip-$name"
    New-Item -ItemType Directory -Path $zipStage -Force | Out-Null
    Copy-Item -LiteralPath $runtimeExecutable -Destination (Join-Path $zipStage $target.PackageExecutable)
    Copy-Item -LiteralPath $helperDirectory -Destination $zipStage -Recurse
    Copy-Item -LiteralPath (Join-Path $root $target.Readme) -Destination (Join-Path $zipStage 'README.md')
    Copy-Item -LiteralPath (Join-Path $root 'LICENSE') -Destination $zipStage
    Copy-Item -LiteralPath (Join-Path $root 'CHANGELOG.md') -Destination $zipStage
    Push-Location $zipStage
    try {
        & /usr/bin/zip -q -r (Join-Path $OutputDirectory $target.ZipName) .
        if ($LASTEXITCODE -ne 0) {
            throw "$name ZIP creation failed with exit code $LASTEXITCODE."
        }
    }
    finally {
        Pop-Location
    }

    & (Join-Path $PSScriptRoot 'Test-DistributionPackages.ps1') `
        -ArchivePath (Join-Path $OutputDirectory $target.ZipName) `
        -ExecutableRelativePath $target.PackageExecutable `
        -MacHelperRelativePath 'helpers/excelmcp-screencapture' `
        -ExpectedArchitecture 'macos-arm64' `
        -Launch
}

if (-not $SkillsDirectory) {
    $SkillsDirectory = Join-Path $OutputDirectory 'generated-skills'
    & (Join-Path $PSScriptRoot 'Build-AgentSkills.ps1') `
        -GenerateOnly `
        -Version $Version `
        -OutputDir $SkillsDirectory
}

$extensionStage = Join-Path ([IO.Path]::GetTempPath()) "ExcelMcpMacExtension-$([Guid]::NewGuid().ToString('N'))"
try {
    New-Item -ItemType Directory -Path $extensionStage -Force | Out-Null
    Get-ChildItem (Join-Path $root 'vscode-extension') -Force |
        Where-Object { $_.Name -notin @('node_modules', 'bin', 'out', 'skills') -and $_.Extension -ne '.vsix' } |
        Copy-Item -Destination $extensionStage -Recurse

    $extensionRuntime = Join-Path $extensionStage 'bin/darwin-arm64'
    New-Item -ItemType Directory -Path $extensionRuntime -Force | Out-Null
    Copy-Item -LiteralPath (Join-Path $runtimeRoot 'Mcp/Sbroenne.ExcelMcp.McpServer') -Destination $extensionRuntime
    Copy-Item -LiteralPath (Join-Path $runtimeRoot 'Mcp/helpers') -Destination $extensionRuntime -Recurse
    $extensionSkills = Join-Path $extensionStage 'skills'
    New-Item -ItemType Directory -Path $extensionSkills -Force | Out-Null
    Copy-Item (Join-Path $SkillsDirectory 'excel-mcp') $extensionSkills -Recurse
    Set-Content (Join-Path $extensionSkills 'excel-mcp/VERSION') $Version -NoNewline
    Copy-Item (Join-Path $root 'CHANGELOG.md') $extensionStage -Force

    $manifestPath = Join-Path $extensionStage 'package.json'
    $manifest = Get-Content $manifestPath -Raw | ConvertFrom-Json
    $manifest.version = $Version
    $manifest.scripts.'vscode:prepublish' = 'npm run compile'
    $manifest | ConvertTo-Json -Depth 20 | Set-Content $manifestPath -Encoding utf8

    $vsixPath = Join-Path $OutputDirectory "excelmcp-$Version-darwin-arm64.vsix"
    Push-Location $extensionStage
    try {
        npm ci --ignore-scripts
        if ($LASTEXITCODE -ne 0) { throw 'Extension dependency installation failed.' }
        npm run lint
        if ($LASTEXITCODE -ne 0) { throw 'Extension lint failed.' }
        npm exec -- vsce package --target darwin-arm64 --no-dependencies --out $vsixPath
        if ($LASTEXITCODE -ne 0) { throw 'Extension packaging failed.' }
    }
    finally {
        Pop-Location
    }

    & (Join-Path $PSScriptRoot 'Test-DistributionPackages.ps1') `
        -ArchivePath $vsixPath `
        -ExecutableRelativePath 'extension/bin/darwin-arm64/Sbroenne.ExcelMcp.McpServer' `
        -MacHelperRelativePath 'extension/bin/darwin-arm64/helpers/excelmcp-screencapture' `
        -ExpectedArchitecture 'macos-arm64' `
        -Launch
}
finally {
    if (Test-Path -LiteralPath $extensionStage) {
        Remove-Item -LiteralPath $extensionStage -Recurse -Force
    }
}

$mcpbOutput = Join-Path $OutputDirectory 'mcpb'
& (Join-Path $root 'mcpb/Build-McpBundle.ps1') `
    -Version $Version `
    -RuntimeIdentifier $runtimeIdentifier `
    -RuntimeExecutable (Join-Path $runtimeRoot 'Mcp/Sbroenne.ExcelMcp.McpServer') `
    -OutputDir $mcpbOutput
$mcpbPath = Join-Path $mcpbOutput "excel-mcp-$Version-macos-arm64.mcpb"
& (Join-Path $PSScriptRoot 'Test-DistributionPackages.ps1') `
    -ArchivePath $mcpbPath `
    -ExecutableRelativePath 'server/excel-mcp-server' `
    -MacHelperRelativePath 'server/helpers/excelmcp-screencapture' `
    -ExpectedArchitecture 'macos-arm64' `
    -Launch

if ($Notarize) {
    foreach ($archive in @(
        (Join-Path $OutputDirectory $targets.Cli.ZipName),
        (Join-Path $OutputDirectory $targets.Mcp.ZipName),
        (Join-Path $OutputDirectory "excelmcp-$Version-darwin-arm64.vsix"),
        $mcpbPath
    )) {
        & (Join-Path $PSScriptRoot 'Submit-MacNotarization.ps1') -ArchivePath $archive
    }
}

Write-Host "Apple Silicon release packages created in $OutputDirectory" -ForegroundColor Green
