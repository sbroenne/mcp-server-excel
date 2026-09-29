<#
.SYNOPSIS
    Creates one platform-specific MCPB package for Claude Desktop.

.DESCRIPTION
    Builds the MCP Server as a self-contained Windows x64 or Apple Silicon macOS
    executable and packages it as an .mcpb file for one-click installation.

.PARAMETER Version
    Package version. Defaults to the version in Directory.Build.props.

.PARAMETER RuntimeIdentifier
    Native runtime to package: win-x64 or osx-arm64.

.PARAMETER OutputDir
    Output directory relative to mcpb/. Defaults to ./artifacts.
#>

[CmdletBinding()]
param(
    [Parameter()]
    [string]$Version,

    [Parameter()]
    [ValidateSet("win-x64", "osx-arm64")]
    [string]$RuntimeIdentifier = "win-x64",

    [Parameter()]
    [string]$OutputDir = "./artifacts"
)

$ErrorActionPreference = "Stop"
Set-StrictMode -Version Latest
. (Join-Path $PSScriptRoot "McpbPackaging.ps1")

$McpbDir = $PSScriptRoot
$RootDir = Split-Path $McpbDir -Parent
$McpServerDir = Join-Path $RootDir "src/ExcelMcp.McpServer"
$Target = if ($RuntimeIdentifier -eq "win-x64") {
    @{
        Platform = "win32"
        Slug = "windows"
        SourceExecutable = "Sbroenne.ExcelMcp.McpServer.exe"
        BundleExecutable = "excel-mcp-server.exe"
        DisplayName = "Excel (Windows)"
        LongDescription = "Automate the real Microsoft Excel application from Claude on Windows. 31 specialized tools with 326 operations cover Power Query, DAX and the Data Model, VBA, PivotTables, Charts, Conditional Formatting, and more through Excel's COM API. Requires Windows x64 and Microsoft Excel 2016 or later."
    }
}
else {
    @{
        Platform = "darwin"
        Slug = "macos-arm64"
        SourceExecutable = "Sbroenne.ExcelMcp.McpServer"
        BundleExecutable = "excel-mcp-server"
        DisplayName = "Excel (Apple Silicon macOS)"
        LongDescription = "Automate the real Microsoft Excel application from Claude on Apple Silicon macOS. The capability-gated backend supports session lifecycle; worksheet create, list, rename, and delete; range values, formulas, number formats, row and column sizing, clearing, and calculation; plus Power Query list and view for clean saved workbooks without a Data Model. Unavailable operations fail explicitly. Requires Excel for Mac 16.112 or later."
    }
}

if (-not $Version) {
    $PropsFile = Join-Path $RootDir "Directory.Build.props"
    if (Test-Path $PropsFile) {
        $xml = [xml](Get-Content $PropsFile)
        $Version = $xml.Project.PropertyGroup.Version | Where-Object { $_ } | Select-Object -First 1
    }
    if (-not $Version) {
        $Version = "1.0.0"
    }
}

$OutputDir = Join-Path $McpbDir $OutputDir
if (Test-Path $OutputDir) {
    Remove-Item -LiteralPath $OutputDir -Recurse -Force
}
New-Item -ItemType Directory -Path $OutputDir -Force | Out-Null

$StagingDir = Join-Path $OutputDir "staging"
$PublishDir = Join-Path $StagingDir "publish"
$ServerDir = Join-Path $StagingDir "server"
New-Item -ItemType Directory -Path $PublishDir -Force | Out-Null
New-Item -ItemType Directory -Path $ServerDir -Force | Out-Null

Write-Host "Building MCPB package $($Target.Slug) version $Version..." -ForegroundColor Cyan

$PublishArgs = @(
    "publish"
    "$McpServerDir/ExcelMcp.McpServer.csproj"
    "-c", "Release"
    "-r", $RuntimeIdentifier
    "--self-contained", "true"
    "-p:PublishSingleFile=true"
    "-p:IncludeNativeLibrariesForSelfExtract=true"
    "-p:PublishTrimmed=false"
    "-p:PublishReadyToRun=false"
    "-p:NuGetAudit=false"
    "-p:Version=$Version"
    "-o", $PublishDir
    "--verbosity", "quiet"
)

& dotnet @PublishArgs
if ($LASTEXITCODE -ne 0) {
    throw "MCP Server publish failed for $RuntimeIdentifier with exit code $LASTEXITCODE."
}

$FinalExecutable = Join-Path $ServerDir $Target.BundleExecutable
Move-Item (Join-Path $PublishDir $Target.SourceExecutable) $FinalExecutable -Force
Remove-Item -LiteralPath $PublishDir -Recurse -Force

if ($RuntimeIdentifier.StartsWith("osx-", [StringComparison]::Ordinal) -and -not $IsWindows) {
    $HelperBuildRoot = Join-Path $StagingDir "native"
    $HelperSource = & (Join-Path $RootDir "scripts/Build-MacScreenCaptureHelper.ps1") `
        -RuntimeIdentifier $RuntimeIdentifier `
        -OutputRoot $HelperBuildRoot
    if ($LASTEXITCODE -ne 0) {
        throw "Could not build the macOS ScreenCaptureKit helper."
    }
    $HelperDirectory = Join-Path $ServerDir "helpers"
    New-Item -ItemType Directory -Path $HelperDirectory -Force | Out-Null
    $FinalHelper = Join-Path $HelperDirectory "excelmcp-screencapture"
    Copy-Item -LiteralPath $HelperSource -Destination $FinalHelper
    Remove-Item -LiteralPath $HelperBuildRoot -Recurse -Force
    & /bin/chmod +x $FinalHelper
    & (Join-Path $RootDir "scripts/Sign-MacBinary.ps1") -Path $FinalHelper
    if ($LASTEXITCODE -ne 0) {
        throw "Could not sign the macOS ScreenCaptureKit helper."
    }

    & /bin/chmod +x $FinalExecutable
    if ($LASTEXITCODE -ne 0) {
        throw "Could not mark the macOS MCP Server executable."
    }

    & (Join-Path $RootDir "scripts/Sign-MacBinary.ps1") -Path $FinalExecutable
    if ($LASTEXITCODE -ne 0) {
        throw "Could not sign the macOS MCP Server executable."
    }
}

$CanRunTarget = ($RuntimeIdentifier -eq "win-x64" -and $IsWindows) -or
    ($RuntimeIdentifier -eq "osx-arm64" -and $IsMacOS -and
        [Runtime.InteropServices.RuntimeInformation]::OSArchitecture -eq
            [Runtime.InteropServices.Architecture]::Arm64)
if ($CanRunTarget) {
    $VersionOutput = & $FinalExecutable --version 2>&1
    if ($LASTEXITCODE -ne 0) {
        throw "Packaged executable verification failed with exit code $LASTEXITCODE."
    }
    Write-Host "Verified $VersionOutput" -ForegroundColor Green
}
else {
    Write-Host "Skipped executable launch verification on this build host." -ForegroundColor DarkGray
}

$Manifest = Get-Content (Join-Path $McpbDir "manifest.json") -Raw | ConvertFrom-Json
$Manifest.version = $Version
$Manifest.display_name = $Target.DisplayName
$Manifest.long_description = $Target.LongDescription
$Manifest.compatibility.platforms = @($Target.Platform)
$ManifestDst = Join-Path $StagingDir "manifest.json"
[IO.File]::WriteAllText(
    $ManifestDst,
    ($Manifest | ConvertTo-Json -Depth 20),
    [Text.UTF8Encoding]::new($false))

Copy-Item (Join-Path $McpbDir "icon-512.png") (Join-Path $StagingDir "icon-512.png")
Copy-Item (Join-Path $McpbDir "README.md") (Join-Path $StagingDir "README.md")
Copy-Item (Join-Path $RootDir "LICENSE") (Join-Path $StagingDir "LICENSE")
Copy-Item (Join-Path $RootDir "CHANGELOG.md") (Join-Path $StagingDir "CHANGELOG.md")

$McpbFileName = "excel-mcp-$Version-$($Target.Slug).mcpb"
$McpbPath = Join-Path $OutputDir $McpbFileName
$MacExecutableRelativePath = if ($RuntimeIdentifier.StartsWith("osx-", [StringComparison]::Ordinal)) {
    @("server/excel-mcp-server", "server/helpers/excelmcp-screencapture")
}
else {
    ""
}
New-McpbArchive `
    -SourceDirectory $StagingDir `
    -DestinationPath $McpbPath `
    -MacExecutableRelativePath $MacExecutableRelativePath

$Archive = [IO.Compression.ZipFile]::OpenRead($McpbPath)
try {
    $ExpectedEntry = "server/$($Target.BundleExecutable)"
    $ExecutableEntry = $Archive.GetEntry($ExpectedEntry)
    if ($null -eq $ExecutableEntry) {
        throw "MCPB package is missing $ExpectedEntry."
    }
    if ($RuntimeIdentifier.StartsWith("osx-", [StringComparison]::Ordinal) -and
        ($ExecutableEntry.ExternalAttributes -band 0x00400000) -eq 0) {
        throw "macOS MCPB executable does not have an executable mode."
    }
    if ($RuntimeIdentifier.StartsWith("osx-", [StringComparison]::Ordinal)) {
        $HelperEntry = $Archive.GetEntry("server/helpers/excelmcp-screencapture")
        if ($null -eq $HelperEntry -or ($HelperEntry.ExternalAttributes -band 0x00400000) -eq 0) {
            throw "macOS MCPB helper is missing or does not have an executable mode."
        }
    }
}
finally {
    $Archive.Dispose()
}

Copy-Item $ManifestDst (Join-Path $OutputDir "manifest-$($Target.Slug).json")
Remove-McpbStagingDirectory -Path $StagingDir

$McpbSize = (Get-Item $McpbPath).Length / 1MB
Write-Host "Created $McpbFileName ($([math]::Round($McpbSize, 1)) MB)." -ForegroundColor Green
Write-Output $McpbPath
