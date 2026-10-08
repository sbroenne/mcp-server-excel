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
    [string]$RuntimeExecutable,

    [Parameter()]
    [string]$OutputDir = "./artifacts"
)

$ErrorActionPreference = "Stop"
Set-StrictMode -Version Latest
. (Join-Path $PSScriptRoot "McpbPackaging.ps1")

$McpbDir = $PSScriptRoot
$RootDir = Split-Path $McpbDir -Parent
. (Join-Path $RootDir "scripts/PackageHelpers.ps1")
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
        DisplayName = "Excel (Apple Silicon macOS - experimental beta)"
        LongDescription = "Experimental beta for Apple Silicon macOS, not Windows feature parity. The capability-gated backend enables verified session, worksheet, basic range, named-range, calculation, Goal Seek, Data Table, and Python formula-write actions. Power Query and VBA are unavailable. Data Model/DAX/OLAP, Tables, PivotTables, charts, slicers, connections, QueryTables, XML Maps, screenshots, advanced visual formatting, and Python result reads are also unsupported. Unavailable actions fail explicitly; optional bridge installation does not enable them. Requires Excel for Mac 16.112 or later. Failed or cancelled mutations can partly apply; reconcile the surviving session before retrying. See https://github.com/sbroenne/mcp-server-excel/blob/main/specs/MACOS-SUPPORT.md#not-supported-in-the-macos-beta."
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

$OutputDir = if ([IO.Path]::IsPathRooted($OutputDir)) {
    [IO.Path]::GetFullPath($OutputDir)
}
else {
    [IO.Path]::GetFullPath((Join-Path $McpbDir $OutputDir))
}
Assert-PackageOutputPath -RepoRoot $RootDir -Path $OutputDir -Inputs @($RuntimeExecutable)
New-Item -ItemType Directory -Path $OutputDir -Force | Out-Null

if ($RuntimeIdentifier -eq "win-x64") {
    $MetadataStage = Join-Path $OutputDir "staging-windows"
    $StagedMcpbPath = Join-Path $OutputDir "staging-windows.mcpb"
    if (Test-Path -LiteralPath $MetadataStage) {
        Remove-McpbStagingDirectory -Path $MetadataStage
    }
    if (Test-Path -LiteralPath $StagedMcpbPath) {
        Remove-Item -LiteralPath $StagedMcpbPath -Force
    }
    New-Item -ItemType Directory -Path $MetadataStage | Out-Null
    try {
        $MetadataManifest = Get-Content (Join-Path $McpbDir "manifest.json") -Raw | ConvertFrom-Json
        $MetadataManifest.version = $Version
        $MetadataManifest.display_name = "Excel (Windows)"
        $MetadataManifest.compatibility.platforms = @("win32")
        $MetadataManifest | ConvertTo-Json -Depth 20 |
            Set-Content (Join-Path $MetadataStage "manifest.json") -Encoding utf8
        foreach ($name in @("icon-512.png", "README.md")) {
            Copy-Item (Join-Path $McpbDir $name) $MetadataStage
        }
        foreach ($name in @("LICENSE", "CHANGELOG.md")) {
            Copy-Item (Join-Path $RootDir $name) $MetadataStage
        }
        $McpbPath = Join-Path $OutputDir "excel-mcp-$Version.mcpb"
        New-McpbArchive -SourceDirectory $MetadataStage -DestinationPath $StagedMcpbPath
        Install-PackageOutput -Source $StagedMcpbPath -Destination $McpbPath
        Copy-Item (Join-Path $MetadataStage "manifest.json") (Join-Path $OutputDir "manifest.json") -Force
        Write-Output $McpbPath
        return
    }
    finally {
        if (Test-Path -LiteralPath $StagedMcpbPath) {
            Remove-Item -LiteralPath $StagedMcpbPath -Force
        }
        Remove-McpbStagingDirectory -Path $MetadataStage
    }
}

$StagingDir = Join-Path $OutputDir "staging-$($Target.Slug)"
if (Test-Path $StagingDir) {
    Remove-McpbStagingDirectory -Path $StagingDir
}
$PublishDir = Join-Path $StagingDir "publish"
$ServerDir = Join-Path $StagingDir "server"
New-Item -ItemType Directory -Path $PublishDir -Force | Out-Null
New-Item -ItemType Directory -Path $ServerDir -Force | Out-Null

Write-Host "Building MCPB package $($Target.Slug) version $Version..." -ForegroundColor Cyan

if ($RuntimeExecutable) {
    $resolvedRuntime = (Resolve-Path -LiteralPath $RuntimeExecutable).Path
    Copy-Item -LiteralPath $resolvedRuntime -Destination (Join-Path $PublishDir $Target.SourceExecutable)
}
else {
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

    & (Join-Path $RootDir "scripts/Sign-MacBinary.ps1") -Path $FinalExecutable -AutomationClient
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
if (Test-Path -LiteralPath $McpbPath) {
    Remove-Item -LiteralPath $McpbPath -Force
}
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
