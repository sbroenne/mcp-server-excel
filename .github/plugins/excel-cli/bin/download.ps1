[CmdletBinding()]
param(
    [switch]$Force,
    [switch]$PassThru,
    [switch]$Quiet
)

$ErrorActionPreference = "Stop"
Set-StrictMode -Version Latest

$PluginName = "excel-cli"
$ExecutableName = "excelcli.exe"
$RepoOwner = "sbroenne"
$RepoName = "mcp-server-excel"
$ReleaseApiUrl = "https://api.github.com/repos/$RepoOwner/$RepoName/releases/latest"
$ReleasePageUrl = "https://github.com/$RepoOwner/$RepoName/releases/latest"
$CacheRoot = Join-Path $env:USERPROFILE ".copilot\plugin-runtime\mcp-server-excel\$PluginName"
$DownloadsDir = Join-Path $CacheRoot "downloads"
$ReleasesDir = Join-Path $CacheRoot "releases"
$StatePath = Join-Path $CacheRoot "bootstrap-state.json"
$SessionId = if ([string]::IsNullOrWhiteSpace($env:COPILOT_AGENT_SESSION_ID)) { "standalone" } else { $env:COPILOT_AGENT_SESSION_ID }

function Write-Status {
    param(
        [string]$Message,

        [string]$Color = "Gray"
    )

    if (-not $Quiet) {
        Write-Host $Message -ForegroundColor $Color
    }
}

function Ensure-Directory {
    param([Parameter(Mandatory = $true)][string]$Path)

    if (-not (Test-Path $Path)) {
        New-Item -ItemType Directory -Path $Path -Force | Out-Null
    }
}

function New-State {
    return [pscustomobject]@{
        checkedSessionId = $null
        checkedAtUtc = $null
        latestTag = $null
        latestVersion = $null
        assetName = $null
        assetUrl = $null
        cachedReleaseTag = $null
        binaryPath = $null
    }
}

function Get-State {
    if (-not (Test-Path $StatePath)) {
        return New-State
    }

    try {
        $loadedState = Get-Content $StatePath -Raw | ConvertFrom-Json
        foreach ($name in @("checkedSessionId", "checkedAtUtc", "latestTag", "latestVersion", "assetName", "assetUrl", "cachedReleaseTag", "binaryPath")) {
            if ($null -eq $loadedState.PSObject.Properties[$name]) {
                $loadedState | Add-Member -MemberType NoteProperty -Name $name -Value $null
            }
        }

        return $loadedState
    } catch {
        Write-Status "[excel-cli] Ignoring unreadable bootstrap state and starting fresh." "DarkYellow"
        return New-State
    }
}

function Save-State {
    param([Parameter(Mandatory = $true)]$State)

    Ensure-Directory -Path $CacheRoot
    $json = $State | ConvertTo-Json -Depth 6
    [System.IO.File]::WriteAllText($StatePath, "$json`n", [System.Text.UTF8Encoding]::new($false))
}

function Get-LatestReleaseMetadata {
    Write-Status "[excel-cli] Checking latest GitHub release..." "Cyan"

    try {
        $headers = @{
            Accept       = "application/vnd.github+json"
            "User-Agent" = "excel-cli-plugin-bootstrap"
        }

        $release = Invoke-RestMethod -Uri $ReleaseApiUrl -Headers $headers
        $releaseVersion = $release.tag_name -replace '^v', ''
        $assetName = "ExcelMcp-CLI-$releaseVersion-windows.zip"
        $asset = $release.assets | Where-Object { $_.name -eq $assetName } | Select-Object -First 1

        if ($null -eq $asset) {
            throw "Latest release '$($release.tag_name)' does not contain asset '$assetName'."
        }

        return [pscustomobject]@{
            Tag = $release.tag_name
            Version = $releaseVersion
            AssetName = $asset.name
            AssetUrl = $asset.browser_download_url
        }
    } catch {
        throw "Failed to resolve the latest excelcli release. $_`nRelease page: $ReleasePageUrl"
    }
}

function Resolve-BinaryPath {
    param([Parameter(Mandatory = $true)]$State)

    if (-not [string]::IsNullOrWhiteSpace($State.binaryPath) -and (Test-Path $State.binaryPath)) {
        return $State.binaryPath
    }

    if ([string]::IsNullOrWhiteSpace($State.latestVersion)) {
        return $null
    }

    $releaseDir = Join-Path $ReleasesDir $State.latestVersion
    if (-not (Test-Path $releaseDir)) {
        return $null
    }

    $binary = Get-ChildItem -Path $releaseDir -Recurse -File -Filter $ExecutableName | Select-Object -First 1
    if ($null -eq $binary) {
        return $null
    }

    return $binary.FullName
}

function Ensure-LatestRuntime {
    param([Parameter(Mandatory = $true)]$State)

    $downloadZipPath = Join-Path $DownloadsDir $State.assetName
    $releaseDir = Join-Path $ReleasesDir $State.latestVersion

    Ensure-Directory -Path $DownloadsDir
    Ensure-Directory -Path $ReleasesDir

    $downloadRequired = $Force -or -not (Test-Path $downloadZipPath) -or $State.cachedReleaseTag -ne $State.latestTag
    if ($downloadRequired) {
        Write-Status "[excel-cli] Downloading $($State.assetName)..." "Yellow"
        Invoke-WebRequest -Uri $State.assetUrl -OutFile $downloadZipPath
    } else {
        Write-Status "[excel-cli] Reusing cached package $($State.assetName)." "DarkGray"
    }

    $binaryPath = Resolve-BinaryPath -State $State
    if (-not $Force -and $State.cachedReleaseTag -eq $State.latestTag -and -not [string]::IsNullOrWhiteSpace($binaryPath) -and (Test-Path $binaryPath)) {
        return $binaryPath
    }

    if (Test-Path $releaseDir) {
        Remove-Item -Path $releaseDir -Recurse -Force
    }

    Ensure-Directory -Path $releaseDir
    Write-Status "[excel-cli] Extracting $($State.assetName)..." "Yellow"
    Expand-Archive -Path $downloadZipPath -DestinationPath $releaseDir -Force

    $binary = Get-ChildItem -Path $releaseDir -Recurse -File -Filter $ExecutableName | Select-Object -First 1
    if ($null -eq $binary) {
        throw "Downloaded package '$($State.assetName)' did not contain $ExecutableName."
    }

    $State.cachedReleaseTag = $State.latestTag
    $State.binaryPath = $binary.FullName
    Save-State -State $State

    return $binary.FullName
}

if ($env:OS -ne "Windows_NT") {
    throw "excel-cli plugin bootstrap is Windows-only."
}

$state = Get-State
$sessionNeedsFreshnessCheck = $Force -or [string]::IsNullOrWhiteSpace($state.checkedSessionId) -or $state.checkedSessionId -ne $SessionId

if ($sessionNeedsFreshnessCheck) {
    $latest = Get-LatestReleaseMetadata
    $state.checkedSessionId = $SessionId
    $state.checkedAtUtc = [DateTime]::UtcNow.ToString("o")
    $state.latestTag = $latest.Tag
    $state.latestVersion = $latest.Version
    $state.assetName = $latest.AssetName
    $state.assetUrl = $latest.AssetUrl
    Save-State -State $state
} else {
    Write-Status "[excel-cli] Freshness already checked for this Copilot session." "DarkGray"
}

$binaryPath = Ensure-LatestRuntime -State $state
$state.binaryPath = $binaryPath
Save-State -State $state

if ($PassThru) {
    Write-Output $binaryPath
    return
}

$binaryInfo = Get-Item $binaryPath
Write-Status
Write-Status "✅ excelcli runtime ready." "Green"
Write-Status "   Release: $($state.latestTag)" "Gray"
Write-Status "   Binary:  $binaryPath" "Gray"
Write-Status "   Size:    $([math]::Round($binaryInfo.Length / 1MB, 2)) MB" "Gray"
