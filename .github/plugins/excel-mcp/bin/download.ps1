[CmdletBinding()]
param(
    [switch]$Force,
    [switch]$PassThru,
    [switch]$Quiet
)

$ErrorActionPreference = "Stop"
Set-StrictMode -Version Latest

$PluginName = "excel-mcp"
$ExecutableName = "mcp-excel.exe"
$RepoOwner = "sbroenne"
$RepoName = "mcp-server-excel"
$ReleaseApiUrl = "https://api.github.com/repos/$RepoOwner/$RepoName/releases/latest"
$ReleasePageUrl = "https://github.com/$RepoOwner/$RepoName/releases/latest"
$ChecksumAssetName = "SHA256SUMS"
$CacheRoot = if (-not [string]::IsNullOrWhiteSpace($env:PLUGIN_DATA)) {
    Join-Path $env:PLUGIN_DATA "runtime"
} else {
    Join-Path $env:USERPROFILE ".copilot\plugin-runtime\mcp-server-excel\$PluginName"
}
$DownloadsDir = Join-Path $CacheRoot "downloads"
$ReleasesDir = Join-Path $CacheRoot "releases"
$StatePath = Join-Path $CacheRoot "bootstrap-state.json"
$HasCopilotSession = -not [string]::IsNullOrWhiteSpace($env:COPILOT_AGENT_SESSION_ID)
$SessionId = if ($HasCopilotSession) { $env:COPILOT_AGENT_SESSION_ID } else { "standalone" }

# Outside a Copilot session the session id is the constant "standalone", so it always equals the
# previously recorded one and the freshness check would never fire again. PATH and shim installs
# would then be pinned forever to whatever they first downloaded. Fall back to elapsed time there.
$StandaloneRecheckHours = 24

try {
    Add-Type -AssemblyName System.IO.Compression.FileSystem -ErrorAction Stop
} catch {
    # Already loaded, or running on a host where the type is available without an explicit load.
}

function Write-Status {
    param(
        [string]$Message,

        [string]$Color = "Gray"
    )

    if (-not $Quiet) {
        Write-Host $Message -ForegroundColor $Color
    }
}

function Write-StatusError {
    param([string]$Message)

    # Always stderr. Stdout carries the -PassThru binary path, and for the MCP plugin it is the
    # MCP stdio transport, so diagnostics must never be written there. This is intentionally not
    # gated on -Quiet: -Quiet suppresses progress chatter, not warnings.
    [Console]::Error.WriteLine($Message)
}

function Ensure-Directory {
    param([Parameter(Mandatory = $true)][string]$Path)

    if (-not (Test-Path $Path)) {
        New-Item -ItemType Directory -Path $Path -Force | Out-Null
    }
}

function Test-ZipArchive {
    param([Parameter(Mandatory = $true)][string]$Path)

    if (-not (Test-Path $Path)) {
        return $false
    }

    # Presence is not integrity. An interrupted transfer leaves a truncated file that Test-Path
    # happily reports as a usable cached download, which then fails at extraction on every run.
    $archive = $null
    try {
        $archive = [System.IO.Compression.ZipFile]::OpenRead($Path)
        return $archive.Entries.Count -gt 0
    } catch {
        return $false
    } finally {
        if ($null -ne $archive) {
            $archive.Dispose()
        }
    }
}

function Test-FileLocked {
    param([Parameter(Mandatory = $true)][string]$Path)

    if (-not (Test-Path $Path)) {
        return $false
    }

    $stream = $null
    try {
        $stream = [System.IO.File]::Open($Path, [System.IO.FileMode]::Open, [System.IO.FileAccess]::ReadWrite, [System.IO.FileShare]::None)
        return $false
    } catch {
        return $true
    } finally {
        if ($null -ne $stream) {
            $stream.Dispose()
        }
    }
}

function Get-BinaryProductVersion {
    param([Parameter(Mandatory = $true)][string]$Path)

    # Read the version stamped into the file rather than running the binary with --version: this
    # is instant, works offline, and has no side effects. The runtime's own --version performs an
    # update check, which is precisely wrong inside a bootstrap that must survive being offline.
    try {
        $productVersion = (Get-Item $Path).VersionInfo.ProductVersion
    } catch {
        return $null
    }

    if ([string]::IsNullOrWhiteSpace($productVersion)) {
        return $null
    }

    # Release builds stamp SemVer 2 build metadata, for example "1.10.7+d526b22d0eda".
    return ($productVersion -split '\+', 2)[0].Trim()
}

function Test-BinaryMatchesVersion {
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [string]$ExpectedVersion
    )

    if ([string]::IsNullOrWhiteSpace($ExpectedVersion)) {
        return $true
    }

    $actualVersion = Get-BinaryProductVersion -Path $Path

    # Missing version metadata is not treated as a mismatch. An unidentifiable runtime is still
    # better than no runtime, and the install path already validated the package contents.
    if ([string]::IsNullOrWhiteSpace($actualVersion)) {
        return $true
    }

    return $actualVersion -eq $ExpectedVersion
}

function Get-RuntimeCacheMutex {
    # Local\ rather than Global\ deliberately: creating a global kernel object requires
    # SeCreateGlobalPrivilege, which standard users do not hold. The cache lives under the user
    # profile, so a per-logon-session lock is exactly the right scope.
    return [System.Threading.Mutex]::new($false, "Local\excelmcp-plugin-$PluginName")
}

function Test-FreshnessWindowElapsed {
    param([Parameter(Mandatory = $true)]$State)

    if ([string]::IsNullOrWhiteSpace($State.checkedAtUtc)) {
        return $true
    }

    try {
        $checkedAtUtc = [DateTime]::Parse(
            $State.checkedAtUtc,
            [System.Globalization.CultureInfo]::InvariantCulture,
            [System.Globalization.DateTimeStyles]::RoundtripKind)
    } catch {
        return $true
    }

    return ([DateTime]::UtcNow - $checkedAtUtc.ToUniversalTime()) -ge [TimeSpan]::FromHours($StandaloneRecheckHours)
}

function New-State {
    return [pscustomobject]@{
        checkedSessionId = $null
        checkedAtUtc = $null
        latestTag = $null
        latestVersion = $null
        assetName = $null
        assetUrl = $null
        expectedSha256 = $null
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
        foreach ($name in @("checkedSessionId", "checkedAtUtc", "latestTag", "latestVersion", "assetName", "assetUrl", "expectedSha256", "cachedReleaseTag", "binaryPath")) {
            if ($null -eq $loadedState.PSObject.Properties[$name]) {
                $loadedState | Add-Member -MemberType NoteProperty -Name $name -Value $null
            }
        }

        return $loadedState
    } catch {
        Write-Status "[excel-mcp] Ignoring unreadable bootstrap state and starting fresh." "DarkYellow"
        return New-State
    }
}

function Save-State {
    param([Parameter(Mandatory = $true)]$State)

    Ensure-Directory -Path $CacheRoot
    $json = $State | ConvertTo-Json -Depth 6
    [System.IO.File]::WriteAllText($StatePath, "$json`n", [System.Text.UTF8Encoding]::new($false))
}

function Get-ExpectedSha256 {
    param(
        [Parameter(Mandatory = $true)][string]$Uri,
        [Parameter(Mandatory = $true)][string]$AssetName
    )

    $manifestPath = Join-Path ([System.IO.Path]::GetTempPath()) ("excelmcp-" + [Guid]::NewGuid().ToString("N") + "-SHA256SUMS")

    try {
        try {
            Invoke-WebRequest -Uri $Uri -OutFile $manifestPath
        } catch {
            $exception = [System.InvalidOperationException]::new(
                "Failed to download checksum metadata '$ChecksumAssetName'. $($_.Exception.Message)",
                $_.Exception)
            $exception.Data["AllowCachedFallback"] = $true
            throw $exception
        }

        if (-not (Test-Path $manifestPath -PathType Leaf)) {
            throw "Checksum metadata '$ChecksumAssetName' was not downloaded."
        }

        $matchingHashes = @()
        $lineNumber = 0
        foreach ($line in @(Get-Content -Path $manifestPath)) {
            $lineNumber++
            if ([string]::IsNullOrWhiteSpace($line)) {
                continue
            }

            if ($line -notmatch '^([0-9A-Fa-f]{64})[ \t]+[*]?(.+)$') {
                throw "Checksum metadata '$ChecksumAssetName' is malformed at line $lineNumber."
            }

            if ($Matches[2] -eq $AssetName) {
                $matchingHashes += $Matches[1].ToLowerInvariant()
            }
        }

        if ($matchingHashes.Count -eq 0) {
            throw "Checksum metadata '$ChecksumAssetName' does not contain an entry for '$AssetName'."
        }

        if ($matchingHashes.Count -ne 1) {
            throw "Checksum metadata '$ChecksumAssetName' contains multiple entries for '$AssetName'."
        }

        return $matchingHashes[0]
    } finally {
        if (Test-Path $manifestPath) {
            Remove-Item -Path $manifestPath -Force -ErrorAction SilentlyContinue
        }
    }
}

function Get-LatestReleaseMetadata {
    Write-Status "[excel-mcp] Checking latest GitHub release..." "Cyan"

    try {
        $headers = @{
            Accept       = "application/vnd.github+json"
            "User-Agent" = "excel-mcp-plugin-bootstrap"
        }

        # Unauthenticated GitHub API access is 60 requests/hour per source IP. Behind corporate
        # NAT that budget is shared by everyone on the network and is routinely exhausted, which
        # is the most common cause of a bootstrap failing to reach the release metadata.
        $token = if (-not [string]::IsNullOrWhiteSpace($env:GITHUB_TOKEN)) { $env:GITHUB_TOKEN } elseif (-not [string]::IsNullOrWhiteSpace($env:GH_TOKEN)) { $env:GH_TOKEN } else { $null }
        if ($null -ne $token) {
            $headers["Authorization"] = "Bearer $token"
        }

        try {
            $release = Invoke-RestMethod -Uri $ReleaseApiUrl -Headers $headers
        } catch {
            $exception = [System.InvalidOperationException]::new(
                "Failed to resolve the latest ExcelMcp MCP server release. Failed to query the GitHub release API. $($_.Exception.Message)`nRelease page: $ReleasePageUrl",
                $_.Exception)
            $exception.Data["AllowCachedFallback"] = $true
            throw $exception
        }
        $releaseVersion = $release.tag_name -replace '^v', ''
        $assetName = "ExcelMcp-MCP-Server-$releaseVersion-windows.zip"
        $asset = $release.assets | Where-Object { $_.name -eq $assetName } | Select-Object -First 1

        if ($null -eq $asset) {
            throw "Latest release '$($release.tag_name)' does not contain asset '$assetName'."
        }

        $checksumAsset = $release.assets | Where-Object { $_.name -eq $ChecksumAssetName } | Select-Object -First 1
        if ($null -eq $checksumAsset) {
            throw "Latest release '$($release.tag_name)' does not contain checksum asset '$ChecksumAssetName'."
        }

        $expectedSha256 = Get-ExpectedSha256 -Uri $checksumAsset.browser_download_url -AssetName $asset.name

        return [pscustomobject]@{
            Tag = $release.tag_name
            Version = $releaseVersion
            AssetName = $asset.name
            AssetUrl = $asset.browser_download_url
            ExpectedSha256 = $expectedSha256
        }
    } catch {
        if ($_.Exception.Data.Contains("AllowCachedFallback")) {
            throw $_.Exception
        }

        throw "Failed to resolve the latest ExcelMcp MCP server release. $_`nRelease page: $ReleasePageUrl"
    }
}

function Find-ReleaseBinary {
    param(
        [string]$Version,
        [string]$ExpectedVersion
    )

    if ([string]::IsNullOrWhiteSpace($Version)) {
        return $null
    }

    $releaseDir = Join-Path $ReleasesDir $Version
    if (-not (Test-Path $releaseDir)) {
        return $null
    }

    $binary = Get-ChildItem -Path $releaseDir -Recurse -File -Filter $ExecutableName | Select-Object -First 1
    if ($null -eq $binary) {
        return $null
    }

    if (-not (Test-BinaryMatchesVersion -Path $binary.FullName -ExpectedVersion $ExpectedVersion)) {
        return $null
    }

    return $binary.FullName
}

function Resolve-BinaryPath {
    param(
        [Parameter(Mandatory = $true)]$State,
        [string]$ExpectedVersion
    )

    if (-not [string]::IsNullOrWhiteSpace($State.binaryPath) -and (Test-Path $State.binaryPath)) {
        if (Test-BinaryMatchesVersion -Path $State.binaryPath -ExpectedVersion $ExpectedVersion) {
            return $State.binaryPath
        }
    }

    return Find-ReleaseBinary -Version $State.latestVersion -ExpectedVersion $ExpectedVersion
}

function Save-RuntimeArchive {
    param(
        [Parameter(Mandatory = $true)][string]$Uri,
        [Parameter(Mandatory = $true)][string]$DestinationPath,
        [Parameter(Mandatory = $true)][string]$ExpectedSha256
    )

    # Download to a sibling temp file and rename into place, so an interrupted transfer can never
    # leave a truncated archive that a later run mistakes for a complete cached download.
    $partialPath = "$DestinationPath.$([Guid]::NewGuid().ToString('N')).part"

    try {
        Ensure-Directory -Path (Split-Path -Parent $DestinationPath)
        Invoke-WebRequest -Uri $Uri -OutFile $partialPath

        $actualSha256 = (Get-FileHash -Path $partialPath -Algorithm SHA256).Hash.ToLowerInvariant()
        if ($actualSha256 -ne $ExpectedSha256) {
            throw "SHA-256 mismatch for '$(Split-Path $DestinationPath -Leaf)': expected $ExpectedSha256, got $actualSha256."
        }

        if (-not (Test-ZipArchive -Path $partialPath)) {
            throw "Downloaded package '$(Split-Path $DestinationPath -Leaf)' is not a readable archive."
        }

        if (Test-Path $DestinationPath) {
            Remove-Item -Path $DestinationPath -Force
        }

        Move-Item -Path $partialPath -Destination $DestinationPath
    } finally {
        if (Test-Path $partialPath) {
            Remove-Item -Path $partialPath -Force -ErrorAction SilentlyContinue
        }
    }
}

function Test-RuntimeArchiveChecksum {
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [Parameter(Mandatory = $true)][string]$ExpectedSha256
    )

    if (-not (Test-Path $Path -PathType Leaf)) {
        return $false
    }

    $actualSha256 = (Get-FileHash -Path $Path -Algorithm SHA256).Hash.ToLowerInvariant()
    return $actualSha256 -eq $ExpectedSha256
}

function Install-RuntimeArchive {
    param(
        [Parameter(Mandatory = $true)][string]$ZipPath,
        [Parameter(Mandatory = $true)][string]$ReleaseDir
    )

    # Extract into a staging directory and swap it into place, so a failed or concurrent
    # extraction never leaves a half-populated release directory behind.
    $stagingDir = Join-Path $ReleasesDir (".staging-" + [Guid]::NewGuid().ToString("N"))

    try {
        Ensure-Directory -Path $stagingDir
        Expand-Archive -Path $ZipPath -DestinationPath $stagingDir -Force

        $staged = Get-ChildItem -Path $stagingDir -Recurse -File -Filter $ExecutableName | Select-Object -First 1
        if ($null -eq $staged) {
            throw "Downloaded package '$(Split-Path $ZipPath -Leaf)' did not contain $ExecutableName."
        }

        if (Test-Path $ReleaseDir) {
            $installed = Get-ChildItem -Path $ReleaseDir -Recurse -File -Filter $ExecutableName | Select-Object -First 1

            # Probe for a lock *before* deleting anything. Remove-Item -Recurse deletes every
            # unlocked file it reaches before failing on the locked executable, which destroys a
            # working install and leaves nothing usable behind.
            if ($null -ne $installed -and (Test-FileLocked -Path $installed.FullName)) {
                Write-StatusError "[excel-mcp] $ExecutableName is in use and cannot be replaced right now; keeping the existing install."
                return [pscustomobject]@{ Path = $installed.FullName; Installed = $false }
            }

            Remove-Item -Path $ReleaseDir -Recurse -Force
        }

        Ensure-Directory -Path (Split-Path -Parent $ReleaseDir)
        Move-Item -Path $stagingDir -Destination $ReleaseDir
        $stagingDir = $null

        $binary = Get-ChildItem -Path $ReleaseDir -Recurse -File -Filter $ExecutableName | Select-Object -First 1
        if ($null -eq $binary) {
            throw "Downloaded package '$(Split-Path $ZipPath -Leaf)' did not contain $ExecutableName."
        }

        return [pscustomobject]@{ Path = $binary.FullName; Installed = $true }
    } finally {
        if (-not [string]::IsNullOrWhiteSpace($stagingDir) -and (Test-Path $stagingDir)) {
            Remove-Item -Path $stagingDir -Recurse -Force -ErrorAction SilentlyContinue
        }
    }
}

function Ensure-LatestRuntime {
    param([Parameter(Mandatory = $true)]$State)

    Ensure-Directory -Path $DownloadsDir
    Ensure-Directory -Path $ReleasesDir

    # Fast path: an already-extracted runtime matching the resolved release is usable as-is.
    # This is checked before any download so that a warm cache never needs the network, even
    # when the cached .zip has been removed by a disk cleanup tool.
    if (-not $Force -and $State.cachedReleaseTag -eq $State.latestTag) {
        $cachedBinary = Resolve-BinaryPath -State $State -ExpectedVersion $State.latestVersion
        if (-not [string]::IsNullOrWhiteSpace($cachedBinary) -and (Test-Path $cachedBinary)) {
            return $cachedBinary
        }
    }

    if ([string]::IsNullOrWhiteSpace($State.assetName) -or
        [string]::IsNullOrWhiteSpace($State.assetUrl) -or
        [string]::IsNullOrWhiteSpace($State.expectedSha256)) {
        throw "excel-mcp must download the runtime but has no verified release metadata available. Check your network connection and try again.`nRelease page: $ReleasePageUrl"
    }

    $downloadZipPath = Join-Path $DownloadsDir $State.assetName
    $releaseDir = Join-Path $ReleasesDir $State.latestVersion

    # Serialize installs across concurrent sessions. Without this, two bootstraps race on the same
    # zip path and release directory, and each can observe the other's partially written files.
    $mutex = Get-RuntimeCacheMutex
    $acquired = $false

    try {
        try {
            $acquired = $mutex.WaitOne([TimeSpan]::FromMinutes(10))
        } catch [System.Threading.AbandonedMutexException] {
            # The previous holder died mid-install. We now own the lock and re-validate below.
            $acquired = $true
        }

        # Another session may have completed the install while this one waited for the lock. This
        # deliberately looks only inside the target release directory rather than trusting the
        # recorded binaryPath, which still points at the previous release during an upgrade.
        if (-not $Force) {
            $installedBinary = Find-ReleaseBinary -Version $State.latestVersion -ExpectedVersion $State.latestVersion
            if (-not [string]::IsNullOrWhiteSpace($installedBinary) -and (Test-Path $installedBinary)) {
                $State.cachedReleaseTag = $State.latestTag
                $State.binaryPath = $installedBinary
                Save-State -State $State
                return $installedBinary
            }
        }

        $lastError = $null

        for ($attempt = 1; $attempt -le 2; $attempt++) {
            try {
                # A readable ZIP is not necessarily the published ZIP. Verify cached bytes against
                # the checksum resolved from this exact release before any reuse or extraction.
                $downloadRequired = $Force -or
                    $attempt -gt 1 -or
                    $State.cachedReleaseTag -ne $State.latestTag -or
                    -not (Test-RuntimeArchiveChecksum -Path $downloadZipPath -ExpectedSha256 $State.expectedSha256) -or
                    -not (Test-ZipArchive -Path $downloadZipPath)

                if ($downloadRequired) {
                    Write-Status "[excel-mcp] Downloading $($State.assetName)..." "Yellow"
                    Save-RuntimeArchive -Uri $State.assetUrl -DestinationPath $downloadZipPath -ExpectedSha256 $State.expectedSha256
                } else {
                    Write-Status "[excel-mcp] Reusing cached package $($State.assetName)." "DarkGray"
                }

                Write-Status "[excel-mcp] Extracting $($State.assetName)..." "Yellow"
                $result = Install-RuntimeArchive -ZipPath $downloadZipPath -ReleaseDir $releaseDir

                # Only record the tag when the new runtime actually landed. If the install was
                # skipped because the executable was in use, the next run must try again.
                if ($result.Installed) {
                    $State.cachedReleaseTag = $State.latestTag
                }

                $State.binaryPath = $result.Path
                Save-State -State $State

                return $result.Path
            } catch {
                $lastError = $_

                # Discard the cached archive so the retry starts from a clean download rather than
                # repeating the same failure against the same bytes.
                if (Test-Path $downloadZipPath) {
                    Remove-Item -Path $downloadZipPath -Force -ErrorAction SilentlyContinue
                }

                if ($attempt -lt 2) {
                    Write-StatusError "[excel-mcp] Runtime install failed: $($_.Exception.Message)"
                    Write-StatusError "[excel-mcp] Retrying once with a fresh download."
                }
            }
        }

        throw $lastError
    } finally {
        if ($acquired) {
            $mutex.ReleaseMutex()
        }

        $mutex.Dispose()
    }
}

if ($env:OS -ne "Windows_NT") {
    throw "excel-mcp plugin bootstrap is Windows-only."
}

$state = Get-State
$sessionNeedsFreshnessCheck = $Force -or
    [string]::IsNullOrWhiteSpace($state.checkedSessionId) -or
    $state.checkedSessionId -ne $SessionId -or
    [string]::IsNullOrWhiteSpace($state.expectedSha256)

if (-not $sessionNeedsFreshnessCheck -and -not $HasCopilotSession) {
    $sessionNeedsFreshnessCheck = Test-FreshnessWindowElapsed -State $state
}

if ($sessionNeedsFreshnessCheck) {
    try {
        $latest = Get-LatestReleaseMetadata
        $state.checkedSessionId = $SessionId
        $state.checkedAtUtc = [DateTime]::UtcNow.ToString("o")
        $state.latestTag = $latest.Tag
        $state.latestVersion = $latest.Version
        $state.assetName = $latest.AssetName
        $state.assetUrl = $latest.AssetUrl
        $state.expectedSha256 = $latest.ExpectedSha256
        Save-State -State $state
    } catch {
        $allowCachedFallback = $_.Exception.Data.Contains("AllowCachedFallback") -and
            [bool]$_.Exception.Data["AllowCachedFallback"]
        if (-not $allowCachedFallback -or [string]::IsNullOrWhiteSpace($state.expectedSha256)) {
            throw
        }

        # A failed update check must not take down a working installation. If a usable runtime
        # is already cached, degrade to it instead of aborting.
        $cachedBinary = Resolve-BinaryPath -State $state
        if ([string]::IsNullOrWhiteSpace($cachedBinary) -or -not (Test-Path $cachedBinary)) {
            throw
        }

        Write-StatusError "[excel-mcp] Could not check for updates: $($_.Exception.Message)"
        Write-StatusError "[excel-mcp] Continuing with the cached runtime $($state.latestTag)."

        # Record the attempt so that every command in this session does not retry a failing
        # endpoint. The next Copilot session checks again.
        $state.checkedSessionId = $SessionId
        $state.checkedAtUtc = [DateTime]::UtcNow.ToString("o")
        Save-State -State $state
    }
} else {
    Write-Status "[excel-mcp] Freshness already checked for this Copilot session." "DarkGray"
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
Write-Status "✅ ExcelMcp MCP runtime ready." "Green"
Write-Status "   Release: $($state.latestTag)" "Gray"
Write-Status "   Binary:  $binaryPath" "Gray"
Write-Status "   Size:    $([math]::Round($binaryInfo.Length / 1MB, 2)) MB" "Gray"
