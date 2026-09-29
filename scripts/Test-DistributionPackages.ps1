#!/usr/bin/env pwsh
[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [string]$ArchivePath,

    [Parameter(Mandatory)]
    [string]$ExecutableRelativePath,

    [Parameter(Mandatory)]
    [ValidateSet("windows-x64", "macos-arm64")]
    [string]$ExpectedArchitecture,

    [string[]]$ForbiddenExecutableRelativePath = @(),

    [string]$MacHelperRelativePath,

    [switch]$Launch
)

$ErrorActionPreference = "Stop"
Set-StrictMode -Version Latest
Add-Type -AssemblyName System.IO.Compression.FileSystem

if (-not (Test-Path -LiteralPath $ArchivePath -PathType Leaf)) {
    throw "Package not found: $ArchivePath"
}

$normalizedEntry = $ExecutableRelativePath.Replace('\', '/').TrimStart('/')
$normalizedHelperEntry = if ([string]::IsNullOrWhiteSpace($MacHelperRelativePath)) {
    $null
}
else {
    $MacHelperRelativePath.Replace('\', '/').TrimStart('/')
}
$archive = [IO.Compression.ZipFile]::OpenRead((Resolve-Path $ArchivePath))
try {
    $entry = $archive.GetEntry($normalizedEntry)
    if ($null -eq $entry) {
        throw "Package '$ArchivePath' is missing '$normalizedEntry'."
    }
    if ($entry.Length -eq 0) {
        throw "Packaged executable '$normalizedEntry' is empty."
    }
    foreach ($forbiddenPath in $ForbiddenExecutableRelativePath) {
        $normalizedForbiddenPath = $forbiddenPath.Replace('\', '/').TrimStart('/')
        if ($null -ne $archive.GetEntry($normalizedForbiddenPath)) {
            throw "Package '$ArchivePath' unexpectedly contains non-target runtime '$normalizedForbiddenPath'."
        }
    }
    if ($ExpectedArchitecture.StartsWith("macos-", [StringComparison]::Ordinal) -and
        ($entry.ExternalAttributes -band 0x00400000) -eq 0) {
        throw "Packaged macOS executable '$normalizedEntry' is not marked executable."
    }
    if ($ExpectedArchitecture.StartsWith("macos-", [StringComparison]::Ordinal) -and
        $null -ne $normalizedHelperEntry) {
        $helperEntry = $archive.GetEntry($normalizedHelperEntry)
        if ($null -eq $helperEntry -or $helperEntry.Length -eq 0) {
            throw "Package '$ArchivePath' is missing nonempty helper '$normalizedHelperEntry'."
        }
        if (($helperEntry.ExternalAttributes -band 0x00400000) -eq 0) {
            throw "Packaged macOS helper '$normalizedHelperEntry' is not marked executable."
        }
    }
}
finally {
    $archive.Dispose()
}

$extractDirectory = Join-Path ([IO.Path]::GetTempPath()) "excelmcp-package-$([Guid]::NewGuid().ToString('N'))"
try {
    [IO.Compression.ZipFile]::ExtractToDirectory((Resolve-Path $ArchivePath), $extractDirectory)
    $executable = Join-Path $extractDirectory $normalizedEntry
    if ($ExpectedArchitecture -eq "windows-x64") {
        $bytes = [IO.File]::ReadAllBytes($executable)
        if ($bytes.Length -lt 64 -or $bytes[0] -ne 0x4D -or $bytes[1] -ne 0x5A) {
            throw "'$normalizedEntry' is not a Windows PE executable."
        }
    }
    else {
        & /bin/chmod +x $executable
        $architecture = "arm64"
        $reported = (& /usr/bin/lipo -archs $executable 2>&1).Trim()
        if ($LASTEXITCODE -ne 0 -or $reported -notmatch "(^|\s)$([regex]::Escape($architecture))(\s|$)") {
            throw "'$normalizedEntry' does not contain expected Mach-O architecture '$architecture' (reported: '$reported')."
        }

        & /usr/bin/codesign --verify --strict --verbose=2 $executable
        if ($LASTEXITCODE -ne 0) {
            throw "Packaged macOS executable '$normalizedEntry' has an invalid code signature."
        }

        if ($null -ne $normalizedHelperEntry) {
            $helper = Join-Path $extractDirectory $normalizedHelperEntry
            & /bin/chmod +x $helper
            $helperReported = (& /usr/bin/lipo -archs $helper 2>&1).Trim()
            if ($LASTEXITCODE -ne 0 -or $helperReported -notmatch "(^|\s)$([regex]::Escape($architecture))(\s|$)") {
                throw "'$normalizedHelperEntry' does not contain expected Mach-O architecture '$architecture' (reported: '$helperReported')."
            }
            & /usr/bin/codesign --verify --strict --verbose=2 $helper
            if ($LASTEXITCODE -ne 0) {
                throw "Packaged macOS helper '$normalizedHelperEntry' has an invalid code signature."
            }
        }
    }

    if ($Launch) {
        if ($ExpectedArchitecture -eq "windows-x64") {
            throw "Windows artifacts are structure-inspected, never executed by this test."
        }

        $hostArchitecture = [Runtime.InteropServices.RuntimeInformation]::OSArchitecture.ToString().ToLowerInvariant()
        $expectedHost = "arm64"
        if ($hostArchitecture -ne $expectedHost) {
            throw "Cannot launch $ExpectedArchitecture on $hostArchitecture without changing the verification claim."
        }

        $output = & $executable --version 2>&1
        if ($LASTEXITCODE -ne 0 -or [string]::IsNullOrWhiteSpace(($output -join "`n"))) {
            throw "Clean-install launch verification failed for '$normalizedEntry'."
        }
    }
}
finally {
    if (Test-Path -LiteralPath $extractDirectory) {
        Remove-Item -LiteralPath $extractDirectory -Recurse -Force
    }
}

Write-Host "Verified $ExpectedArchitecture package '$ArchivePath' (launch=$($Launch.IsPresent))." -ForegroundColor Green
