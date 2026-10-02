#!/usr/bin/env pwsh
[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [string]$ArchivePath,

    [switch]$AllowUnnotarized
)

$ErrorActionPreference = "Stop"
Set-StrictMode -Version Latest

if (-not $IsMacOS) {
    throw "Apple notarization must run on macOS."
}
if (-not (Test-Path -LiteralPath $ArchivePath -PathType Leaf)) {
    throw "Archive not found: $ArchivePath"
}
$resolvedArchivePath = (Resolve-Path -LiteralPath $ArchivePath).Path

$key = $env:MACOS_NOTARY_API_KEY_P8
$keyId = $env:MACOS_NOTARY_KEY_ID
$issuer = $env:MACOS_NOTARY_ISSUER_ID
$configured = @(@($key, $keyId, $issuer) | Where-Object {
    -not [string]::IsNullOrWhiteSpace($_)
})

if ($configured.Count -eq 0) {
    if (-not $AllowUnnotarized) {
        throw "Notarization credentials are required. Use -AllowUnnotarized only for explicit local validation."
    }
    Write-Warning "Notarization credentials are not configured. '$ArchivePath' remains explicitly unnotarized for local validation."
    return
}
if ($configured.Count -ne 3) {
    throw "MACOS_NOTARY_API_KEY_P8, MACOS_NOTARY_KEY_ID, and MACOS_NOTARY_ISSUER_ID must be configured together."
}
if ([string]::IsNullOrWhiteSpace($env:MACOS_SIGNING_IDENTITY)) {
    throw "Refusing to notarize an archive whose executable was not Developer ID signed."
}

$temporaryDirectory = Join-Path ([IO.Path]::GetTempPath()) "excelmcp-notary-$([Guid]::NewGuid().ToString('N'))"
$keyPath = Join-Path $temporaryDirectory "AuthKey.p8"
$submissionPath = $resolvedArchivePath
[IO.Directory]::CreateDirectory(
    $temporaryDirectory,
    [IO.UnixFileMode]::UserRead -bor
        [IO.UnixFileMode]::UserWrite -bor
        [IO.UnixFileMode]::UserExecute) | Out-Null
try {
    [IO.File]::WriteAllText($keyPath, $key, [Text.UTF8Encoding]::new($false))
    [IO.File]::SetUnixFileMode(
        $keyPath,
        [IO.UnixFileMode]::UserRead -bor [IO.UnixFileMode]::UserWrite)
    if ([IO.Path]::GetExtension($resolvedArchivePath) -notin @(".zip", ".pkg", ".dmg")) {
        $submissionPath = Join-Path $temporaryDirectory "submission.zip"
        Copy-Item -LiteralPath $resolvedArchivePath -Destination $submissionPath
    }

    $result = & /usr/bin/xcrun notarytool submit $submissionPath `
        --key $keyPath `
        --key-id $keyId `
        --issuer $issuer `
        --wait `
        --output-format json
    if ($LASTEXITCODE -ne 0) {
        throw "Apple notarization submission failed for $ArchivePath."
    }

    $response = $result | ConvertFrom-Json
    if ($response.status -ne "Accepted") {
        throw "Apple notarization did not accept '$ArchivePath' (status: $($response.status), id: $($response.id))."
    }

    Write-Host "Apple notarization accepted '$ArchivePath' (submission $($response.id))." -ForegroundColor Green
}
finally {
    Remove-Item -LiteralPath $temporaryDirectory -Recurse -Force -ErrorAction SilentlyContinue
}
