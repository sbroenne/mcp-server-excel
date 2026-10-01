#!/usr/bin/env pwsh
[CmdletBinding()]
param()

$ErrorActionPreference = "Stop"
Set-StrictMode -Version Latest

if (-not $IsMacOS) {
    throw "macOS code-signing initialization must run on macOS."
}

$certificate = $env:MACOS_CERTIFICATE_P12
$password = $env:MACOS_CERTIFICATE_PASSWORD
$identity = $env:MACOS_SIGNING_IDENTITY
$configured = @(@($certificate, $password, $identity) | Where-Object {
    -not [string]::IsNullOrWhiteSpace($_)
})

if ($configured.Count -eq 0) {
    Write-Warning "Developer ID signing secrets are not configured. macOS binaries will be ad-hoc signed and must not be described as notarized."
    exit 0
}

if ($configured.Count -ne 3) {
    throw "MACOS_CERTIFICATE_P12, MACOS_CERTIFICATE_PASSWORD, and MACOS_SIGNING_IDENTITY must be configured together."
}

$temporaryDirectory = Join-Path ([IO.Path]::GetTempPath()) "excelmcp-signing-$([Guid]::NewGuid().ToString('N'))"
$certificatePath = Join-Path $temporaryDirectory "developer-id.p12"
$keychainPath = Join-Path $temporaryDirectory "excelmcp.keychain-db"
$keychainPassword = [Guid]::NewGuid().ToString("N")
New-Item -ItemType Directory -Path $temporaryDirectory | Out-Null

function Invoke-Security {
    param(
        [Parameter(Mandatory)]
        [string[]]$Arguments
    )

    $output = & /usr/bin/security @Arguments 2>&1
    if ($LASTEXITCODE -ne 0) {
        throw "security $($Arguments[0]) failed: $($output -join [Environment]::NewLine)"
    }
    return @($output)
}

try {
    [IO.File]::WriteAllBytes($certificatePath, [Convert]::FromBase64String($certificate))
    [IO.File]::SetUnixFileMode(
        $certificatePath,
        [IO.UnixFileMode]::UserRead -bor [IO.UnixFileMode]::UserWrite)
    Invoke-Security @("create-keychain", "-p", $keychainPassword, $keychainPath) | Out-Null
    Invoke-Security @("set-keychain-settings", "-lut", "21600", $keychainPath) | Out-Null
    Invoke-Security @("unlock-keychain", "-p", $keychainPassword, $keychainPath) | Out-Null
    Invoke-Security @(
        "import", $certificatePath, "-k", $keychainPath, "-P", $password,
        "-T", "/usr/bin/codesign") | Out-Null
    Invoke-Security @(
        "set-key-partition-list", "-S", "apple-tool:,apple:", "-s",
        "-k", $keychainPassword, $keychainPath) | Out-Null
    Invoke-Security @("list-keychains", "-d", "user", "-s", $keychainPath) | Out-Null
    $identities = Invoke-Security @("find-identity", "-v", "-p", "codesigning", $keychainPath)
    if (($identities -join "`n") -notmatch [regex]::Escape($identity)) {
        throw "Configured Developer ID signing identity '$identity' was not imported."
    }

    if (-not [string]::IsNullOrWhiteSpace($env:GITHUB_ENV)) {
        "EXCELMCP_SIGNING_KEYCHAIN=$keychainPath" | Add-Content $env:GITHUB_ENV
    }
}
finally {
    Remove-Item -LiteralPath $certificatePath -Force -ErrorAction SilentlyContinue
}

Write-Host "Imported the configured Developer ID signing identity into an ephemeral keychain."
