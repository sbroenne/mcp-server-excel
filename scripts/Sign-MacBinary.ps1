#!/usr/bin/env pwsh
[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [string]$Path,

    [string]$Identity = $env:MACOS_SIGNING_IDENTITY,

    [switch]$AutomationClient
)

$ErrorActionPreference = "Stop"
Set-StrictMode -Version Latest

if (-not $IsMacOS) {
    throw "macOS binaries can only be signed and verified on macOS."
}
if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) {
    throw "macOS executable not found: $Path"
}

$signingIdentity = if ([string]::IsNullOrWhiteSpace($Identity)) { "-" } else { $Identity }
$arguments = @("--force", "--sign", $signingIdentity)
if ($signingIdentity -ne "-") {
    $arguments += @("--options", "runtime", "--timestamp")
    if ($AutomationClient) {
        $entitlementsPath = Join-Path $PSScriptRoot "macos/ExcelMcp.Automation.entitlements.plist"
        if (-not (Test-Path -LiteralPath $entitlementsPath -PathType Leaf)) {
            throw "Apple Events entitlements file not found: $entitlementsPath"
        }
        $arguments += @("--entitlements", $entitlementsPath)
    }
    if (-not [string]::IsNullOrWhiteSpace($env:EXCELMCP_SIGNING_KEYCHAIN)) {
        $arguments += @("--keychain", $env:EXCELMCP_SIGNING_KEYCHAIN)
    }
}
$arguments += $Path

& /usr/bin/codesign @arguments
if ($LASTEXITCODE -ne 0) {
    throw "codesign failed for $Path."
}
& /usr/bin/codesign --verify --strict --verbose=2 $Path
if ($LASTEXITCODE -ne 0) {
    throw "codesign verification failed for $Path."
}
if ($AutomationClient -and $signingIdentity -ne "-") {
    $signedEntitlements = & /usr/bin/codesign --display --entitlements - $Path 2>&1
    if ($LASTEXITCODE -ne 0 -or
        ($signedEntitlements -join "`n") -notmatch "com\.apple\.security\.automation\.apple-events") {
        throw "Signed automation client is missing the Apple Events entitlement: $Path"
    }
}

if ($signingIdentity -eq "-") {
    Write-Warning "Ad-hoc signed '$Path'. This verifies package integrity only; it is not Developer ID signed or notarized."
}
else {
    Write-Host "Developer ID signed and verified '$Path'." -ForegroundColor Green
}
