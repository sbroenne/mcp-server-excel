#Requires -Version 7.0
<#
.SYNOPSIS
Builds an ad-hoc-signed, standalone native permission investigation app.
.DESCRIPTION
Development probe only, not notarized or ready for distribution. Does not launch
the app, request permissions, or install to Applications. Keep its path and
binary unchanged between setup and repeat runs.
#>
[CmdletBinding()]
param([Parameter(Mandatory)][string]$OutputDirectory)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
if (-not $IsMacOS) { throw 'The native helper build requires macOS and Xcode command-line tools.' }
$output = [IO.Path]::GetFullPath($OutputDirectory)
$bundle = Join-Path $output 'ExcelMcp Permission Probe.app'
$contents = Join-Path $bundle 'Contents'
$binaryDirectory = Join-Path $contents 'MacOS'
$resources = Join-Path $contents 'Resources'
$source = Join-Path $PSScriptRoot 'native-helper'
if ([IO.Directory]::Exists($bundle)) {
    throw 'Bundle already exists. Choose a fresh output directory rather than replacing an app with existing grants.'
}
[void][IO.Directory]::CreateDirectory($binaryDirectory)
[void][IO.Directory]::CreateDirectory($resources)
Copy-Item -LiteralPath (Join-Path $source 'Info.plist') -Destination (Join-Path $contents 'Info.plist')
Copy-Item -LiteralPath (Join-Path $PSScriptRoot 'ExcelSpike.applescript') -Destination $resources
& /usr/bin/xcrun swiftc (Join-Path $source 'main.swift') -o (Join-Path $binaryDirectory 'ExcelMcpMacPermissionProbe')
if ($LASTEXITCODE -ne 0) { throw 'Native helper compilation failed.' }
& /usr/bin/codesign --force --sign - --options runtime --entitlements (Join-Path $source 'Entitlements.plist') $bundle
if ($LASTEXITCODE -ne 0) { throw 'Development code signing failed.' }
& /usr/bin/codesign --verify --strict $bundle
if ($LASTEXITCODE -ne 0) { throw 'Development signature verification failed.' }
Write-Output $bundle
