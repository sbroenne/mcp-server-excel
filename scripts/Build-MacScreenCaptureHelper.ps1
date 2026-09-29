#Requires -Version 7.0
[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [ValidateSet('osx-arm64')]
    [string]$RuntimeIdentifier,
    [string]$OutputRoot = (Join-Path $PSScriptRoot '../artifacts/native')
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
if (-not $IsMacOS) {
    throw 'The ScreenCaptureKit helper can only be built on macOS.'
}

$architecture = 'arm64'
$root = Split-Path -Parent $PSScriptRoot
$source = Join-Path $root 'src/ExcelMcp.MacScreenCapture/main.swift'
$outputDirectory = Join-Path $OutputRoot "$RuntimeIdentifier/helpers"
$output = Join-Path $outputDirectory 'excelmcp-screencapture'
$temporary = "$output.tmp"

New-Item -ItemType Directory -Path $outputDirectory -Force | Out-Null
try {
    & swiftc `
        -parse-as-library `
        -warnings-as-errors `
        -target "$architecture-apple-macos14.0" `
        -O `
        -o $temporary `
        $source `
        -framework AppKit `
        -framework CoreGraphics `
        -framework ImageIO `
        -framework ScreenCaptureKit `
        -framework UniformTypeIdentifiers
    if ($LASTEXITCODE -ne 0) {
        throw "swiftc failed for $RuntimeIdentifier with exit code $LASTEXITCODE."
    }

    $actualArchitecture = (& lipo -archs $temporary).Trim()
    if ($actualArchitecture -ne $architecture) {
        throw "Expected $architecture helper but swiftc produced '$actualArchitecture'."
    }

    & chmod 755 $temporary
    if ($LASTEXITCODE -ne 0) {
        throw "chmod failed for $temporary."
    }
    Move-Item -LiteralPath $temporary -Destination $output -Force
}
finally {
    Remove-Item -LiteralPath $temporary -Force -ErrorAction SilentlyContinue
}

Write-Output $output
