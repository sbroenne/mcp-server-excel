#Requires -Version 7.0
<#
.SYNOPSIS
Packages an Excel-authored, SelfCert-signed macOS VBA helper with provenance.

.DESCRIPTION
Treats ExcelMcpHelper.xlam as an opaque whole file. The script never opens or
inspects workbook package contents. It records whole-file hashes, reviewed BAS
source, a public-only self-signed code-signing certificate, and the maintainer's
explicit confirmation that Excel displayed the expected VBA signature.

The SelfCert private key must remain on the controlled Windows signing host.
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [string]$HelperPath,
    [Parameter(Mandatory)]
    [string]$SourcePath,
    [Parameter(Mandatory)]
    [string]$PublicCertificatePath,
    [Parameter(Mandatory)]
    [string]$OutputDirectory,
    [Parameter(Mandatory)]
    [ValidatePattern('^[0-9a-fA-F]{7,64}$')]
    [string]$SourceCommit,
    [switch]$ExcelSignatureVerifiedConfirmed,
    [switch]$SelfCertTrustModelConfirmed
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Resolve-InputFile {
    param([string]$Path, [string]$Name)

    if (-not [IO.Path]::IsPathFullyQualified($Path)) {
        throw "$Name must be an absolute path."
    }

    $resolved = [IO.Path]::GetFullPath($Path)
    if (-not [IO.File]::Exists($resolved)) {
        throw "$Name does not exist at '$resolved'."
    }

    return $resolved
}

function Get-Sha256 {
    param([string]$Path)

    return (Get-FileHash -LiteralPath $Path -Algorithm SHA256).Hash.ToUpperInvariant()
}

if (-not $ExcelSignatureVerifiedConfirmed) {
    throw '-ExcelSignatureVerifiedConfirmed is required. Confirm the expected VBA signature in Windows Excel after the final save and reopen.'
}
if (-not $SelfCertTrustModelConfirmed) {
    throw '-SelfCertTrustModelConfirmed is required. SelfCert does not provide CA-validated publisher identity and requires manual trust on every machine.'
}

$helper = Resolve-InputFile $HelperPath 'HelperPath'
$source = Resolve-InputFile $SourcePath 'SourcePath'
$certificateFile = Resolve-InputFile $PublicCertificatePath 'PublicCertificatePath'
if (-not [string]::Equals(
        [IO.Path]::GetFileName($helper),
        'ExcelMcpHelper.xlam',
        [StringComparison]::Ordinal)) {
    throw "HelperPath must name the exact artifact 'ExcelMcpHelper.xlam'."
}
if (-not [string]::Equals(
        [IO.Path]::GetExtension($certificateFile),
        '.cer',
        [StringComparison]::OrdinalIgnoreCase)) {
    throw "PublicCertificatePath must name an exported public '.cer' file."
}

$certificateText = [Text.Encoding]::ASCII.GetString([IO.File]::ReadAllBytes($certificateFile))
$privateKeyMarkers = @(
    '-----BEGIN PRIVATE KEY-----',
    '-----BEGIN ENCRYPTED PRIVATE KEY-----',
    '-----BEGIN RSA PRIVATE KEY-----',
    '-----BEGIN EC PRIVATE KEY-----'
)
foreach ($marker in $privateKeyMarkers) {
    if ($certificateText.Contains($marker, [StringComparison]::Ordinal)) {
        throw 'PublicCertificatePath must not contain private-key material.'
    }
}

$sourceText = [IO.File]::ReadAllText($source)
$helperVersionMatch = [regex]::Match(
    $sourceText,
    'Private Const HELPER_VERSION As String = "([^"]+)"',
    [Text.RegularExpressions.RegexOptions]::CultureInvariant)
$protocolVersionMatch = [regex]::Match(
    $sourceText,
    'Private Const PROTOCOL_VERSION As Long = ([0-9]+)',
    [Text.RegularExpressions.RegexOptions]::CultureInvariant)
if (-not $helperVersionMatch.Success -or -not $protocolVersionMatch.Success) {
    throw 'SourcePath does not declare the required helper and protocol versions.'
}

$certificate = [Security.Cryptography.X509Certificates.X509CertificateLoader]::LoadCertificateFromFile(
    $certificateFile)
try {
    if ($certificate.HasPrivateKey) {
        throw 'PublicCertificatePath must not contain a private key.'
    }
    if (-not [string]::Equals(
            $certificate.Subject,
            $certificate.Issuer,
            [StringComparison]::OrdinalIgnoreCase)) {
        throw 'PublicCertificatePath must contain the dedicated self-signed helper certificate.'
    }

    $codeSigningOid = '1.3.6.1.5.5.7.3.3'
    $hasCodeSigningUsage = $false
    foreach ($extension in $certificate.Extensions) {
        if ($extension -is [Security.Cryptography.X509Certificates.X509EnhancedKeyUsageExtension]) {
            foreach ($usage in $extension.EnhancedKeyUsages) {
                if ($usage.Value -eq $codeSigningOid) {
                    $hasCodeSigningUsage = $true
                }
            }
        }
    }
    if (-not $hasCodeSigningUsage) {
        throw 'PublicCertificatePath must include the Code Signing enhanced key usage.'
    }

    $now = [DateTime]::UtcNow
    if ($now -lt $certificate.NotBefore.ToUniversalTime() -or
        $now -gt $certificate.NotAfter.ToUniversalTime()) {
        throw 'PublicCertificatePath is not currently valid.'
    }

    $output = [IO.Path]::GetFullPath($OutputDirectory)
    if ([IO.Directory]::Exists($output) -or [IO.File]::Exists($output)) {
        throw "OutputDirectory already exists: '$output'."
    }

    $parent = [IO.Path]::GetDirectoryName($output)
    if ([string]::IsNullOrWhiteSpace($parent)) {
        throw 'OutputDirectory must have a parent directory.'
    }
    [IO.Directory]::CreateDirectory($parent) | Out-Null
    $staging = Join-Path $parent (
        ".$([IO.Path]::GetFileName($output)).staging-$([Guid]::NewGuid().ToString('N'))")
    [IO.Directory]::CreateDirectory($staging) | Out-Null

    try {
        $packagedHelper = Join-Path $staging 'ExcelMcpHelper.xlam'
        $packagedSource = Join-Path $staging 'ExcelMcpHelper.bas'
        $packagedCertificate = Join-Path $staging 'ExcelMcpHelper.cer'
        [IO.File]::Copy($helper, $packagedHelper, $false)
        [IO.File]::Copy($source, $packagedSource, $false)
        [IO.File]::Copy($certificateFile, $packagedCertificate, $false)

        $manifest = [ordered]@{
            schemaVersion = 1
            helperVersion = $helperVersionMatch.Groups[1].Value
            protocolVersion = [int]$protocolVersionMatch.Groups[1].Value
            sourceCommit = $SourceCommit.ToLowerInvariant()
            trustModel = 'self-signed-manual-per-machine'
            excelSignatureVerified = $true
            artifactFileName = 'ExcelMcpHelper.xlam'
            artifactSha256 = Get-Sha256 $packagedHelper
            sourceSha256 = Get-Sha256 $packagedSource
            certificateSha256 = Get-Sha256 $packagedCertificate
            certificateSubject = $certificate.Subject
            certificateIssuer = $certificate.Issuer
            certificateThumbprint = $certificate.Thumbprint.ToUpperInvariant()
            certificateNotBeforeUtc = $certificate.NotBefore.ToUniversalTime().ToString('O')
            certificateNotAfterUtc = $certificate.NotAfter.ToUniversalTime().ToString('O')
            privateKeyIncluded = $false
            selfCertWarning = 'SelfCert does not validate publisher identity through a public CA. Verify the published SHA-256 fingerprint and trust the public certificate manually on each machine.'
        }
        $manifestPath = Join-Path $staging 'ExcelMcpHelper.manifest.json'
        [IO.File]::WriteAllText(
            $manifestPath,
            ($manifest | ConvertTo-Json -Depth 4) + [Environment]::NewLine,
            [Text.UTF8Encoding]::new($false))

        [IO.Directory]::Move($staging, $output)
        Write-Host "Created opaque signed-helper package: $output" -ForegroundColor Green
        Write-Host "Certificate SHA-256 fingerprint: $($manifest.certificateSha256)"
        Write-Host 'The SelfCert public certificate must be trusted manually on every machine.'
    }
    catch {
        if ([IO.Directory]::Exists($staging)) {
            [IO.Directory]::Delete($staging, $true)
        }
        throw
    }
}
finally {
    $certificate.Dispose()
}
