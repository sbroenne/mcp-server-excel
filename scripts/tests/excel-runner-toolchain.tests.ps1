$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
. (Join-Path $root 'infrastructure\azure\install-excel-toolchain.ps1')

$global:ExcelToolchainSignatureStatus = 'Valid'
$global:ExcelToolchainSignatureSubject = 'CN=Microsoft Corporation, O=Microsoft Corporation'
function Get-AuthenticodeSignature {
    @{
        Status = $global:ExcelToolchainSignatureStatus
        SignerCertificate = @{ Subject = $global:ExcelToolchainSignatureSubject }
    }
}
Assert-ToolchainInstallerSignature 'synthetic.exe' Microsoft
$global:ExcelToolchainSignatureSubject = 'CN=.NET, O=Microsoft Corporation, C=US'
Assert-ToolchainInstallerSignature 'synthetic.exe' DotNet
$global:ExcelToolchainSignatureSubject = 'CN=.NET, O=Unrelated Publisher, C=US'
$failed = $false
try { Assert-ToolchainInstallerSignature 'synthetic.exe' DotNet }
catch { $failed = $true }
if (-not $failed) { throw 'The .NET certificate must belong to Microsoft, not only use the .NET name.' }
$global:ExcelToolchainSignatureSubject = 'CN=Microsoft Corporation, O=Microsoft Corporation'
foreach ($case in @('NotSigned', 'HashMismatch', 'UnknownError')) {
    $global:ExcelToolchainSignatureStatus = $case
    $failed = $false
    try { Assert-ToolchainInstallerSignature 'synthetic.exe' Microsoft }
    catch { $failed = $true }
    if (-not $failed) { throw 'Invalid installer signatures must be rejected.' }
}
$global:ExcelToolchainSignatureStatus = 'Valid'
$global:ExcelToolchainSignatureSubject = 'CN=Unrelated Publisher, O=Unrelated Publisher'
$failed = $false
try { Assert-ToolchainInstallerSignature 'synthetic.exe' Microsoft }
catch { $failed = $true }
if (-not $failed) { throw 'An unrelated valid publisher must be rejected.' }

$global:ExcelToolchainSignatureSubject = 'CN=OpenJS Foundation, O=OpenJS Foundation'
Assert-ToolchainInstallerSignature 'synthetic.msi' Node
$global:ExcelToolchainSignatureSubject = 'CN=Open Source Developer, Johannes Schindelin, O=Open Source Developer'
Assert-ToolchainInstallerSignature 'synthetic.exe' Git

$global:ExcelToolchainJqHash = (Get-RunnerJqRelease).sha256
function Get-FileHash {
    param($LiteralPath, $Algorithm)
    if ($Algorithm -ne 'SHA256') { throw 'Cloud tools must use SHA256.' }
    @{ Hash = $global:ExcelToolchainJqHash }
}
Assert-RunnerJqPackage 'synthetic-jq.exe'
$global:ExcelToolchainJqHash = '0' * 64
$failed = $false
try { Assert-RunnerJqPackage 'synthetic-jq.exe' } catch { $failed = $true }
if (-not $failed) { throw 'Unverified cloud jq binaries must be rejected before execution.' }

$required = (Get-Content (Join-Path $root 'global.json') -Raw | ConvertFrom-Json).sdk.version
Assert-RequiredRunnerSdk -Required $required -Installed @("$required [synthetic]")
$requiredVersion = [version]$required
$newer = "$($requiredVersion.Major).$($requiredVersion.Minor).$($requiredVersion.Build + 100)"
Assert-RequiredRunnerSdk -Required $required -Installed @("$newer [synthetic]") -RollForward latestFeature
$failed = $false
try { Assert-RequiredRunnerSdk -Required $required -Installed @("$newer [synthetic]") -RollForward disable }
catch { $failed = $true }
if (-not $failed) { throw 'Disabled SDK roll-forward must still require the exact version.' }
$failed = $false
try { Assert-RequiredRunnerSdk -Required $required -Installed @('10.0.100 [synthetic]') }
catch { $failed = $true }
if (-not $failed) { throw 'An older SDK must not substitute for the requested SDK.' }

Write-Output 'Runner SDK and trusted installer tests passed.'
