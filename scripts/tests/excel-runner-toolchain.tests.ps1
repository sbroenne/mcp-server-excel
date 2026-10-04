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
$hashFixture = ${function:Get-FileHash}
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

$global:ExcelToolchainJqHash = (Get-RunnerJqRelease).sha256
$global:ExcelToolchainBashExit = 0
$global:ExcelToolchainBashVersion = 'GNU bash, version 5.3.15(2)-release (x86_64-pc-cygwin)'
$global:ExcelToolchainDevelopmentMode = 1
$bashPath = Join-Path $env:ProgramFiles 'Git\bin\bash.exe'
$jqPath = Join-Path $env:ProgramFiles 'ExcelMcp\Tools\jq.exe'
function Test-Path { param($LiteralPath, $PathType) return $true }
function Get-ItemPropertyValue {
    param($LiteralPath, $Name, $ErrorAction)
    if ($LiteralPath -ne 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\AppModelUnlock' -or
        $Name -ne 'AllowDevelopmentWithoutDevLicense') { throw 'Unexpected development-mode lookup.' }
    $global:ExcelToolchainDevelopmentMode
}
function Get-Command {
    param($Name, $CommandType, $ErrorAction)
    switch ($Name) {
        bash { @{ Source = $bashPath } }
        jq { @{ Source = $jqPath } }
        default { throw 'Unexpected cloud command lookup.' }
    }
}
Set-Item -Path "Function:\$bashPath" -Value {
    if ($args[0] -eq '--version') {
        Write-Output $global:ExcelToolchainBashVersion
        Write-Output 'Additional version output must be consumed before checking command completion.'
        $global:LASTEXITCODE = $global:ExcelToolchainBashExit
    }
    else {
        Write-Output 'jq-1.8.2'
        $global:LASTEXITCODE = 0
    }
}
Set-Item -Path Function:\Get-FileHash -Value $hashFixture
try {
    $global:LASTEXITCODE = 239
    $cloud = Get-RunnerCloudToolState
    if ($cloud.bash -ne $global:ExcelToolchainBashVersion -or $cloud.jq -ne 'jq-1.8.2' -or
        $cloud.developmentMode -ne $true) {
        throw 'Cloud verification must drain complete Bash output before checking its result.'
    }
    foreach ($case in @(
        @{ exit = 1; version = $global:ExcelToolchainBashVersion },
        @{ exit = 0; version = 'Invalid Bash version output' }
    )) {
        $global:ExcelToolchainBashExit = $case.exit
        $global:ExcelToolchainBashVersion = $case.version
        $failed = $false
        try { Get-RunnerCloudToolState | Out-Null } catch { $failed = $true }
        if (-not $failed) { throw 'Unsuccessful or invalid Bash verification must still fail.' }
    }
    $global:ExcelToolchainBashExit = 0
    $global:ExcelToolchainBashVersion = 'GNU bash, version 5.3.15(2)-release (x86_64-pc-cygwin)'
    $global:ExcelToolchainDevelopmentMode = 0
    $failed = $false
    try { Get-RunnerCloudToolState | Out-Null } catch { $failed = $true }
    if (-not $failed) { throw 'A desktop without limited-user symbolic-link support must not be admitted.' }
}
finally { Remove-Item -LiteralPath "Function:\$bashPath" }

Write-Output 'Runner SDK and trusted installer tests passed.'
