$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
. (Join-Path $root 'infrastructure\azure\install-excel-toolchain.ps1')

$nativePython = @(Get-Command python -CommandType Application -ErrorAction Stop)[0].Source
if ((Get-RunnerPythonArchitecture $nativePython) -ne '64') {
    throw 'The actual native Python architecture probe must survive both supported shells without losing quoted arguments.'
}

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
$global:ExcelToolchainSignatureSubject = 'CN=Python Software Foundation, O=Python Software Foundation, C=US'
Assert-ToolchainInstallerSignature 'synthetic.exe' Python
foreach ($subject in @(
    'CN=Python Software Foundation, O=Unrelated Publisher, C=US',
    'CN=Unrelated Publisher, O=Python Software Foundation, C=US'
)) {
    $global:ExcelToolchainSignatureSubject = $subject
    $failed = $false
    try { Assert-ToolchainInstallerSignature 'synthetic.exe' Python } catch { $failed = $true }
    if (-not $failed) { throw 'Python installers require the verified foundation publisher and organization.' }
}

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
$global:ExcelToolchainPythonPresent = $true
$global:ExcelToolchainPythonVersion = 'Python ' + (Get-RunnerPythonRelease).version
$global:ExcelToolchainPythonArchitecture = '64'
$global:ExcelToolchainPythonPip = 'pip 26.2 from synthetic (python 3.13)'
$global:ExcelToolchainPythonExit = 0
$global:ExcelToolchainPythonPathMatches = $true
$global:ExcelToolchainDuplicatePythonCommands = $false
$bashPath = Join-Path $env:ProgramFiles 'Git\bin\bash.exe'
$jqPath = Join-Path $env:ProgramFiles 'ExcelMcp\Tools\jq.exe'
$pythonPath = Join-Path $env:ProgramFiles 'Python313\python.exe'
function Test-Path {
    param($LiteralPath, $PathType)
    if ($LiteralPath -eq $pythonPath) { return $global:ExcelToolchainPythonPresent }
    return $true
}
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
        python {
            @{ Source = $(if ($global:ExcelToolchainPythonPathMatches) { $pythonPath } else { 'unprotected-python.exe' }) }
            if ($global:ExcelToolchainDuplicatePythonCommands) { @{ Source = 'later-python-alias.exe' } }
        }
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
Set-Item -Path "Function:\$pythonPath" -Value {
    switch ($args[0]) {
        --version { Write-Output $global:ExcelToolchainPythonVersion }
        -c { Write-Output $global:ExcelToolchainPythonArchitecture }
        -m { Write-Output $global:ExcelToolchainPythonPip }
        default { throw 'Unexpected Python verification invocation.' }
    }
    $global:LASTEXITCODE = $global:ExcelToolchainPythonExit
}
Set-Item -Path Function:\Get-FileHash -Value $hashFixture
try {
    $global:LASTEXITCODE = 239
    $cloud = Get-RunnerCloudToolState
    if ($cloud.bash -ne $global:ExcelToolchainBashVersion -or $cloud.jq -ne 'jq-1.8.2' -or
        $cloud.developmentMode -ne $true -or $cloud.python -ne $global:ExcelToolchainPythonVersion) {
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
    foreach ($case in @('missing', 'wrong-version', 'outdated-version', 'wrong-architecture', 'missing-pip', 'failed-command', 'unprotected-path')) {
        $global:ExcelToolchainPythonPresent = $case -ne 'missing'
        $global:ExcelToolchainPythonVersion = switch ($case) {
            wrong-version { 'Python 3.12.13' }
            outdated-version { 'Python 3.13.0' }
            default { 'Python ' + (Get-RunnerPythonRelease).version }
        }
        $global:ExcelToolchainPythonArchitecture = if ($case -eq 'wrong-architecture') { '32' } else { '64' }
        $global:ExcelToolchainPythonPip = if ($case -eq 'missing-pip') { 'No module named pip' } else { 'pip 26.2 from synthetic (python 3.13)' }
        $global:ExcelToolchainPythonExit = if ($case -eq 'failed-command') { 1 } else { 0 }
        $global:ExcelToolchainPythonPathMatches = $case -ne 'unprotected-path'
        $failed = $false
        try { Get-RunnerCloudToolState | Out-Null } catch { $failed = $true }
        if (-not $failed) { throw "Python readiness must reject $case." }
    }
    $global:ExcelToolchainPythonPresent = $true
    $global:ExcelToolchainPythonVersion = 'Python ' + (Get-RunnerPythonRelease).version
    $global:ExcelToolchainPythonArchitecture = '64'
    $global:ExcelToolchainPythonPip = 'pip 26.2 from synthetic (python 3.13)'
    $global:ExcelToolchainPythonExit = 0
    $global:ExcelToolchainPythonPathMatches = $true
    $global:ExcelToolchainDuplicatePythonCommands = $true
    $cloud = Get-RunnerCloudToolState
    if ($cloud.python -ne $global:ExcelToolchainPythonVersion) {
        throw 'A later Python alias must not override the first protected application on PATH.'
    }
    $global:ExcelToolchainDuplicatePythonCommands = $false
    $global:ExcelToolchainPythonInstalls = 0
    function Install-RunnerPrerequisite {
        param($Uri, $FileName, $Publisher, $Arguments, [switch]$Msi)
        if ($Uri -ne (Get-RunnerPythonRelease).uri -or $FileName -ne 'python.exe' -or $Publisher -ne 'Python' -or $Msi -or
            $Arguments -notmatch '/quiet InstallAllUsers=1' -or
            -not $Arguments.Contains(('TargetDir="' + (Split-Path -Parent $pythonPath) + '"')) -or
            $Arguments -notmatch 'Include_launcher=0') {
            throw 'Python must use the bounded, verified all-users installer in protected Program Files.'
        }
        $global:ExcelToolchainPythonInstalls++
        $global:ExcelToolchainPythonPresent = $true
    }
    $global:ExcelToolchainPythonPresent = $false
    Install-RunnerPythonPrerequisite
    Install-RunnerPythonPrerequisite
    if ($global:ExcelToolchainPythonInstalls -ne 1) { throw 'Missing Python must be installed exactly once and verified on reuse.' }
    $global:ExcelToolchainDevelopmentMode = 0
    $failed = $false
    try { Get-RunnerCloudToolState | Out-Null } catch { $failed = $true }
    if (-not $failed) { throw 'A desktop without limited-user symbolic-link support must not be admitted.' }
    $global:ExcelToolchainProvisioningCalls = [Collections.Generic.List[string]]::new()
    function New-Item {
        param($ItemType, $Path, [switch]$Force)
        if ($Path -ne (Join-Path $env:ProgramFiles 'ExcelMcp\Tools') -and
            $Path -ne 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\AppModelUnlock') {
            throw 'Unexpected prerequisite directory or registry key.'
        }
    }
    function New-ItemProperty {
        param($LiteralPath, $Name, $PropertyType, $Value, [switch]$Force)
        if ($LiteralPath -ne 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\AppModelUnlock' -or
            $Name -ne 'AllowDevelopmentWithoutDevLicense' -or $PropertyType -ne 'DWord' -or $Value -ne 1) {
            throw 'Provisioning must enable the exact supported development-mode DWORD.'
        }
        $global:ExcelToolchainDevelopmentMode = $Value
        $global:ExcelToolchainProvisioningCalls.Add('development-mode')
    }
    function Set-RunnerCloudToolPath {
        if ($global:ExcelToolchainDevelopmentMode -ne 1) { throw 'Development Mode must be enabled before PATH/readiness.' }
        $global:ExcelToolchainProvisioningCalls.Add('path')
    }
    foreach ($initialMode in @($null, 0)) {
        $global:ExcelToolchainDevelopmentMode = $initialMode
        $global:ExcelToolchainProvisioningCalls.Clear()
        $cloud = Install-RunnerCloudPrerequisites
        if ($cloud.developmentMode -ne $true -or
            ($global:ExcelToolchainProvisioningCalls -join ',') -ne 'development-mode,path') {
            throw 'Absent or disabled Developer Mode must be provisioned before actual readiness succeeds.'
        }
    }
}
finally {
    Remove-Item -LiteralPath "Function:\$bashPath"
    Remove-Item -LiteralPath "Function:\$pythonPath"
}

$setup = Get-Content (Join-Path $root '.github\workflows\copilot-setup-steps.yml') -Raw
if ($setup -notmatch '(?m)^      - name: Setup Python\r?\n        if: runner.environment != ''self-hosted''\r?\n        uses: actions/setup-python@') {
    throw 'Only hosted setup may invoke the first-time administrative setup-python installer.'
}

Write-Output 'Runner SDK and trusted installer tests passed.'
