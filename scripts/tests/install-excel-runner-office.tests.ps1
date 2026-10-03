$ErrorActionPreference = 'Stop'
$guest = Join-Path (Split-Path -Parent (Split-Path -Parent $PSScriptRoot)) 'infrastructure\azure\install-excel-office.ps1'
. $guest

$global:ExcelOfficeTestSignature = 'Valid'
$global:ExcelOfficeTestSigner = 'CN=Microsoft Corporation, O=Microsoft Corporation, C=US'
$global:ExcelOfficeTestProduct = 'Excel2024Retail'
$global:ExcelOfficeTestPlatform = 'x64'
$global:ExcelOfficeTestExecutable = $true

function Get-AuthenticodeSignature {
    return @{
        Status = $global:ExcelOfficeTestSignature
        SignerCertificate = @{ Subject = $global:ExcelOfficeTestSigner }
    }
}

function Get-ItemProperty {
    return @{
        ProductReleaseIds = $global:ExcelOfficeTestProduct
        Platform = $global:ExcelOfficeTestPlatform
        VersionToReport = '16.0.1.2'
        InstallationPath = 'C:\SyntheticOffice'
    }
}

function Test-Path { return $global:ExcelOfficeTestExecutable }

Assert-MicrosoftOfficeSignature -Path 'synthetic-setup.exe'
$installation = Get-VerifiedExcelInstallation
if ($installation.product -ne 'Excel2024Retail' -or $installation.version -ne '16.0.1.2') {
    throw 'Verification must return the installed product and version.'
}

foreach ($case in @('unsigned', 'wrong-signer', 'wrong-product', 'wrong-platform', 'missing-excel')) {
    $global:ExcelOfficeTestSignature = if ($case -eq 'unsigned') { 'NotSigned' } else { 'Valid' }
    $global:ExcelOfficeTestSigner = if ($case -eq 'wrong-signer') { 'CN=Unrelated Publisher' } else {
        'CN=Microsoft Corporation, O=Microsoft Corporation, C=US'
    }
    $global:ExcelOfficeTestProduct = if ($case -eq 'wrong-product') { 'ProPlus2024Volume' } else { 'Excel2024Retail' }
    $global:ExcelOfficeTestPlatform = if ($case -eq 'wrong-platform') { 'x86' } else { 'x64' }
    $global:ExcelOfficeTestExecutable = $case -ne 'missing-excel'
    $failure = $null
    try {
        if ($case -in @('unsigned', 'wrong-signer')) {
            Assert-MicrosoftOfficeSignature -Path 'synthetic-setup.exe'
        }
        else { $null = Get-VerifiedExcelInstallation }
    }
    catch { $failure = $_.Exception.Message }
    if (-not $failure) { throw "$case must not be reported as a verified Excel installation." }
}

$shell = (Get-Command powershell.exe).Source
$exitCode = Invoke-OfficeSetupProcess -Path $shell -Arguments '-NoProfile -Command "exit 21"' -TimeoutSeconds 15
if ($exitCode -ne 21) { throw 'The installer process must retain its actual exit code.' }
$failure = $null
try {
    $null = Invoke-OfficeSetupProcess -Path $shell `
        -Arguments '-NoProfile -Command "Start-Sleep -Seconds 30"' -TimeoutSeconds 1
}
catch { $failure = $_.Exception }
if ($failure -isnot [TimeoutException]) { throw 'A hung installer must fail with a hard timeout.' }
$owned = Get-Process -Id $failure.Data['ProcessId'] -ErrorAction SilentlyContinue
if ($owned -and $owned.StartTime -eq $failure.Data['StartTime']) {
    throw 'The timed-out owned installer process must be stopped.'
}
Write-Output 'Excel installer identity, installed-state and process-deadline tests passed.'
