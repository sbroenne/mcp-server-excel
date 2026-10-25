#Requires -Version 7.4
<#
.SYNOPSIS
Installs the official Microsoft ADO import tool on explicitly provisioned build machines.
#>
[CmdletBinding()]
param()

$ErrorActionPreference = 'Stop'
$PSNativeCommandUseErrorActionPreference = $true

if (-not $IsWindows) { throw 'The Microsoft ADO import tool requires Windows.' }
$tool = Join-Path ${env:ProgramFiles(x86)} 'Microsoft SDKs\Windows\v10.0A\bin\NETFX 4.8.1 Tools\TlbImp.exe'
if (Test-Path -LiteralPath $tool -PathType Leaf) {
    & (Join-Path $PSScriptRoot 'Import-AdoDefinitions.ps1') -ValidateOnly
    return
}

$directory = Join-Path ([IO.Path]::GetTempPath()) "ExcelMcp-AdoBuildTools-$([Guid]::NewGuid().ToString('N'))"
New-Item -ItemType Directory -Path $directory | Out-Null
$installationFailure = $null
try {
    $installer = Join-Path $directory 'ndp481-devpack-enu.exe'
    Invoke-WebRequest -Uri 'https://go.microsoft.com/fwlink/?linkid=2203306' -OutFile $installer
    $expectedHash = 'FD9AE4B63CBC8F7637249C4B29D25281B1F70FBE6D650109DCBF2BDF596CF994'
    if ((Get-FileHash -LiteralPath $installer -Algorithm SHA256).Hash -ne $expectedHash) {
        throw 'The Microsoft .NET Framework 4.8.1 Developer Pack download does not match the reviewed SHA256.'
    }
    $signature = Get-AuthenticodeSignature -LiteralPath $installer
    if ($signature.Status -ne 'Valid' -or
        $signature.SignerCertificate.Subject -notmatch '(?:^|,\s*)O=Microsoft Corporation(?:,|$)') {
        throw 'The Developer Pack installer must have a valid Microsoft signature.'
    }
    $process = Start-Process -FilePath $installer -ArgumentList '/quiet', '/norestart' -Wait -PassThru
    try {
        if ($process.ExitCode -ne 0) {
            throw "Developer Pack installation returned $($process.ExitCode). A restart or installation repair may be required; no restart was requested."
        }
    }
    finally {
        $process.Dispose()
    }
    & (Join-Path $PSScriptRoot 'Import-AdoDefinitions.ps1') -ValidateOnly
}
catch {
    $installationFailure = $_.Exception
    throw
}
finally {
    try {
        Remove-Item -LiteralPath $directory -Recurse -Force
    }
    catch {
        if ($installationFailure) {
            throw [AggregateException]::new('ADO build-tool installation and staging cleanup both failed.',
                [Exception[]]@($installationFailure, $_.Exception))
        }
        throw
    }
}
