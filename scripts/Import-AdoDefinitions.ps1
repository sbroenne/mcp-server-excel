#Requires -Version 7.4
<#
.SYNOPSIS
Imports Windows ADO 2.8 definitions for embedding by the C# compiler.
#>
[CmdletBinding()]
param(
    [string]$OutputPath,
    [string]$ImporterPath = (Join-Path ${env:ProgramFiles(x86)} 'Microsoft SDKs\Windows\v10.0A\bin\NETFX 4.8.1 Tools\TlbImp.exe'),
    [string]$TypeLibraryPath = (Join-Path $env:CommonProgramFiles 'System\ado\msado28.tlb'),
    [switch]$ValidateOnly
)

$ErrorActionPreference = 'Stop'
$PSNativeCommandUseErrorActionPreference = $true

if (-not $IsWindows) {
    throw 'ADO definition import requires Windows and the Microsoft .NET Framework 4.8.1 Developer Pack.'
}
if (-not (Test-Path -LiteralPath $ImporterPath -PathType Leaf)) {
    throw 'TlbImp.exe is missing. Install the Microsoft .NET Framework 4.8.1 Developer Pack: https://go.microsoft.com/fwlink/?linkid=2203306. See docs/agents/development.md#windows-setup.'
}
if (-not (Test-Path -LiteralPath $TypeLibraryPath -PathType Leaf)) {
    throw "Windows ADO 2.8 definitions are missing: $TypeLibraryPath"
}

foreach ($path in @($ImporterPath, $TypeLibraryPath)) {
    $signature = Get-AuthenticodeSignature -LiteralPath $path
    if ($signature.Status -ne 'Valid' -or
        $signature.SignerCertificate.Subject -notmatch '(?:^|,\s*)O=Microsoft Corporation(?:,|$)') {
        throw "A valid Microsoft signature is required: $path"
    }
}

$tool = Get-Item -LiteralPath $ImporterPath
if ($tool.VersionInfo.ProductVersion -ne '4.8.9037.0') {
    throw "Expected Microsoft TlbImp.exe 4.8.9037.0; found $($tool.VersionInfo.ProductVersion). Install the .NET Framework 4.8.1 Developer Pack."
}

foreach ($path in @($ImporterPath, $TypeLibraryPath)) {
    $item = Get-Item -LiteralPath $path
    $hash = (Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash
    Write-Host "ADO compiler input: $($item.Name), version $($item.VersionInfo.FileVersion), SHA256 $hash"
}
if ($ValidateOnly) { return }
if ([string]::IsNullOrWhiteSpace($OutputPath)) {
    throw 'OutputPath is required when importing ADO definitions.'
}

$OutputPath = [IO.Path]::GetFullPath($OutputPath)
New-Item -ItemType Directory -Path (Split-Path -Parent $OutputPath) -Force | Out-Null
if (Test-Path -LiteralPath $OutputPath) {
    Remove-Item -LiteralPath $OutputPath
}
try {
    # strictref:nopia ignores registered PIAs; TI3001 only reports that one exists.
    & $ImporterPath $TypeLibraryPath /namespace:ADODB /machine:Agnostic /asmversion:2.8.0.0 "/out:$OutputPath" /nologo /strictref:nopia /silence:3001
    if ($LASTEXITCODE -ne 0) {
        throw "Microsoft TlbImp.exe failed with exit code $LASTEXITCODE."
    }
    if (-not (Test-Path -LiteralPath $OutputPath -PathType Leaf)) {
        throw 'Microsoft TlbImp.exe did not produce the requested ADO definitions.'
    }
}
catch {
    $importFailure = $_.Exception
    try {
        if (Test-Path -LiteralPath $OutputPath) {
            Remove-Item -LiteralPath $OutputPath
        }
    }
    catch {
        throw [AggregateException]::new('ADO import and incomplete-output cleanup both failed.',
            [Exception[]]@($importFailure, $_.Exception))
    }
    throw
}
