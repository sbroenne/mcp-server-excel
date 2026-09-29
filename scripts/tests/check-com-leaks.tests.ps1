#!/usr/bin/env pwsh

$ErrorActionPreference = "Stop"
$rootDir = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
$checker = Join-Path $rootDir "scripts\check-com-leaks.ps1"
$fixtureDir = Join-Path ([IO.Path]::GetTempPath()) "excelmcp-com-leak-check-$([Guid]::NewGuid())"

try {
    New-Item -ItemType Directory -Path $fixtureDir | Out-Null

    $safeFixture = Join-Path $fixtureDir "Safe.cs"
    @'
dynamic? rows = null;
try
{
    rows = range.Rows;
    int count = rows.Count;
}
finally
{
    ComUtilities.Release(ref rows);
}
'@ | Set-Content -LiteralPath $safeFixture

    & pwsh -NoProfile -File $checker -InputPath $safeFixture
    if ($LASTEXITCODE -ne 0) {
        throw "Expected the safe fixture to pass."
    }

    $unsafeFixture = Join-Path $fixtureDir "Unsafe.cs"
    @'
int count = range.Rows.Count;
int row = table.Range.Row;
int column = table.Range.Column;
namesCollection.Add("Example", "=Sheet1!A1");
'@ | Set-Content -LiteralPath $unsafeFixture

    & pwsh -NoProfile -File $checker -InputPath $unsafeFixture
    if ($LASTEXITCODE -eq 0) {
        throw "Expected the unsafe fixture to fail."
    }

    Write-Host "COM leak checker fixture tests passed." -ForegroundColor Green
}
finally {
    if (Test-Path -LiteralPath $fixtureDir) {
        Remove-Item -LiteralPath $fixtureDir -Recurse -Force
    }
}
