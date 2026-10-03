$ErrorActionPreference = 'Stop'
. (Join-Path (Split-Path -Parent (Split-Path -Parent $PSScriptRoot)) 'infrastructure\azure\test-excel-desktop.ps1')
$licensed = @"
No licenses found in user scope
Type: Device|Perpetual
Product: Excel2024Retail
LicenseState: Licensed
"@
if (-not (Test-ExcelRunnerDeviceLicense @($licensed))) { throw 'Empty user scope must not reject the licensed device.' }
if (-not (Test-ExcelRunnerDeviceLicense @([pscustomobject]@{
    Type = 'Device|Perpetual'; Product = 'Excel2024Retail'; LicenseState = 'Licensed'
}))) { throw 'Structured device evidence must work.' }
foreach ($records in @(
    @('Type: User', 'Product: Excel2024Retail', 'LicenseState: Licensed'),
    @('Type: Device|Perpetual', 'Product: Excel2024Retail', 'LicenseState: Unlicensed'),
    @('Type: Device|Perpetual', 'Product: OtherRetail', 'LicenseState: Licensed'),
    @('Type: Device|Perpetual', 'Product: Excel2024Retail'),
    @('Type: Device|Perpetual', 'Product: Excel2024Retail', 'Type: User', 'Product: OtherRetail', 'LicenseState: Licensed')
)) {
    if (Test-ExcelRunnerDeviceLicense $records) { throw 'Unrelated or incomplete licence evidence was accepted.' }
}
$source = Join-Path (Split-Path -Parent (Split-Path -Parent $PSScriptRoot)) 'infrastructure\azure\test-excel-desktop.ps1'
$ast = [Management.Automation.Language.Parser]::ParseFile($source, [ref]$null, [ref]$null)
$native = @($ast.FindAll({
    param($node)
    $node -is [Management.Automation.Language.StringConstantExpressionAst] -and
        $node.Value -match 'public static class ExcelRunnerDesktopNative'
}, $true))
if ($native.Count -ne 1) { throw 'Native credential metadata implementation was not discovered.' }
Add-Type -TypeDefinition $native[0].Value
Write-Output 'Retail device licence grouping and safe empty-user-scope tests passed.'
