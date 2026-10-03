$ErrorActionPreference = 'Stop'
. (Join-Path (Split-Path -Parent (Split-Path -Parent $PSScriptRoot)) 'infrastructure\azure\configure-excel-runner.ps1')
$hostSource = Join-Path (Split-Path -Parent $PSScriptRoot) 'Register-ExcelAgentRunner.ps1'
$hostAst = [Management.Automation.Language.Parser]::ParseFile($hostSource, [ref]$null, [ref]$null)
$templates = @($hostAst.FindAll({
    param($node)
    $node -is [Management.Automation.Language.ExpandableStringExpressionAst] -and
        $node.Value -match 'configured-offline'
}, $true))
if ($templates.Count -ne 1) { throw 'Actual runner task configuration command was not discovered.' }
$configured = [pscustomobject]@{ operationId = 'synthetic-operation'; agentId = 123; state = 'configured' }
$guestScript = & ([scriptblock]::Create($templates[0].Extent.Text))
if ($guestScript -notmatch 'agentId = 123\b' -or $guestScript -match 'synthetic-operation') {
    throw 'The protected registration marker must contain the numeric agent ID, not an interpolated result object.'
}
$asset = @{
    name = 'actions-runner-win-x64-2.337.0.zip'
    browser_download_url = 'https://github.com/actions/runner/releases/download/v2.337.0/actions-runner-win-x64-2.337.0.zip'
    digest = 'sha256:' + ('a' * 64)
}
Assert-ExcelRunnerPackage $asset '2.337.0'
foreach ($field in @('name', 'browser_download_url', 'digest')) {
    $candidate = $asset.Clone()
    $candidate[$field] = 'unverified'
    $failure = $null
    try { Assert-ExcelRunnerPackage $candidate '2.337.0' } catch { $failure = $_.Exception }
    if (-not $failure) { throw 'Unverified or unrelated runner package was accepted.' }
}
$rsa = New-ExcelRunnerRegistrationKey
try {
    if ($rsa.KeySize -ne 2048 -or $rsa.PersistKeyInCsp) { throw 'Registration keys must be transient 2048-bit keys.' }
    $key = [Text.Encoding]::UTF8.GetBytes($rsa.ToXmlString($true))
    $entropy = [Text.Encoding]::UTF8.GetBytes('synthetic-operation')
    $protected = [Security.Cryptography.ProtectedData]::Protect(
        $key, $entropy, [Security.Cryptography.DataProtectionScope]::CurrentUser)
    $unprotected = [Security.Cryptography.ProtectedData]::Unprotect(
        $protected, $entropy, [Security.Cryptography.DataProtectionScope]::CurrentUser)
    if ([Convert]::ToBase64String($key) -ne [Convert]::ToBase64String($unprotected)) {
        throw 'Actual private-key Windows data protection did not round-trip.'
    }
    $plain = [Text.Encoding]::UTF8.GetBytes('synthetic-registration-token')
    $encrypted = $rsa.Encrypt($plain, [Security.Cryptography.RSAEncryptionPadding]::OaepSHA1)
    $decoded = $rsa.Decrypt($encrypted, [Security.Cryptography.RSAEncryptionPadding]::OaepSHA1)
    if ([Text.Encoding]::UTF8.GetString($decoded) -ne 'synthetic-registration-token') { throw 'Transient token transport failed.' }
}
finally { $rsa.Dispose() }
Write-Output 'Pinned runner package and private registration transport tests passed.'
