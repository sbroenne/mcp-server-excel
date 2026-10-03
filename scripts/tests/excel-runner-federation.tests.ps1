$ErrorActionPreference = 'Stop'
. (Join-Path (Split-Path -Parent (Split-Path -Parent $PSScriptRoot)) 'infrastructure\azure\configure-excel-control-oidc.ps1')
$definition = Get-ExcelRunnerControlApplicationDefinition
$json = $definition | ConvertTo-Json -Depth 4
$decoded = $json | ConvertFrom-Json
if ($decoded.displayName -ne 'excelmcp-github-runner-control' -or
    $decoded.signInAudience -ne 'AzureADMyOrg' -or $decoded.tags -isnot [array] -or
    $decoded.tags.Count -ne 1 -or $decoded.tags[0] -ne 'ExcelMcpRunnerControl') {
    throw 'Atomic application creation must include a JSON array ownership tag.'
}
$environment = @{
    deployment_branch_policy = @{ custom_branch_policies = $true; protected_branches = $false }
    protection_rules = @(@{ type = 'required_reviewers'; reviewers = @('synthetic-reviewer') })
}
$policies = @{ branch_policies = @(@{ name = 'main'; type = 'branch' }) }
Assert-ExcelRunnerControlEnvironment $environment $policies 'main'
if ($environment.protection_rules[0].reviewers[0] -ne 'synthetic-reviewer') { throw 'Existing reviewer protection must be retained.' }
foreach ($branches in @(
    @(),
    @(@{ name = '*'; type = 'branch' }),
    @(@{ name = 'main'; type = 'tag' }),
    @(@{ name = 'main'; type = 'branch' }, @{ name = 'feature'; type = 'branch' })
)) {
    $failed = $false
    try { Assert-ExcelRunnerControlEnvironment $environment @{ branch_policies = $branches } 'main' }
    catch { $failed = $true }
    if (-not $failed) { throw 'Unexpected deployment trust must be rejected before modifying environment settings.' }
}
$environment.deployment_branch_policy.protected_branches = $true
$failed = $false
try { Assert-ExcelRunnerControlEnvironment $environment $policies 'main' } catch { $failed = $true }
if (-not $failed) { throw 'All-protected-branch trust is broader than the default branch.' }
Write-Output 'Default-branch federation and preservation of existing environment protections passed.'
