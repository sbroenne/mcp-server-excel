$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
. (Join-Path $root 'scripts\ExcelRunnerPolicy.ps1')
$source = Join-Path $root 'infrastructure\azure\excel-job-started.ps1'
$ast = [Management.Automation.Language.Parser]::ParseFile($source, [ref]$null, [ref]$null)
$guards = @($ast.FindAll({
    param($node)
    $node -is [Management.Automation.Language.IfStatementAst] -and
        $node.Extent.Text -match '\$permit\.runId'
}, $true))
if ($guards.Count -ne 1) { throw 'The actual job admission guard was not discovered.' }
$guard = [scriptblock]::Create($guards[0].Extent.Text)
$values = @{
    GITHUB_REPOSITORY = 'sbroenne/mcp-server-excel'
    GITHUB_RUN_ID = '37152390285'
    GITHUB_RUN_ATTEMPT = '1'
    GITHUB_JOB = 'copilot'
    GITHUB_ACTOR = 'Copilot'
    GITHUB_EVENT_NAME = 'dynamic'
    GITHUB_REF = 'refs/heads/copilot/qualification'
}
$saved = @{}
foreach ($name in $values.Keys) {
    $saved[$name] = [Environment]::GetEnvironmentVariable($name)
    [Environment]::SetEnvironmentVariable($name, $values[$name])
}
try {
    $boot = [DateTime]::UtcNow.AddMinutes(-5)
    $permit = @{
        runId = $values.GITHUB_RUN_ID; job = 'copilot'; kind = 'cloud'; state = 'admitted'
        bootTime = $boot.ToString('o'); expiresAt = [DateTime]::UtcNow.AddMinutes(10).ToString('o')
        defaultBranch = 'main'
    }
    & $guard
    $env:GITHUB_EVENT_NAME = 'workflow_dispatch'
    $failure = $null
    try { & $guard } catch { $failure = $_.Exception }
    if (-not $failure) { throw 'A different cloud event was admitted.' }
    if ($failure.Message -notmatch '"event":"workflow_dispatch"' -or
        $failure.Message -notmatch '"job":"copilot"' -or
        $failure.Message -notmatch '"attempt":"1"') {
        throw 'Rejected cloud admission must identify the actual non-secret runtime event, job and attempt.'
    }
}
finally {
    foreach ($name in $saved.Keys) { [Environment]::SetEnvironmentVariable($name, $saved[$name]) }
}
Write-Output 'Actual cloud admission guard and safe rejection context passed.'
