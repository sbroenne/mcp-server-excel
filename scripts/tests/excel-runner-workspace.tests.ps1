$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
$source = Join-Path $root 'infrastructure\azure\excel-job-completed.ps1'
$directory = Join-Path ([IO.Path]::GetTempPath()) ('excel-workspace-test-' + [Guid]::NewGuid().ToString('N'))
$workspace = Join-Path $directory 'checkout'
$outside = Join-Path $directory 'outside'
$junction = Join-Path $workspace 'linked'
$childScript = Join-Path $directory 'check.ps1'
$executable = (Get-Process -Id $PID).Path

function Invoke-WorkspaceCleanup {
    $start = [Diagnostics.ProcessStartInfo]::new()
    $start.FileName = $executable
    $start.Arguments = '-NoLogo -NoProfile -NonInteractive -File "' + $childScript + '" -Source "' + $source + '" -Workspace "' + $workspace + '"'
    $start.WorkingDirectory = $workspace
    $start.UseShellExecute = $false
    $process = [Diagnostics.Process]::Start($start)
    try {
        $null = $process.Handle
        if (-not $process.WaitForExit(30000)) {
            $process.Kill()
            throw 'Workspace cleanup did not finish within its test deadline.'
        }
        return $process.ExitCode
    }
    finally { $process.Dispose() }
}

try {
    New-Item -ItemType Directory -Path $workspace, $outside | Out-Null
    Set-Content -LiteralPath (Join-Path $outside 'sentinel.txt') -Value 'untouched'
    @'
param([string]$Source, [string]$Workspace)
$ErrorActionPreference = 'Stop'
$ast = [Management.Automation.Language.Parser]::ParseFile($Source, [ref]$null, [ref]$null)
$helpers = @($ast.EndBlock.Statements | Where-Object {
    $_ -is [Management.Automation.Language.FunctionDefinitionAst] -and $_.Name -eq 'Remove-ExcelRunnerJobWorkspace'
})
if ($helpers.Count -ne 1) { throw 'Workspace cleanup requires its actual isolated production helper.' }
. ([scriptblock]::Create($helpers[0].Extent.Text))
function Assert-ExcelRunnerWorkspace {
    param([string]$Path)
    if ($Path -ne $Workspace) { throw 'Cleanup escaped the isolated test checkout.' }
}
Remove-ExcelRunnerJobWorkspace $Workspace
'@ | Set-Content -LiteralPath $childScript -Encoding UTF8

    New-Item -ItemType Directory -Path (Join-Path $workspace 'nested') | Out-Null
    Set-Content -LiteralPath (Join-Path $workspace 'nested\workbook-marker.txt') -Value 'job-owned'
    Set-Content -LiteralPath (Join-Path $workspace '.hidden-marker') -Value 'job-owned'
    [IO.File]::SetAttributes((Join-Path $workspace '.hidden-marker'), [IO.FileAttributes]::Hidden)
    if ((Invoke-WorkspaceCleanup) -ne 0) { throw 'Cleanup must succeed while its process uses the checkout as its current directory.' }
    if (-not (Test-Path -LiteralPath $workspace) -or @(Get-ChildItem -LiteralPath $workspace -Force).Count) {
        throw 'Cleanup must remove all checkout contents and retain its empty working directory.'
    }
    if ((Get-Content -LiteralPath (Join-Path $outside 'sentinel.txt')) -ne 'untouched') { throw 'Cleanup changed a neighboring directory.' }
    if ((Invoke-WorkspaceCleanup) -ne 0) { throw 'Repeating cleanup of the empty checkout must succeed.' }

    New-Item -ItemType Junction -Path $junction -Target $outside | Out-Null
    if ((Invoke-WorkspaceCleanup) -eq 0) { throw 'A linked checkout descendant must fail cleanup explicitly.' }
    if ((Get-Content -LiteralPath (Join-Path $outside 'sentinel.txt')) -ne 'untouched') { throw 'Cleanup followed the checkout link.' }
}
finally {
    if (Test-Path -LiteralPath $junction) { [IO.Directory]::Delete($junction) }
    if (Test-Path -LiteralPath $directory) { Remove-Item -LiteralPath $directory -Recurse -Force }
}
Write-Output 'Actual checkout cleanup passed with a held working directory, nested and hidden contents, repeated cleanup and link rejection.'
