$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
$source = Join-Path $root 'infrastructure\azure\excel-job-completed.ps1'
$ast = [Management.Automation.Language.Parser]::ParseFile($source, [ref]$null, [ref]$null)
$helpers = @($ast.EndBlock.Statements | Where-Object {
    $_ -is [Management.Automation.Language.FunctionDefinitionAst] -and $_.Name -eq 'Stop-ExcelRunnerJobProcess'
})
if ($helpers.Count -ne 1) { throw 'Job process cleanup requires a callable, isolated ownership helper.' }
. ([scriptblock]::Create($helpers[0].Extent.Text))
$script:Stopped = [Collections.Generic.List[int]]::new()
$script:Disposed = 0
$script:Owner = @{ ReturnValue = 0; User = 'excelrunner'; Domain = 'synthetic-machine' }
$created = [DateTime]::UtcNow
$script:Process = [pscustomobject]@{ Id = 777; Handle = 1; StartTime = $created; HasExited = $false }
$script:Process | Add-Member ScriptMethod WaitForExit { param($Timeout) return $true }
$script:Process | Add-Member ScriptMethod Dispose { $script:Disposed++ }
function Get-Process { param($Id, $ErrorAction) return $script:Process }
function Invoke-CimMethod { param($InputObject, $MethodName) return $script:Owner }
function Stop-Process { param($Id, [switch]$Force) $script:Stopped.Add($Id) }
$entry = @{ ProcessId = 777; CreationDate = [DateTime]::new($created.Ticks - $created.Ticks % 10, [DateTimeKind]::Utc) }
Stop-ExcelRunnerJobProcess $entry 'synthetic-machine'
if ($script:Stopped.Count -ne 1 -or $script:Stopped[0] -ne 777 -or $script:Disposed -ne 1) {
    throw 'Exact microsecond identity must stop its owned process and dispose once.'
}
$script:Stopped.Clear()
$script:Disposed = 0
$script:Owner.User = 'unrelated-user'
Stop-ExcelRunnerJobProcess $entry 'synthetic-machine'
if ($script:Stopped.Count -or $script:Disposed -ne 1) { throw 'An unrelated process must remain untouched, with its handle disposed.' }
$script:Owner.User = 'excelrunner'
$script:Owner.ReturnValue = 2
$failed = $false
try { Stop-ExcelRunnerJobProcess $entry 'synthetic-machine' } catch { $failed = $true }
if (-not $failed -or $script:Stopped.Count -or $script:Disposed -ne 2) { throw 'Unknown ownership must fail explicitly and release the handle.' }
$script:Owner.ReturnValue = 0
$script:Process.StartTime = $created.AddSeconds(1)
Stop-ExcelRunnerJobProcess $entry 'synthetic-machine'
if ($script:Stopped.Count -or $script:Disposed -ne 3) { throw 'A reused PID must not be stopped.' }
$script:Process = $null
Stop-ExcelRunnerJobProcess $entry 'synthetic-machine'
Write-Output 'Owned process cleanup, timestamp precision, handle lifetime and exit races passed.'
