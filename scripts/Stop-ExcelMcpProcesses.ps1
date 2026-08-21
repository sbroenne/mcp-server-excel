<#
.SYNOPSIS
    Stops the ExcelMCP CLI service and Excel processes owned by one CLI pipe.
.DESCRIPTION
    Invokes the CLI-owned service stop path for EXCELMCP_CLI_PIPE (or the
    default user pipe when no override is set). The CLI validates tracked
    process start times before stopping anything.

    If an existing CLI is older than the ownership lifecycle sources, a
    current cleanup client is built in an isolated temporary output first.
    If no CLI binary exists yet, the script safely does nothing. It never
    scans for or stops unrelated Excel or service processes.
.NOTES
    Called once from Directory.Build.props before the CLI project builds.
    Safe to run when no processes are running (silently succeeds).
#>

param(
    [string]$PipeName = $env:EXCELMCP_CLI_PIPE,
    [switch]$Verbose
)

$ErrorActionPreference = 'Stop'

function Write-Status($message) {
    if ($Verbose) {
        Write-Host "  [pre-build] $message" -ForegroundColor DarkGray
    }
}

function Remove-StagingClient([string]$path) {
    if ([string]::IsNullOrWhiteSpace($path) -or -not (Test-Path -LiteralPath $path)) {
        return
    }

    try {
        Remove-Item -LiteralPath $path -Recurse -Force
    }
    catch {
        Write-Host "  Isolated owned-cleanup client could not be removed: $($_.Exception.Message)" -ForegroundColor Yellow
    }
}

$repoRoot = Split-Path -Parent $PSScriptRoot
$cliPaths = @(
    (Join-Path $repoRoot 'src\ExcelMcp.CLI\bin\Release\net10.0-windows\excelcli.exe'),
    (Join-Path $repoRoot 'src\ExcelMcp.CLI\bin\Debug\net10.0-windows\excelcli.exe')
)
$safetyInputs = @(
    $PSCommandPath,
    (Join-Path $repoRoot 'src\ExcelMcp.CLI\Commands\ServiceCommands.cs'),
    (Join-Path $repoRoot 'src\ExcelMcp.CLI\Infrastructure\DaemonAutoStart.cs'),
    (Join-Path $repoRoot 'src\ExcelMcp.CLI\Infrastructure\DaemonPipeIdentity.cs'),
    (Join-Path $repoRoot 'src\ExcelMcp.CLI\Infrastructure\DaemonProcessTracker.cs'),
    (Join-Path $repoRoot 'src\ExcelMcp.CLI\Infrastructure\DaemonStartupLock.cs'),
    (Join-Path $repoRoot 'src\ExcelMcp.CLI\Infrastructure\DaemonTrackingJson.cs'),
    (Join-Path $repoRoot 'src\ExcelMcp.CLI\Infrastructure\OwnedProcessCleanup.cs'),
    (Join-Path $repoRoot 'src\ExcelMcp.Cleanup\ExcelMcp.Cleanup.csproj'),
    (Join-Path $repoRoot 'src\ExcelMcp.Cleanup\Program.cs'),
    (Join-Path $repoRoot 'src\ExcelMcp.Cleanup\ParameterTransforms.cs'),
    (Join-Path $repoRoot 'src\ExcelMcp.Service\ServiceClient.cs'),
    (Join-Path $repoRoot 'src\ExcelMcp.Service\ServiceProtocol.cs'),
    (Join-Path $repoRoot 'src\ExcelMcp.Service\Rpc\IExcelDaemonRpc.cs'),
    (Join-Path $repoRoot 'src\ExcelMcp.ComInterop\ServiceClient\ServiceProtocol.cs'),
    (Join-Path $repoRoot 'src\ExcelMcp.CLI\Program.cs'),
    (Join-Path $repoRoot 'src\ExcelMcp.ComInterop\Session\ExcelBatch.cs'),
    (Join-Path $repoRoot 'src\ExcelMcp.ComInterop\Session\ExcelProcessIdentity.cs'),
    (Join-Path $repoRoot 'src\ExcelMcp.ComInterop\Session\OwnedProcessGuard.cs'),
    (Join-Path $repoRoot 'src\ExcelMcp.ComInterop\Session\SessionManager.cs'),
    (Join-Path $repoRoot 'src\ExcelMcp.Service\ServiceSecurity.cs')
)
$latestSafetyInput = $safetyInputs |
    Where-Object { Test-Path $_ } |
    ForEach-Object { (Get-Item $_).LastWriteTimeUtc } |
    Sort-Object -Descending |
    Select-Object -First 1
$availableClis = @($cliPaths |
    Where-Object { Test-Path $_ } |
    ForEach-Object { Get-Item $_ } |
    Sort-Object LastWriteTimeUtc -Descending |
    Select-Object -First 1)
$excelcli = $availableClis |
    Where-Object { $_.LastWriteTimeUtc -ge $latestSafetyInput } |
    Select-Object -First 1
$stagingRoot = $null
$useCleanupBootstrap = $false

if (-not $excelcli) {
    if ($availableClis.Count -eq 0) {
        Write-Status 'excelcli is missing; owned cleanup skipped safely'
        exit 0
    }

    $stagingRoot = Join-Path ([System.IO.Path]::GetTempPath()) "excelmcp-cleanup-$PID-$([Guid]::NewGuid().ToString('N'))"
    $cleanupProject = Join-Path $repoRoot 'src\ExcelMcp.Cleanup\ExcelMcp.Cleanup.csproj'
    $stagedCliPath = Join-Path $stagingRoot 'bin\ExcelMcp.Cleanup\Release\net10.0-windows\excelmcp-cleanup.exe'
    Write-Status "Existing CLI is older than the cleanup sources; building isolated cleanup client in $stagingRoot"

    try {
        $buildOutput = & dotnet build $cleanupProject `
            --configuration Release `
            -p:ExcelMcpCleanupRoot=$stagingRoot `
            -p:ExcelMcpSkipCleanup=true `
            -p:NuGetAudit=false `
            -maxcpucount:1 `
            -nodeReuse:false `
            --verbosity quiet 2>&1
        $buildExitCode = $LASTEXITCODE
    }
    catch {
        Write-Host "  Isolated owned-cleanup client could not be built: $($_.Exception.Message). No process sweep was attempted." -ForegroundColor Yellow
        Remove-StagingClient $stagingRoot
        exit 1
    }

    if ($buildExitCode -ne 0 -or -not (Test-Path $stagedCliPath)) {
        Write-Host "  Isolated owned-cleanup client build failed with exit code $buildExitCode. No process sweep was attempted." -ForegroundColor Yellow
        if ($buildOutput) {
            Write-Status ($buildOutput | Out-String).Trim()
        }
        Remove-StagingClient $stagingRoot
        exit $(if ($buildExitCode -ne 0) { $buildExitCode } else { 1 })
    }

    $excelcli = Get-Item $stagedCliPath
    $useCleanupBootstrap = $true
}

$previousPipeName = $env:EXCELMCP_CLI_PIPE
try {
    if ([string]::IsNullOrWhiteSpace($PipeName)) {
        Remove-Item Env:EXCELMCP_CLI_PIPE -ErrorAction SilentlyContinue
        Write-Status 'Using the default CLI pipe'
    }
    else {
        $env:EXCELMCP_CLI_PIPE = $PipeName
        Write-Status "Using CLI pipe: $PipeName"
    }

    Write-Status "Using CLI: $($excelcli.FullName)"
    try {
        if ($useCleanupBootstrap) {
            $output = & $excelcli.FullName 2>&1
        }
        else {
            $output = & $excelcli.FullName service stop --quiet 2>&1
        }
        $exitCode = $LASTEXITCODE
    }
    catch {
        Write-Host "  Owned CLI cleanup could not start: $($_.Exception.Message). No process sweep was attempted." -ForegroundColor Yellow
        exit 1
    }

    if ($exitCode -ne 0) {
        Write-Host "  Owned CLI cleanup failed with exit code $exitCode. No process sweep was attempted." -ForegroundColor Yellow
        if ($output) {
            Write-Status ($output | Out-String).Trim()
        }
        exit $exitCode
    }

    Write-Status 'Owned CLI cleanup completed'
}
finally {
    if ($null -eq $previousPipeName) {
        Remove-Item Env:EXCELMCP_CLI_PIPE -ErrorAction SilentlyContinue
    }
    else {
        $env:EXCELMCP_CLI_PIPE = $previousPipeName
    }

    Remove-StagingClient $stagingRoot
}

exit 0
