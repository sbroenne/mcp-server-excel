#Requires -Version 7.0
<#
.SYNOPSIS
Runs opt-in public CLI and MCP acceptance for macOS VBA helper candidates.
.DESCRIPTION
This script never installs the helper, changes Excel security, or creates a VBA
fixture. The workbook must be an original Excel-authored .xlsm containing the
standard-module procedure named by -MarkerProcedure. That procedure must write
its sole string argument to the exact marker cell.

Each entry point opens the exact workbook, exercises public VBA source lifecycle
commands, runs the marker procedure, verifies the marker through the public
range API, deletes the reserved acceptance module, and closes without saving.
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [string]$HelperPath,
    [Parameter(Mandatory)]
    [string]$WorkbookPath,
    [switch]$MacroApprovalConfirmed,
    [switch]$VbaProjectTrustConfirmed,
    [switch]$ExcelAuthoredWorkbookConfirmed,
    [string]$MarkerProcedure = 'ExcelMcpFixtureRun.WriteMarker',
    [string]$MarkerSheet = 'ExcelMcpAcceptance',
    [string]$MarkerCell = 'A1',
    [string]$MarkerValue = 'excelmcp-public-vba-accepted',
    [ValidateRange(10, 120)]
    [int]$OperationTimeoutSeconds = 30,
    [switch]$SkipBuild,
    [switch]$ValidateOnly
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $PSScriptRoot
$helperPath = [IO.Path]::GetFullPath($HelperPath)
$workbookPath = [IO.Path]::GetFullPath($WorkbookPath)
$moduleName = 'ExcelMcpPublicAcceptanceModule'
$moduleSource = "Option Explicit`n`nPublic Function ExcelMcpAcceptanceValue() As String`n    ExcelMcpAcceptanceValue = `"fixture`"`nEnd Function"
$updatedModuleSource = "Option Explicit`n`nPublic Function ExcelMcpAcceptanceValue() As String`n    ExcelMcpAcceptanceValue = `"updated`"`nEnd Function"

function Assert-Confirmed {
    param([bool]$Value, [string]$Name)
    if (-not $Value) { throw "-$Name is required." }
}

if (-not $ValidateOnly) {
    Assert-Confirmed $MacroApprovalConfirmed 'MacroApprovalConfirmed'
    Assert-Confirmed $VbaProjectTrustConfirmed 'VbaProjectTrustConfirmed'
    Assert-Confirmed $ExcelAuthoredWorkbookConfirmed 'ExcelAuthoredWorkbookConfirmed'
    if (-not [IO.File]::Exists($helperPath)) { throw "Helper not found: '$helperPath'." }
    if (-not [IO.File]::Exists($workbookPath)) { throw "Workbook not found: '$workbookPath'." }
    if ([IO.Path]::GetExtension($workbookPath) -cne '.xlsm') {
        throw 'Public VBA acceptance requires an .xlsm workbook.'
    }
}

if (-not $SkipBuild) {
    & dotnet build (Join-Path $root 'Sbroenne.ExcelMcp.sln') -c Release `
        -p:EnableWindowsTargeting=true --nologo
    if ($LASTEXITCODE -ne 0) { throw 'Release build failed.' }
}

$cliAssembly = Join-Path $root 'src/ExcelMcp.CLI/bin/Release/net10.0/excelcli.dll'
$mcpAssembly = Join-Path $root 'src/ExcelMcp.McpServer/bin/Release/net10.0/Sbroenne.ExcelMcp.McpServer.dll'
foreach ($assembly in @($cliAssembly, $mcpAssembly)) {
    if (-not [IO.File]::Exists($assembly)) {
        throw "Required built entry point is missing at '$assembly'."
    }
}

if ($ValidateOnly) {
    [ordered]@{
        success = $true
        helperVersion = '1.3.0'
        acceptanceScope = 'public-vba-cli-mcp'
        publicCommandAcceptance = $false
        executed = $false
    } | ConvertTo-Json -Depth 5
    exit 0
}

$environment = @{
    EXCELMCP_MAC_VBA_HELPER_PATH = $helperPath
    EXCELMCP_MAC_VBA_CANDIDATE_ACTIONS =
        'vba.list,vba.view,vba.import,vba.update,vba.delete,vba.run'
}

function ConvertFrom-StrictJson {
    param([string]$Json, [string]$Context)
    try {
        return $Json | ConvertFrom-Json -AsHashtable -Depth 32
    }
    catch {
        throw "$Context returned invalid JSON: $($_.Exception.Message)"
    }
}

function Invoke-BoundedProcess {
    param([string]$FileName, [string[]]$Arguments)
    $start = [Diagnostics.ProcessStartInfo]::new()
    $start.FileName = $FileName
    $start.UseShellExecute = $false
    $start.RedirectStandardOutput = $true
    $start.RedirectStandardError = $true
    foreach ($argument in $Arguments) { $null = $start.ArgumentList.Add($argument) }
    foreach ($entry in $environment.GetEnumerator()) {
        $start.Environment[$entry.Key] = $entry.Value
    }
    $process = [Diagnostics.Process]::new()
    $process.StartInfo = $start
    $null = $process.Start()
    if (-not $process.WaitForExit($OperationTimeoutSeconds * 1000)) {
        $process.Kill($true)
        throw "Process timed out: $FileName $($Arguments -join ' ')"
    }
    $stdout = $process.StandardOutput.ReadToEnd()
    $stderr = $process.StandardError.ReadToEnd()
    if ($process.ExitCode -ne 0) {
        throw "Process failed ($($process.ExitCode)): $stdout $stderr"
    }
    return ConvertFrom-StrictJson $stdout 'CLI command'
}

function Invoke-Cli {
    param([string[]]$Arguments)
    return Invoke-BoundedProcess dotnet (@($cliAssembly, '-q') + $Arguments)
}

function Assert-PublicSuccess {
    param([hashtable]$Result, [string]$Context)
    if ($Result.success -ne $true) {
        throw "$Context failed: $($Result | ConvertTo-Json -Compress -Depth 20)"
    }
}

function Invoke-CliAcceptance {
    $open = Invoke-Cli @('session', 'open', $workbookPath, '--timeout', "$OperationTimeoutSeconds")
    Assert-PublicSuccess $open 'CLI open'
    $session = [string]$open.sessionId
    try {
        Assert-PublicSuccess (Invoke-Cli @('vba', 'list', '--session', $session)) 'CLI list'
        Assert-PublicSuccess (Invoke-Cli @('vba', 'import', '--session', $session, '--module-name', $moduleName, '--vba-code', $moduleSource)) 'CLI import'
        $view = Invoke-Cli @('vba', 'view', '--session', $session, '--module-name', $moduleName)
        Assert-PublicSuccess $view 'CLI view'
        if ([string]$view.code -cne $moduleSource) { throw 'CLI view source mismatch.' }
        Assert-PublicSuccess (Invoke-Cli @('vba', 'update', '--session', $session, '--module-name', $moduleName, '--vba-code', $updatedModuleSource)) 'CLI update'
        Assert-PublicSuccess (Invoke-Cli @('vba', 'run', '--session', $session, '--procedure-name', $MarkerProcedure, '--timeout', "$OperationTimeoutSeconds", '--parameters', $MarkerValue)) 'CLI run'
        $marker = Invoke-Cli @('range', 'get-values', '--session', $session, '--sheet', $MarkerSheet, '--range', $MarkerCell)
        Assert-PublicSuccess $marker 'CLI marker read'
        if ([string]$marker.values[0][0] -cne $MarkerValue) { throw 'CLI marker value mismatch.' }
        Assert-PublicSuccess (Invoke-Cli @('vba', 'delete', '--session', $session, '--module-name', $moduleName)) 'CLI delete'
    }
    finally {
        $null = Invoke-Cli @('session', 'close', '--session', $session)
    }
}

function Start-Mcp {
    $start = [Diagnostics.ProcessStartInfo]::new()
    $start.FileName = 'dotnet'
    $null = $start.ArgumentList.Add($mcpAssembly)
    $start.UseShellExecute = $false
    $start.RedirectStandardInput = $true
    $start.RedirectStandardOutput = $true
    $start.RedirectStandardError = $true
    foreach ($entry in $environment.GetEnumerator()) {
        $start.Environment[$entry.Key] = $entry.Value
    }
    $process = [Diagnostics.Process]::new()
    $process.StartInfo = $start
    $null = $process.Start()
    return $process
}

function Read-McpResponse {
    param([Diagnostics.Process]$Process, [int]$Id)
    while ($true) {
        $task = $Process.StandardOutput.ReadLineAsync()
        $line = $task.WaitAsync(
            [TimeSpan]::FromSeconds($OperationTimeoutSeconds)).GetAwaiter().GetResult()
        if ($null -eq $line) { throw "MCP exited before response $Id." }
        $message = ConvertFrom-StrictJson $line 'MCP response'
        if ($message.id -eq $Id) { return $message }
    }
}

function Write-McpMessage {
    param([Diagnostics.Process]$Process, [hashtable]$Message)
    $Process.StandardInput.WriteLine(
        ($Message | ConvertTo-Json -Compress -Depth 30))
    $Process.StandardInput.Flush()
}

function Invoke-McpTool {
    param(
        [Diagnostics.Process]$Process,
        [ref]$NextId,
        [string]$Name,
        [hashtable]$Arguments
    )
    $id = $NextId.Value
    $NextId.Value++
    Write-McpMessage $Process @{
        jsonrpc = '2.0'
        id = $id
        method = 'tools/call'
        params = @{ name = $Name; arguments = $Arguments }
    }
    $response = Read-McpResponse $Process $id
    if ($null -ne $response.error) {
        throw "MCP $Name failed: $($response.error | ConvertTo-Json -Compress -Depth 10)"
    }
    $text = [string]$response.result.content[0].text
    $result = ConvertFrom-StrictJson $text "MCP $Name result"
    Assert-PublicSuccess $result "MCP $Name"
    return $result
}

function Invoke-McpAcceptance {
    $process = Start-Mcp
    $nextId = 1
    $session = $null
    try {
        Write-McpMessage $process @{
            jsonrpc = '2.0'
            id = $nextId
            method = 'initialize'
            params = @{
                protocolVersion = '2025-06-18'
                capabilities = @{}
                clientInfo = @{ name = 'excelmcp-public-vba-acceptance'; version = '1.0' }
            }
        }
        $null = Read-McpResponse $process $nextId
        $nextId++
        Write-McpMessage $process @{
            jsonrpc = '2.0'
            method = 'notifications/initialized'
            params = @{}
        }

        $open = Invoke-McpTool $process ([ref]$nextId) 'file' @{
            action = 'open'
            path = $workbookPath
            show = $false
            timeout_seconds = $OperationTimeoutSeconds
        }
        $session = [string]$open.sessionId
        $null = Invoke-McpTool $process ([ref]$nextId) 'vba' @{
            action = 'list'; session_id = $session
        }
        $null = Invoke-McpTool $process ([ref]$nextId) 'vba' @{
            action = 'import'; session_id = $session
            module_name = $moduleName; vba_code = $moduleSource
        }
        $view = Invoke-McpTool $process ([ref]$nextId) 'vba' @{
            action = 'view'; session_id = $session; module_name = $moduleName
        }
        if ([string]$view.code -cne $moduleSource) { throw 'MCP view source mismatch.' }
        $null = Invoke-McpTool $process ([ref]$nextId) 'vba' @{
            action = 'update'; session_id = $session
            module_name = $moduleName; vba_code = $updatedModuleSource
        }
        $null = Invoke-McpTool $process ([ref]$nextId) 'vba' @{
            action = 'run'; session_id = $session
            procedure_name = $MarkerProcedure
            timeout_seconds = $OperationTimeoutSeconds
            parameters = @($MarkerValue)
        }
        $marker = Invoke-McpTool $process ([ref]$nextId) 'range' @{
            action = 'get-values'; session_id = $session
            sheet_name = $MarkerSheet; range_address = $MarkerCell
        }
        if ([string]$marker.values[0][0] -cne $MarkerValue) { throw 'MCP marker value mismatch.' }
        $null = Invoke-McpTool $process ([ref]$nextId) 'vba' @{
            action = 'delete'; session_id = $session; module_name = $moduleName
        }
    }
    finally {
        if ($null -ne $session -and -not $process.HasExited) {
            try {
                $null = Invoke-McpTool $process ([ref]$nextId) 'file' @{
                    action = 'close'; session_id = $session; save = $false
                }
            }
            catch { }
        }
        if (-not $process.HasExited) {
            $process.StandardInput.Close()
            if (-not $process.WaitForExit(5000)) { $process.Kill($true) }
        }
        $process.Dispose()
    }
}

Invoke-CliAcceptance
Invoke-McpAcceptance

[ordered]@{
    success = $true
    helperVersion = '1.3.0'
    acceptanceScope = 'public-vba-cli-mcp'
    publicCommandAcceptance = $true
    executed = $true
    actions = @('vba.list', 'vba.view', 'vba.import', 'vba.update', 'vba.delete', 'vba.run')
    marker = @{
        procedure = $MarkerProcedure
        sheet = $MarkerSheet
        cell = $MarkerCell
        value = $MarkerValue
    }
} | ConvertTo-Json -Depth 10
