#Requires -Version 7.0
<#
.SYNOPSIS
Runs opt-in public CLI and MCP acceptance for disabled Office.js PivotTable candidates.
.DESCRIPTION
This runner does not install or trust certificates, sideload the task pane, edit
the bridge allowlist, launch Excel, or change security settings. Before running,
the user must complete docs/MACOS-OFFICEJS.md, start the bridge, activate the
task pane for the exact saved workbook, and deliberately add the required
candidate actions to bridge.json.

The dedicated workbook must contain two ordinary local PivotTables per entry
point. The primary pivots need a placed row field and value field. The removal
pivots are sacrificial because Office.js placement actions remain disabled.

Example:
  pwsh ./scripts/Test-MacOfficeJsAcceptance.ps1 `
    -WorkbookPath /absolute/path/OfficeJsAcceptance.xlsx `
    -CliPivotTable CliPivot -CliRemovalPivotTable CliRemovalPivot `
    -McpPivotTable McpPivot -McpRemovalPivotTable McpRemovalPivot `
    -RowField Region -ValueField Sales -SelectedItem North `
    -UserSetupConfirmed -CandidateAllowlistConfirmed -ExcelSlotConfirmed
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [string]$WorkbookPath,
    [Parameter(Mandatory)]
    [string]$CliPivotTable,
    [Parameter(Mandatory)]
    [string]$CliRemovalPivotTable,
    [Parameter(Mandatory)]
    [string]$McpPivotTable,
    [Parameter(Mandatory)]
    [string]$McpRemovalPivotTable,
    [Parameter(Mandatory)]
    [string]$RowField,
    [Parameter(Mandatory)]
    [string]$ValueField,
    [Parameter(Mandatory)]
    [string]$SelectedItem,
    [string]$BridgeConfigPath = (Join-Path $HOME 'Library/Application Support/ExcelMcp/officejs/bridge.json'),
    [switch]$UserSetupConfirmed,
    [switch]$CandidateAllowlistConfirmed,
    [switch]$ExcelSlotConfirmed,
    [switch]$SkipBuild,
    [switch]$ValidateOnly,
    [ValidateRange(5, 120)]
    [int]$OperationTimeoutSeconds = 30
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $PSScriptRoot
$requiredActions = @(
    'pivottablefield.remove-field',
    'pivottablefield.set-field-name',
    'pivottablefield.set-field-format',
    'pivottablefield.set-field-filter',
    'pivottablefield.sort-field',
    'pivottablecalc.get-data',
    'pivottablecalc.set-layout',
    'pivottablecalc.set-subtotals',
    'pivottablecalc.set-grand-totals'
)

function Assert-Confirmed {
    param([bool]$Value, [string]$Name, [string]$Meaning)
    if (-not $Value) {
        throw "-$Name is required. $Meaning"
    }
}

function Resolve-ExactFile {
    param([string]$Path, [string]$Name)
    if (-not [IO.Path]::IsPathFullyQualified($Path)) {
        throw "$Name must be an absolute path."
    }
    $fullPath = [IO.Path]::GetFullPath($Path)
    if (-not [IO.File]::Exists($fullPath)) {
        throw "$Name does not exist at '$fullPath'."
    }
    return $fullPath
}

function ConvertFrom-StrictJson {
    param([string]$Json, [string]$Context)
    if ([string]::IsNullOrWhiteSpace($Json)) {
        throw "$Context returned no JSON."
    }
    try {
        return $Json | ConvertFrom-Json -AsHashtable -Depth 32
    }
    catch {
        throw "$Context returned invalid JSON: $($_.Exception.Message)"
    }
}

function Invoke-BoundedProcess {
    param(
        [string]$Executable,
        [string[]]$Arguments,
        [hashtable]$Environment = @{},
        [int]$TimeoutSeconds = $OperationTimeoutSeconds
    )
    $start = [Diagnostics.ProcessStartInfo]::new($Executable)
    $start.WorkingDirectory = $root
    $start.UseShellExecute = $false
    $start.RedirectStandardOutput = $true
    $start.RedirectStandardError = $true
    foreach ($argument in $Arguments) { $start.ArgumentList.Add($argument) }
    foreach ($entry in $Environment.GetEnumerator()) {
        $start.Environment[$entry.Key] = $entry.Value
    }
    $process = [Diagnostics.Process]::new()
    $process.StartInfo = $start
    try {
        if (-not $process.Start()) { throw "Could not start '$Executable'." }
        $stdout = $process.StandardOutput.ReadToEndAsync()
        $stderr = $process.StandardError.ReadToEndAsync()
        if (-not $process.WaitForExit($TimeoutSeconds * 1000)) {
            $process.Kill($true)
            $process.WaitForExit()
            throw "TIMEOUT_UNCERTAIN: '$Executable' exceeded $TimeoutSeconds seconds."
        }
        return [ordered]@{
            exitCode = $process.ExitCode
            stdout = $stdout.GetAwaiter().GetResult()
            stderr = $stderr.GetAwaiter().GetResult()
        }
    }
    finally {
        $process.Dispose()
    }
}

Assert-Confirmed $UserSetupConfirmed.IsPresent 'UserSetupConfirmed' `
    'Confirm the localhost leaf is user-trusted and the task pane is active for this exact workbook.'
Assert-Confirmed $CandidateAllowlistConfirmed.IsPresent 'CandidateAllowlistConfirmed' `
    'Confirm bridge.json was deliberately changed for this acceptance run only.'
Assert-Confirmed $ExcelSlotConfirmed.IsPresent 'ExcelSlotConfirmed' `
    'Confirm the coordinator granted this runner the serialized desktop Excel slot.'

$workbookPath = Resolve-ExactFile $WorkbookPath 'WorkbookPath'
$configPath = Resolve-ExactFile $BridgeConfigPath 'Office.js bridge configuration'
$config = ConvertFrom-StrictJson ([IO.File]::ReadAllText($configPath)) 'bridge.json'
$enabledActions = @($config.enabledActions)
$missingActions = @($requiredActions | Where-Object { $_ -notin $enabledActions })
if ($missingActions.Count -gt 0) {
    throw "bridge.json is missing acceptance action(s): $($missingActions -join ', ')."
}

$receipt = [ordered]@{
    schemaVersion = 1
    status = 'validation-only'
    acceptanceScope = 'public-officejs-pivottable-candidates'
    runtimeProof = $false
    publicCommandAcceptance = $false
    workbookPath = $workbookPath
    bridgeConfigPath = $configPath
    requiredActions = $requiredActions
    entryPoints = @('cli', 'mcp')
}
if ($ValidateOnly) {
    $receipt | ConvertTo-Json -Depth 8
    exit 0
}

if (-not $IsMacOS) {
    throw 'Real Office.js acceptance requires macOS desktop Excel.'
}

if (-not $SkipBuild) {
    $build = Invoke-BoundedProcess -Executable dotnet -Arguments @(
        'build', 'Sbroenne.ExcelMcp.sln', '-c', 'Release',
        '-p:ExcelMcpSkipCleanup=true', '-p:EnableWindowsTargeting=true',
        '--nologo', '-v', 'minimal'
    ) -TimeoutSeconds 300
    if ($build.exitCode -ne 0) {
        throw "Release build failed: $($build.stdout) $($build.stderr)"
    }
}

$cliAssembly = Join-Path $root 'src/ExcelMcp.CLI/bin/Release/net10.0/excelcli.dll'
$mcpAssembly = Join-Path $root 'src/ExcelMcp.McpServer/bin/Release/net10.0/Sbroenne.ExcelMcp.McpServer.dll'
foreach ($assembly in @($cliAssembly, $mcpAssembly)) {
    if (-not [IO.File]::Exists($assembly)) {
        throw "Required built entry point is missing at '$assembly'."
    }
}

$pipe = "em-officejs-$([Guid]::NewGuid().ToString('N'))"
$environment = @{ EXCELMCP_CLI_PIPE = $pipe }
$sessionId = $null
$mcp = $null
$requestId = 0

function Invoke-CliJson {
    param([string[]]$Arguments, [string]$Context)
    $process = Invoke-BoundedProcess `
        -Executable dotnet `
        -Arguments (@($cliAssembly, '-q') + $Arguments) `
        -Environment $environment
    if ($process.exitCode -ne 0) {
        throw "$Context failed: $($process.stdout) $($process.stderr)"
    }
    $result = ConvertFrom-StrictJson $process.stdout $Context
    if ($result.success -ne $true -or $null -ne $result.errorMessage) {
        throw "$Context returned failure: $($process.stdout)"
    }
    return $result
}

function Read-McpResponse {
    param([int]$ExpectedId)
    while ($true) {
        $read = $mcp.StandardOutput.ReadLineAsync()
        if (-not $read.Wait($OperationTimeoutSeconds * 1000)) {
            throw "MCP response $ExpectedId exceeded $OperationTimeoutSeconds seconds."
        }
        $line = $read.Result
        if ($null -eq $line) {
            throw "MCP server closed before response $ExpectedId."
        }
        $response = ConvertFrom-StrictJson $line "MCP response $ExpectedId"
        if ($response.ContainsKey('id') -and [int]$response.id -eq $ExpectedId) {
            if ($response.ContainsKey('error')) {
                throw "MCP response $ExpectedId failed: $line"
            }
            return $response.result
        }
    }
}

function Send-McpRequest {
    param([string]$Method, [hashtable]$Params)
    $script:requestId++
    $message = @{
        jsonrpc = '2.0'
        id = $script:requestId
        method = $Method
        params = $Params
    } | ConvertTo-Json -Compress -Depth 20
    $mcp.StandardInput.WriteLine($message)
    $mcp.StandardInput.Flush()
    return Read-McpResponse $script:requestId
}

function Invoke-McpTool {
    param([string]$Name, [hashtable]$Arguments)
    $result = Send-McpRequest 'tools/call' @{ name = $Name; arguments = $Arguments }
    if ($result.isError -eq $true -or $result.content.Count -lt 1) {
        throw "MCP tool '$Name' returned an error."
    }
    $payload = ConvertFrom-StrictJson ([string]$result.content[0].text) "MCP tool '$Name'"
    if ($payload.success -ne $true -or $null -ne $payload.errorMessage) {
        throw "MCP tool '$Name' returned a failed ExcelMcp result."
    }
    return $payload
}

function Invoke-CliCandidateSequence {
    param([string]$PivotName, [string]$RemovalPivotName)
    $common = @('--session', $sessionId, '--pivot-table-name', $PivotName)
    $data = Invoke-CliJson -Arguments (@('pivottablecalc', 'get-data') + $common) -Context 'CLI get-data'
    if ($data.pivotTableName -cne $PivotName) { throw 'CLI get-data returned the wrong PivotTable.' }
    Invoke-CliJson -Arguments (@('pivottablefield', 'sort-field') + $common + @(
        '--field-name', $RowField, '--direction', 'Descending')) -Context 'CLI sort-field' | Out-Null
    $filter = Invoke-CliJson -Arguments (@('pivottablefield', 'set-field-filter') + $common + @(
        '--field-name', $RowField, '--selected-values', "[`"$SelectedItem`"]")) -Context 'CLI set-field-filter'
    Invoke-CliJson -Arguments (@('pivottablefield', 'set-field-filter') + $common + @(
        '--field-name', $RowField,
        '--selected-values', ($filter.availableItems | ConvertTo-Json -Compress))) -Context 'CLI restore filter' | Out-Null
    Invoke-CliJson -Arguments (@('pivottablefield', 'set-field-format') + $common + @(
        '--field-name', $ValueField, '--number-format', '#,##0.00')) -Context 'CLI set-field-format' | Out-Null
    Invoke-CliJson -Arguments (@('pivottablefield', 'set-field-name') + $common + @(
        '--field-name', $RowField, '--custom-name', "$RowField CLI")) -Context 'CLI set-field-name' | Out-Null
    Invoke-CliJson -Arguments (@('pivottablefield', 'set-field-name') + $common + @(
        '--field-name', $RowField, '--custom-name', $RowField)) -Context 'CLI restore field name' | Out-Null
    Invoke-CliJson -Arguments (@('pivottablecalc', 'set-layout') + $common + @(
        '--row-layout', '1')) -Context 'CLI set-layout' | Out-Null
    Invoke-CliJson -Arguments (@('pivottablecalc', 'set-subtotals') + $common + @(
        '--field-name', $RowField, '--show-subtotals', 'false')) -Context 'CLI set-subtotals' | Out-Null
    Invoke-CliJson -Arguments (@('pivottablecalc', 'set-grand-totals') + $common + @(
        '--show-row-grand-totals', 'false',
        '--show-column-grand-totals', 'true')) -Context 'CLI set-grand-totals' | Out-Null
    Invoke-CliJson -Arguments @(
        'pivottablefield', 'remove-field', '--session', $sessionId,
        '--pivot-table-name', $RemovalPivotName, '--field-name', $RowField
    ) -Context 'CLI remove-field' | Out-Null
}

function Invoke-McpCandidateSequence {
    param([string]$PivotName, [string]$RemovalPivotName)
    $common = @{ session_id = $sessionId; pivot_table_name = $PivotName }
    $data = Invoke-McpTool 'pivottable_calc' ($common + @{ action = 'get-data' })
    if ($data.pivotTableName -cne $PivotName) { throw 'MCP get-data returned the wrong PivotTable.' }
    Invoke-McpTool 'pivottable_field' ($common + @{
        action = 'sort-field'; field_name = $RowField; direction = 'Descending'
    }) | Out-Null
    $filter = Invoke-McpTool 'pivottable_field' ($common + @{
        action = 'set-field-filter'; field_name = $RowField; selected_values = @($SelectedItem)
    })
    Invoke-McpTool 'pivottable_field' ($common + @{
        action = 'set-field-filter'; field_name = $RowField; selected_values = @($filter.availableItems)
    }) | Out-Null
    Invoke-McpTool 'pivottable_field' ($common + @{
        action = 'set-field-format'; field_name = $ValueField; number_format = '#,##0.00'
    }) | Out-Null
    Invoke-McpTool 'pivottable_field' ($common + @{
        action = 'set-field-name'; field_name = $RowField; custom_name = "$RowField MCP"
    }) | Out-Null
    Invoke-McpTool 'pivottable_field' ($common + @{
        action = 'set-field-name'; field_name = $RowField; custom_name = $RowField
    }) | Out-Null
    Invoke-McpTool 'pivottable_calc' ($common + @{
        action = 'set-layout'; row_layout = 2
    }) | Out-Null
    Invoke-McpTool 'pivottable_calc' ($common + @{
        action = 'set-subtotals'; field_name = $RowField; show_subtotals = $true
    }) | Out-Null
    Invoke-McpTool 'pivottable_calc' ($common + @{
        action = 'set-grand-totals'
        show_row_grand_totals = $true
        show_column_grand_totals = $false
    }) | Out-Null
    Invoke-McpTool 'pivottable_field' @{
        action = 'remove-field'
        session_id = $sessionId
        pivot_table_name = $RemovalPivotName
        field_name = $RowField
    } | Out-Null
}

try {
    $open = Invoke-CliJson -Arguments @('session', 'open', $workbookPath) -Context 'CLI session open'
    $sessionId = [string]$open.sessionId
    if ([string]::IsNullOrWhiteSpace($sessionId)) {
        throw 'CLI session open returned no session identity.'
    }
    Invoke-CliCandidateSequence $CliPivotTable $CliRemovalPivotTable

    $start = [Diagnostics.ProcessStartInfo]::new('dotnet')
    $start.WorkingDirectory = $root
    $start.UseShellExecute = $false
    $start.RedirectStandardInput = $true
    $start.RedirectStandardOutput = $true
    $start.RedirectStandardError = $true
    $start.ArgumentList.Add($mcpAssembly)
    foreach ($entry in $environment.GetEnumerator()) {
        $start.Environment[$entry.Key] = $entry.Value
    }
    $mcp = [Diagnostics.Process]::new()
    $mcp.StartInfo = $start
    if (-not $mcp.Start()) { throw 'Could not start the MCP server.' }
    Send-McpRequest 'initialize' @{
        protocolVersion = '2025-06-18'
        capabilities = @{}
        clientInfo = @{ name = 'officejs-acceptance'; version = '1.0' }
    } | Out-Null
    $mcp.StandardInput.WriteLine((@{
        jsonrpc = '2.0'
        method = 'notifications/initialized'
        params = @{}
    } | ConvertTo-Json -Compress))
    $mcp.StandardInput.Flush()
    Invoke-McpCandidateSequence $McpPivotTable $McpRemovalPivotTable

    $receipt.status = 'passed'
    $receipt.runtimeProof = $true
    $receipt.publicCommandAcceptance = $true
    $receipt | ConvertTo-Json -Depth 8
}
finally {
    if ($null -ne $mcp) {
        try { $mcp.StandardInput.Close() } catch {}
        if (-not $mcp.WaitForExit(2000)) {
            $mcp.Kill($true)
            $mcp.WaitForExit()
        }
        $mcp.Dispose()
    }
    if (-not [string]::IsNullOrWhiteSpace($sessionId)) {
        try {
            Invoke-BoundedProcess dotnet @(
                $cliAssembly, '-q', 'session', 'close', '--session', $sessionId
            ) -Environment $environment | Out-Null
        }
        catch {
            Write-Warning "Session cleanup failed: $($_.Exception.Message)"
        }
    }
}
