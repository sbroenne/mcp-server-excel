#Requires -Version 7.0
<#
.SYNOPSIS
Runs guarded public CLI and MCP acceptance for macOS Power Query lifecycle candidates.
.DESCRIPTION
This script never installs the helper, changes Excel security, or creates a
workbook. The caller supplies an original Excel-authored dedicated workbook and
confirms that the exact helper 1.3.0 installation is trusted and an exclusive
desktop Excel slot is available.

Each entry point opens the exact workbook, uses only the public Power Query
surface, verifies exact results and unsupported variants, removes only its
unique owned query, and closes without saving.
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [string]$HelperPath,
    [Parameter(Mandatory)]
    [string]$WorkbookPath,
    [switch]$HelperInstalledTrustedConfirmed,
    [switch]$ExcelAuthoredWorkbookConfirmed,
    [switch]$DedicatedWorkbookConfirmed,
    [switch]$ExcelSlotConfirmed,
    [ValidateRange(10, 120)]
    [int]$OperationTimeoutSeconds = 60,
    [switch]$SkipBuild,
    [switch]$ValidateOnly
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $PSScriptRoot
$helperVersion = '1.3.0'
$acceptanceScope = 'public-powerquery-lifecycle-cli-mcp'
$requiredActions = @(
    'powerquery.create',
    'powerquery.update',
    'powerquery.rename',
    'powerquery.delete',
    'powerquery.refresh',
    'powerquery.refresh-all',
    'powerquery.load-to',
    'powerquery.unload',
    'powerquery.evaluate'
)
$publicMethods = @(
    'powerquery.create',
    'powerquery.list',
    'powerquery.view',
    'powerquery.update',
    'powerquery.rename',
    'powerquery.get-load-config',
    'powerquery.load-to',
    'powerquery.refresh',
    'powerquery.unload',
    'powerquery.delete',
    'powerquery.refresh-all',
    'powerquery.evaluate'
)
$candidateActions = $requiredActions -join ','
$requiredUnsupportedDestinations = @('load-to-data-model', 'load-to-both')
$originalFormula = 'let Source = #table({"Value"}, {{"original"}}) in Source'
$updatedFormula = 'let Source = #table({"Value"}, {{"updated"}}) in Source'
$evaluatedFormula = 'let Source = #table({"Value"}, {{"evaluated"}}) in Source'
$script:uncertain = $false

function Assert-Confirmed {
    param([bool]$Value, [string]$Name, [string]$Meaning)
    if (-not $Value) {
        throw "-$Name is required. $Meaning"
    }
}

function Resolve-ExactInput {
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

Assert-Confirmed $HelperInstalledTrustedConfirmed.IsPresent `
    'HelperInstalledTrustedConfirmed' `
    "Confirm that the exact reviewed ExcelMcpHelper.xlam $helperVersion is already installed and trusted."
Assert-Confirmed $ExcelAuthoredWorkbookConfirmed.IsPresent `
    'ExcelAuthoredWorkbookConfirmed' `
    'Confirm that Excel created and saved the supplied workbook.'
Assert-Confirmed $DedicatedWorkbookConfirmed.IsPresent `
    'DedicatedWorkbookConfirmed' `
    'Confirm that the workbook is dedicated to this destructive acceptance run.'
Assert-Confirmed $ExcelSlotConfirmed.IsPresent `
    'ExcelSlotConfirmed' `
    'Confirm that this run owns the exclusive desktop Excel slot.'

$helperPath = Resolve-ExactInput $HelperPath 'HelperPath'
$workbookPath = Resolve-ExactInput $WorkbookPath 'WorkbookPath'
if ([IO.Path]::GetFileName($helperPath) -cne 'ExcelMcpHelper.xlam') {
    throw 'HelperPath must end with the exact case-sensitive name ExcelMcpHelper.xlam.'
}
if ([IO.Path]::GetExtension($workbookPath) -notin @('.xlsx', '.xlsm')) {
    throw 'WorkbookPath must be a dedicated Excel-authored .xlsx or .xlsm workbook.'
}
if ($helperPath -ceq $workbookPath) {
    throw 'The helper add-in cannot also be the acceptance target workbook.'
}

$validationReceipt = [ordered]@{
    schemaVersion = 1
    status = 'validation-only'
    acceptanceScope = $acceptanceScope
    runtimeProof = $false
    publicCommandAcceptance = $false
    helperVersion = $helperVersion
    helperPath = $helperPath
    workbookPath = $workbookPath
    entryPoints = @('cli', 'mcp')
    requiredActions = $requiredActions
    requiredUnsupportedDestinations = $requiredUnsupportedDestinations
    formulas = 'literal-table-only-no-external-source-or-credentials'
}
if ($ValidateOnly) {
    $validationReceipt | ConvertTo-Json -Depth 10
    exit 0
}

if (-not $IsMacOS) {
    throw 'Real Power Query acceptance requires macOS desktop Excel.'
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

$environment = @{
    EXCELMCP_CLI_PIPE = "em-$([Guid]::NewGuid().ToString('N'))"
    EXCELMCP_MAC_VBA_HELPER_PATH = $helperPath
    EXCELMCP_MAC_POWERQUERY_CANDIDATE_ACTIONS = $candidateActions
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
    param([string]$FileName, [string[]]$Arguments)
    $start = [Diagnostics.ProcessStartInfo]::new()
    $start.FileName = $FileName
    $start.WorkingDirectory = $root
    $start.UseShellExecute = $false
    $start.RedirectStandardOutput = $true
    $start.RedirectStandardError = $true
    foreach ($argument in $Arguments) { $null = $start.ArgumentList.Add($argument) }
    foreach ($entry in $environment.GetEnumerator()) {
        $start.Environment[$entry.Key] = $entry.Value
    }
    $process = [Diagnostics.Process]::new()
    $process.StartInfo = $start
    try {
        $null = $process.Start()
        $stdout = $process.StandardOutput.ReadToEndAsync()
        $stderr = $process.StandardError.ReadToEndAsync()
        if (-not $process.WaitForExit($OperationTimeoutSeconds * 1000)) {
            $process.Kill($true)
            $process.WaitForExit()
            $script:uncertain = $true
            throw "TIMEOUT_UNCERTAIN: $FileName $($Arguments -join ' ')"
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

function Invoke-CliRaw {
    param([string[]]$Arguments)
    $process = Invoke-BoundedProcess dotnet (@($cliAssembly, '-q') + $Arguments)
    $result = ConvertFrom-StrictJson $process.stdout 'CLI command'
    return [ordered]@{
        exitCode = $process.exitCode
        stderr = $process.stderr
        result = $result
    }
}

function Invoke-Cli {
    param([string[]]$Arguments, [string]$Context)
    $response = Invoke-CliRaw $Arguments
    if ($response.exitCode -ne 0 -or $response.result.success -ne $true) {
        throw "$Context failed: $($response.result | ConvertTo-Json -Compress -Depth 20) $($response.stderr)"
    }
    return $response.result
}

function Assert-PublicUnsupported {
    param([hashtable]$Result, [string]$Context)
    if ($Result.success -ne $false -or
        [string]$Result.errorCategory -cne 'PlatformNotSupported') {
        throw "$Context did not fail with PlatformNotSupported: $($Result | ConvertTo-Json -Compress -Depth 20)"
    }
}

function Assert-QueryListed {
    param([hashtable]$Result, [string]$QueryName, [bool]$Expected)
    $found = @($Result.queries | Where-Object { [string]$_.name -ceq $QueryName }).Count -eq 1
    if ($found -ne $Expected) {
        throw "Query list expectation failed for '$QueryName'."
    }
}

function Assert-DedicatedWorkbookIsEmpty {
    param([hashtable]$Result, [string]$Context)
    if (@($Result.queries).Count -ne 0) {
        throw "$Context requires a dedicated workbook with no existing Power Queries."
    }
}

function Assert-View {
    param([hashtable]$Result, [string]$QueryName, [string]$Formula)
    if ([string]$Result.queryName -cne $QueryName -or
        [string]$Result.mCode -cne $Formula -or
        [int]$Result.characterCount -ne $Formula.Length) {
        throw "View result mismatch for '$QueryName'."
    }
}

function Assert-LoadConfig {
    param(
        [hashtable]$Result,
        [string]$QueryName,
        [string]$LoadMode,
        [string]$TargetSheet = ''
    )
    if ([string]$Result.queryName -cne $QueryName -or
        [string]$Result.loadMode -cne $LoadMode) {
        throw "Load configuration mismatch for '$QueryName'."
    }
    if ($LoadMode -ceq 'load-to-table' -and
        [string]$Result.targetSheet -cne $TargetSheet) {
        throw "Load target mismatch for '$QueryName'."
    }
}

function Assert-Refresh {
    param([hashtable]$Result, [string]$QueryName, [string]$SheetName)
    $parsedRefreshTime = [DateTimeOffset]::MinValue
    if ([string]$Result.queryName -cne $QueryName -or
        $Result.hasErrors -ne $false -or
        @($Result.errorMessages).Count -ne 0 -or
        $Result.isConnectionOnly -ne $false -or
        [string]$Result.loadedToSheet -cne $SheetName -or
        -not [DateTimeOffset]::TryParse(
            [string]$Result.refreshTime,
            [ref]$parsedRefreshTime)) {
        throw "Refresh result mismatch for '$QueryName'."
    }
}

function Assert-Rename {
    param(
        [hashtable]$Result,
        [string]$OldName,
        [string]$NewName
    )
    if ([string]$Result.objectType -cne 'power-query' -or
        [string]$Result.oldName -cne $OldName -or
        [string]$Result.newName -cne $NewName -or
        [string]$Result.normalizedOldName -cne $OldName -or
        [string]$Result.normalizedNewName -cne $NewName) {
        throw "Rename result mismatch for '$OldName'."
    }
}

function Assert-Evaluate {
    param([hashtable]$Result)
    if ([string]$Result.mCode -cne $evaluatedFormula -or
        [int]$Result.columnCount -ne 1 -or
        [int]$Result.rowCount -ne 1 -or
        [string]$Result.columns[0] -cne 'Value' -or
        [string]$Result.rows[0][0] -cne 'evaluated') {
        throw 'Evaluate result mismatch.'
    }
}

function Assert-LoadedValues {
    param(
        [hashtable]$Result,
        [string]$ExpectedValue,
        [string]$Context
    )
    $values = @($Result.values)
    if ($values.Count -ne 2 -or
        @($values[0]).Count -ne 1 -or
        @($values[1]).Count -ne 1 -or
        [string]$values[0][0] -cne 'Value' -or
        [string]$values[1][0] -cne $ExpectedValue) {
        throw "$Context loaded values mismatch."
    }
}

function New-WorkingCopy {
    param([string]$EntryPoint)
    $extension = [IO.Path]::GetExtension($workbookPath)
    $path = Join-Path ([IO.Path]::GetTempPath()) (
        "excelmcp-pq-public-$EntryPoint-$([Guid]::NewGuid().ToString('N'))$extension")
    [IO.File]::Copy($workbookPath, $path, $false)
    return $path
}

function Open-CliSession {
    param([string]$Path, [ref]$ExactlyClosed)
    $ExactlyClosed.Value = $false
    $open = Invoke-Cli @(
        'session', 'open', $Path, '--timeout', "$OperationTimeoutSeconds"
    ) 'CLI open'
    $sessionId = [string]$open.sessionId
    if ([string]::IsNullOrWhiteSpace($sessionId)) {
        throw 'CLI open returned no session ID.'
    }
    return $sessionId
}

function Checkpoint-CliSession {
    param([string]$Session, [string]$Path, [ref]$ExactlyClosed)
    $null = Invoke-Cli @(
        'session', 'close', '--session', $Session, '--save', 'true'
    ) 'CLI checkpoint close'
    $ExactlyClosed.Value = $true
    return Open-CliSession $Path $ExactlyClosed
}

function Invoke-CliAcceptance {
    $suffix = [Guid]::NewGuid().ToString('N').Substring(0, 10)
    $queryName = "ExcelMcpCliQuery_$suffix"
    $renamedName = "ExcelMcpCliRenamed_$suffix"
    $sheetName = "ExcelMcpCliSheet_$suffix"
    $unsupportedName = "ExcelMcpCliUnsupported_$suffix"
    $workingPath = New-WorkingCopy 'cli'
    $session = $null
    $exactlyClosed = $true
    $closeError = ''
    try {
        $session = Open-CliSession $workingPath ([ref]$exactlyClosed)
        Assert-DedicatedWorkbookIsEmpty (
            Invoke-Cli @('powerquery', 'list', '--session', $session) 'CLI initial list'
        ) 'CLI acceptance'
        foreach ($destination in $requiredUnsupportedDestinations) {
            $unsupported = Invoke-CliRaw @(
                'powerquery', 'create', '--session', $session,
                '--query-name', $unsupportedName,
                '--m-code', $originalFormula,
                '--load-destination', $destination
            )
            Assert-PublicUnsupported $unsupported.result "CLI unsupported $destination"
        }

        $null = Invoke-Cli @(
            'powerquery', 'create', '--session', $session,
            '--query-name', $queryName,
            '--m-code', $originalFormula,
            '--target-sheet', $sheetName,
            '--target-cell-address', 'A1'
        ) 'CLI create'
        $session = Checkpoint-CliSession $session $workingPath ([ref]$exactlyClosed)
        $loaded = Invoke-Cli @(
            'range', 'get-values', '--session', $session,
            '--sheet', $sheetName, '--range', 'A1:A2'
        ) 'CLI created values'
        Assert-LoadedValues $loaded 'original' 'CLI created checkpoint'
        $list = Invoke-Cli @('powerquery', 'list', '--session', $session) 'CLI list'
        Assert-QueryListed $list $queryName $true
        Assert-View (
            Invoke-Cli @(
                'powerquery', 'view', '--session', $session, '--query-name', $queryName
            ) 'CLI view'
        ) $queryName $originalFormula
        $null = Invoke-Cli @(
            'powerquery', 'update', '--session', $session,
            '--query-name', $queryName, '--m-code', $updatedFormula, '--refresh', 'true'
        ) 'CLI update'
        $session = Checkpoint-CliSession $session $workingPath ([ref]$exactlyClosed)
        $loaded = Invoke-Cli @(
            'range', 'get-values', '--session', $session,
            '--sheet', $sheetName, '--range', 'A1:A2'
        ) 'CLI updated values'
        Assert-LoadedValues $loaded 'updated' 'CLI updated checkpoint'
        Assert-View (
            Invoke-Cli @(
                'powerquery', 'view', '--session', $session, '--query-name', $queryName
            ) 'CLI updated view'
        ) $queryName $updatedFormula
        $rename = Invoke-Cli @(
            'powerquery', 'rename', '--session', $session,
            '--old-name', $queryName, '--new-name', $renamedName
        ) 'CLI rename'
        Assert-Rename $rename $queryName $renamedName
        $session = Checkpoint-CliSession $session $workingPath ([ref]$exactlyClosed)
        Assert-LoadConfig (
            Invoke-Cli @(
                'powerquery', 'get-load-config', '--session', $session,
                '--query-name', $renamedName
            ) 'CLI get load config'
        ) $renamedName 'load-to-table' $sheetName
        $null = Invoke-Cli @(
            'powerquery', 'load-to', '--session', $session,
            '--query-name', $renamedName, '--load-destination', 'connection-only'
        ) 'CLI load connection-only'
        $session = Checkpoint-CliSession $session $workingPath ([ref]$exactlyClosed)
        Assert-LoadConfig (
            Invoke-Cli @(
                'powerquery', 'get-load-config', '--session', $session,
                '--query-name', $renamedName
            ) 'CLI connection-only config'
        ) $renamedName 'connection-only'
        $null = Invoke-Cli @(
            'powerquery', 'load-to', '--session', $session,
            '--query-name', $renamedName, '--load-destination', 'load-to-table',
            '--target-sheet', $sheetName, '--target-cell-address', 'A1'
        ) 'CLI load worksheet'
        Assert-Refresh (
            Invoke-Cli @(
                'powerquery', 'refresh', '--session', $session,
                '--query-name', $renamedName, '--timeout', "$OperationTimeoutSeconds"
            ) 'CLI refresh'
        ) $renamedName $sheetName
        $session = Checkpoint-CliSession $session $workingPath ([ref]$exactlyClosed)
        $loaded = Invoke-Cli @(
            'range', 'get-values', '--session', $session,
            '--sheet', $sheetName, '--range', 'A1:A2'
        ) 'CLI refreshed checkpoint values'
        Assert-LoadedValues $loaded 'updated' 'CLI refreshed checkpoint'
        $null = Invoke-Cli @(
            'powerquery', 'refresh-all', '--session', $session,
            '--timeout', "$OperationTimeoutSeconds"
        ) 'CLI refresh all'
        $session = Checkpoint-CliSession $session $workingPath ([ref]$exactlyClosed)
        $loaded = Invoke-Cli @(
            'range', 'get-values', '--session', $session,
            '--sheet', $sheetName, '--range', 'A1:A2'
        ) 'CLI refresh-all checkpoint values'
        Assert-LoadedValues $loaded 'updated' 'CLI refresh-all checkpoint'
        Assert-Evaluate (
            Invoke-Cli @(
                'powerquery', 'evaluate', '--session', $session, '--m-code', $evaluatedFormula
            ) 'CLI evaluate'
        )
        $null = Invoke-Cli @(
            'powerquery', 'unload', '--session', $session, '--query-name', $renamedName
        ) 'CLI unload'
        $session = Checkpoint-CliSession $session $workingPath ([ref]$exactlyClosed)
        Assert-LoadConfig (
            Invoke-Cli @(
                'powerquery', 'get-load-config', '--session', $session,
                '--query-name', $renamedName
            ) 'CLI unloaded config'
        ) $renamedName 'connection-only'
        $null = Invoke-Cli @(
            'powerquery', 'delete', '--session', $session, '--query-name', $renamedName
        ) 'CLI delete'
        $session = Checkpoint-CliSession $session $workingPath ([ref]$exactlyClosed)
        Assert-QueryListed (
            Invoke-Cli @('powerquery', 'list', '--session', $session) 'CLI final list'
        ) $renamedName $false
    }
    finally {
        if ($null -ne $session) {
            if (-not $script:uncertain) {
                try {
                    $null = Invoke-Cli @(
                        'session', 'close', '--session', $session, '--save', 'false'
                    ) 'CLI close without save'
                    $exactlyClosed = $true
                }
                catch {
                    $closeError = $_.Exception.Message
                }
            }
        }
        if ($exactlyClosed -and [IO.File]::Exists($workingPath)) {
            [IO.File]::Delete($workingPath)
        }
        if (-not $exactlyClosed) {
            throw "RECOVERY_REQUIRED: CLI acceptance could not confirm exact close. Preserve '$workingPath' for manual reconciliation. $closeError"
        }
    }
}

function Start-Mcp {
    $start = [Diagnostics.ProcessStartInfo]::new()
    $start.FileName = 'dotnet'
    $null = $start.ArgumentList.Add($mcpAssembly)
    $start.WorkingDirectory = $root
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
    $process | Add-Member -NotePropertyName AcceptanceStandardErrorTask `
        -NotePropertyValue $process.StandardError.ReadToEndAsync()
    return $process
}

function Read-McpResponse {
    param([Diagnostics.Process]$Process, [int]$Id)
    $timer = [Diagnostics.Stopwatch]::StartNew()
    $budget = [TimeSpan]::FromSeconds($OperationTimeoutSeconds)
    while ($true) {
        $remaining = $budget - $timer.Elapsed
        if ($remaining -le [TimeSpan]::Zero) {
            $script:uncertain = $true
            throw "TIMEOUT_UNCERTAIN: MCP response $Id exceeded $OperationTimeoutSeconds seconds."
        }
        $task = $Process.StandardOutput.ReadLineAsync()
        try {
            $line = $task.WaitAsync($remaining).GetAwaiter().GetResult()
        }
        catch [TimeoutException] {
            $script:uncertain = $true
            throw "TIMEOUT_UNCERTAIN: MCP response $Id exceeded $OperationTimeoutSeconds seconds."
        }
        if ($null -eq $line) { throw "MCP exited before response $Id." }
        $message = ConvertFrom-StrictJson $line 'MCP response'
        if ($message.id -eq $Id) { return $message }
    }
}

function Write-McpMessage {
    param([Diagnostics.Process]$Process, [hashtable]$Message)
    $Process.StandardInput.WriteLine(($Message | ConvertTo-Json -Compress -Depth 30))
    $Process.StandardInput.Flush()
}

function Invoke-McpToolRaw {
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
    return ConvertFrom-StrictJson ([string]$response.result.content[0].text) "MCP $Name result"
}

function Invoke-McpTool {
    param(
        [Diagnostics.Process]$Process,
        [ref]$NextId,
        [string]$Name,
        [hashtable]$Arguments,
        [string]$Context
    )
    $result = Invoke-McpToolRaw $Process $NextId $Name $Arguments
    if ($result.success -ne $true) {
        throw "$Context failed: $($result | ConvertTo-Json -Compress -Depth 20)"
    }
    return $result
}

function Open-McpSession {
    param(
        [Diagnostics.Process]$Process,
        [ref]$NextId,
        [string]$Path,
        [ref]$ExactlyClosed
    )
    $ExactlyClosed.Value = $false
    $open = Invoke-McpTool $Process $NextId 'file' @{
        action = 'open'
        path = $Path
        show = $false
        timeout_seconds = $OperationTimeoutSeconds
    } 'MCP open'
    $sessionId = [string]$open.sessionId
    if ([string]::IsNullOrWhiteSpace($sessionId)) {
        throw 'MCP open returned no session ID.'
    }
    return $sessionId
}

function Checkpoint-McpSession {
    param(
        [Diagnostics.Process]$Process,
        [ref]$NextId,
        [string]$Session,
        [string]$Path,
        [ref]$ExactlyClosed
    )
    $null = Invoke-McpTool $Process $NextId 'file' @{
        action = 'close'; session_id = $Session; save = $true
    } 'MCP checkpoint close'
    $ExactlyClosed.Value = $true
    return Open-McpSession $Process $NextId $Path $ExactlyClosed
}

function Invoke-McpAcceptance {
    $suffix = [Guid]::NewGuid().ToString('N').Substring(0, 10)
    $queryName = "ExcelMcpMcpQuery_$suffix"
    $renamedName = "ExcelMcpMcpRenamed_$suffix"
    $sheetName = "ExcelMcpMcpSheet_$suffix"
    $unsupportedName = "ExcelMcpMcpUnsupported_$suffix"
    $workingPath = New-WorkingCopy 'mcp'
    $process = Start-Mcp
    $nextId = 1
    $session = $null
    $exactlyClosed = $true
    $closeError = ''
    try {
        Write-McpMessage $process @{
            jsonrpc = '2.0'
            id = $nextId
            method = 'initialize'
            params = @{
                protocolVersion = '2025-06-18'
                capabilities = @{}
                clientInfo = @{ name = 'excelmcp-public-powerquery-acceptance'; version = '1.0' }
            }
        }
        $null = Read-McpResponse $process $nextId
        $nextId++
        Write-McpMessage $process @{
            jsonrpc = '2.0'
            method = 'notifications/initialized'
            params = @{}
        }

        $session = Open-McpSession $process ([ref]$nextId) $workingPath ([ref]$exactlyClosed)
        Assert-DedicatedWorkbookIsEmpty (
            Invoke-McpTool $process ([ref]$nextId) 'powerquery' @{
                action = 'list'; session_id = $session
            } 'MCP initial list'
        ) 'MCP acceptance'

        foreach ($destination in $requiredUnsupportedDestinations) {
            $unsupported = Invoke-McpToolRaw $process ([ref]$nextId) 'powerquery' @{
                action = 'create'
                session_id = $session
                query_name = $unsupportedName
                m_code = $originalFormula
                load_destination = $destination
            }
            Assert-PublicUnsupported $unsupported "MCP unsupported $destination"
        }

        $null = Invoke-McpTool $process ([ref]$nextId) 'powerquery' @{
            action = 'create'
            session_id = $session
            query_name = $queryName
            m_code = $originalFormula
            target_sheet = $sheetName
            target_cell_address = 'A1'
        } 'MCP create'
        $session = Checkpoint-McpSession $process ([ref]$nextId) $session $workingPath `
            ([ref]$exactlyClosed)
        $loaded = Invoke-McpTool $process ([ref]$nextId) 'range' @{
            action = 'get-values'; session_id = $session
            sheet_name = $sheetName; range_address = 'A1:A2'
        } 'MCP created values'
        Assert-LoadedValues $loaded 'original' 'MCP created checkpoint'
        $list = Invoke-McpTool $process ([ref]$nextId) 'powerquery' @{
            action = 'list'; session_id = $session
        } 'MCP list'
        Assert-QueryListed $list $queryName $true
        Assert-View (
            Invoke-McpTool $process ([ref]$nextId) 'powerquery' @{
                action = 'view'; session_id = $session; query_name = $queryName
            } 'MCP view'
        ) $queryName $originalFormula
        $null = Invoke-McpTool $process ([ref]$nextId) 'powerquery' @{
            action = 'update'
            session_id = $session
            query_name = $queryName
            m_code = $updatedFormula
            refresh = $true
        } 'MCP update'
        $session = Checkpoint-McpSession $process ([ref]$nextId) $session $workingPath `
            ([ref]$exactlyClosed)
        $loaded = Invoke-McpTool $process ([ref]$nextId) 'range' @{
            action = 'get-values'; session_id = $session
            sheet_name = $sheetName; range_address = 'A1:A2'
        } 'MCP updated values'
        Assert-LoadedValues $loaded 'updated' 'MCP updated checkpoint'
        Assert-View (
            Invoke-McpTool $process ([ref]$nextId) 'powerquery' @{
                action = 'view'; session_id = $session; query_name = $queryName
            } 'MCP updated view'
        ) $queryName $updatedFormula
        $rename = Invoke-McpTool $process ([ref]$nextId) 'powerquery' @{
            action = 'rename'
            session_id = $session
            old_name = $queryName
            new_name = $renamedName
        } 'MCP rename'
        Assert-Rename $rename $queryName $renamedName
        $session = Checkpoint-McpSession $process ([ref]$nextId) $session $workingPath `
            ([ref]$exactlyClosed)
        Assert-LoadConfig (
            Invoke-McpTool $process ([ref]$nextId) 'powerquery' @{
                action = 'get-load-config'; session_id = $session; query_name = $renamedName
            } 'MCP get load config'
        ) $renamedName 'load-to-table' $sheetName
        $null = Invoke-McpTool $process ([ref]$nextId) 'powerquery' @{
            action = 'load-to'
            session_id = $session
            query_name = $renamedName
            load_destination = 'connection-only'
        } 'MCP load connection-only'
        $session = Checkpoint-McpSession $process ([ref]$nextId) $session $workingPath `
            ([ref]$exactlyClosed)
        Assert-LoadConfig (
            Invoke-McpTool $process ([ref]$nextId) 'powerquery' @{
                action = 'get-load-config'; session_id = $session; query_name = $renamedName
            } 'MCP connection-only config'
        ) $renamedName 'connection-only'
        $null = Invoke-McpTool $process ([ref]$nextId) 'powerquery' @{
            action = 'load-to'
            session_id = $session
            query_name = $renamedName
            load_destination = 'load-to-table'
            target_sheet = $sheetName
            target_cell_address = 'A1'
        } 'MCP load worksheet'
        Assert-Refresh (
            Invoke-McpTool $process ([ref]$nextId) 'powerquery' @{
                action = 'refresh'
                session_id = $session
                query_name = $renamedName
                timeout = $OperationTimeoutSeconds
            } 'MCP refresh'
        ) $renamedName $sheetName
        $session = Checkpoint-McpSession $process ([ref]$nextId) $session $workingPath `
            ([ref]$exactlyClosed)
        $loaded = Invoke-McpTool $process ([ref]$nextId) 'range' @{
            action = 'get-values'; session_id = $session
            sheet_name = $sheetName; range_address = 'A1:A2'
        } 'MCP refreshed checkpoint values'
        Assert-LoadedValues $loaded 'updated' 'MCP refreshed checkpoint'
        $null = Invoke-McpTool $process ([ref]$nextId) 'powerquery' @{
            action = 'refresh-all'; session_id = $session; timeout = $OperationTimeoutSeconds
        } 'MCP refresh all'
        $session = Checkpoint-McpSession $process ([ref]$nextId) $session $workingPath `
            ([ref]$exactlyClosed)
        $loaded = Invoke-McpTool $process ([ref]$nextId) 'range' @{
            action = 'get-values'; session_id = $session
            sheet_name = $sheetName; range_address = 'A1:A2'
        } 'MCP refresh-all checkpoint values'
        Assert-LoadedValues $loaded 'updated' 'MCP refresh-all checkpoint'
        Assert-Evaluate (
            Invoke-McpTool $process ([ref]$nextId) 'powerquery' @{
                action = 'evaluate'; session_id = $session; m_code = $evaluatedFormula
            } 'MCP evaluate'
        )
        $null = Invoke-McpTool $process ([ref]$nextId) 'powerquery' @{
            action = 'unload'; session_id = $session; query_name = $renamedName
        } 'MCP unload'
        $session = Checkpoint-McpSession $process ([ref]$nextId) $session $workingPath `
            ([ref]$exactlyClosed)
        Assert-LoadConfig (
            Invoke-McpTool $process ([ref]$nextId) 'powerquery' @{
                action = 'get-load-config'; session_id = $session; query_name = $renamedName
            } 'MCP unloaded config'
        ) $renamedName 'connection-only'
        $null = Invoke-McpTool $process ([ref]$nextId) 'powerquery' @{
            action = 'delete'; session_id = $session; query_name = $renamedName
        } 'MCP delete'
        $session = Checkpoint-McpSession $process ([ref]$nextId) $session $workingPath `
            ([ref]$exactlyClosed)
        Assert-QueryListed (
            Invoke-McpTool $process ([ref]$nextId) 'powerquery' @{
                action = 'list'; session_id = $session
            } 'MCP final list'
        ) $renamedName $false
    }
    finally {
        if ($null -ne $session -and -not $process.HasExited) {
            if (-not $script:uncertain) {
                try {
                    $null = Invoke-McpTool $process ([ref]$nextId) 'file' @{
                        action = 'close'; session_id = $session; save = $false
                    } 'MCP close without save'
                    $exactlyClosed = $true
                }
                catch {
                    $closeError = $_.Exception.Message
                }
            }
        }
        if (-not $process.HasExited) {
            $process.StandardInput.Close()
            if (-not $process.WaitForExit(5000)) {
                $process.Kill($true)
                $process.WaitForExit()
            }
        }
        $mcpStandardError = $process.AcceptanceStandardErrorTask.GetAwaiter().GetResult()
        $process.Dispose()
        if ($exactlyClosed -and [IO.File]::Exists($workingPath)) {
            [IO.File]::Delete($workingPath)
        }
        if (-not $exactlyClosed) {
            throw "RECOVERY_REQUIRED: MCP acceptance could not confirm exact close. Preserve '$workingPath' for manual reconciliation. $closeError $mcpStandardError"
        }
    }
}

function Stop-PrivateCliDaemon {
    $null = Invoke-Cli @('service', 'stop') 'Private CLI daemon cleanup'
}

try {
    Invoke-CliAcceptance
}
finally {
    Stop-PrivateCliDaemon
}
Invoke-McpAcceptance

[ordered]@{
    schemaVersion = 1
    status = 'passed'
    acceptanceScope = $acceptanceScope
    runtimeProof = $true
    publicCommandAcceptance = $true
    helperVersion = $helperVersion
    helperPath = $helperPath
    workbookPath = $workbookPath
    entryPoints = @('cli', 'mcp')
    actions = $publicMethods
    candidateActions = $requiredActions
    methodsByEntryPoint = @{
        cli = $publicMethods
        mcp = $publicMethods
    }
    unsupportedDestinationsVerified = $requiredUnsupportedDestinations
    formulas = 'literal-table-only-no-external-source-or-credentials'
    cleanup = 'owned-working-copy-deleted-after-exact-close-without-save'
} | ConvertTo-Json -Depth 10
