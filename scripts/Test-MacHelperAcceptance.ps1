#Requires -Version 7.0
<#
.SYNOPSIS
Runs opt-in direct-engine acceptance for an existing ExcelMcpHelper installation.
.DESCRIPTION
This runner does not install the helper, approve macros, change VBA project-model
trust, or enable public helper-backed commands. The user must first review and
import the committed ExcelMcpHelper.bas into a blank workbook, save it as the
exact ExcelMcpHelper.xlam path, approve that helper in Excel, and create a
dedicated Excel-authored .xlsm test workbook.

The runner opens the exact test workbook through the public CLI session path,
then exercises the fixed helper.dispatch transport in both the CLI and MCP
binaries. Passing proves only direct helper-engine acceptance. It does not prove
that currently gated public CLI/MCP Power Query or VBA commands are enabled.

Example:
  pwsh ./scripts/Test-MacHelperAcceptance.ps1 `
    -HelperPath '/absolute/path/ExcelMcpHelper.xlam' `
    -WorkbookPath '/absolute/path/ExcelMcpHelperAcceptance.xlsm' `
    -MacroApprovalConfirmed -VbaProjectTrustConfirmed `
    -ExcelAuthoredWorkbookConfirmed
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
    [switch]$SkipBuild,
    [switch]$ValidateOnly,
    [ValidateRange(5, 120)]
    [int]$OperationTimeoutSeconds = 30
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $PSScriptRoot
$helperSourcePath = Join-Path $root 'src/ExcelMcp.Service/Mac/ExcelMcpHelper.bas'
$maximumPayloadBytes = 262144
$queryName = 'ExcelMcpFixtureLiteral'
$renamedQueryName = 'ExcelMcpFixtureLiteralRenamed'
$moduleName = 'ExcelMcpFixtureModule'
$queryFormula = 'let Source = #table({"Value"}, {{"fixture"}}) in Source'
$updatedQueryFormula = 'let Source = #table({"Value"}, {{"updated"}}) in Source'
$moduleSource = "Option Explicit`n`nPublic Function ExcelMcpFixtureValue() As String`n    ExcelMcpFixtureValue = `"fixture`"`nEnd Function"
$updatedModuleSource = "Option Explicit`n`nPublic Function ExcelMcpFixtureValue() As String`n    ExcelMcpFixtureValue = `"updated`"`nEnd Function"
$requiredActions = @(
    'helper.capabilities',
    'helper.inspect-engines',
    'powerquery.list',
    'powerquery.view',
    'powerquery.create',
    'powerquery.update',
    'powerquery.rename',
    'powerquery.delete',
    'powerquery.refresh',
    'powerquery.refresh-all',
    'powerquery.load-to',
    'powerquery.unload',
    'powerquery.evaluate',
    'analysis.create-scenario',
    'analysis.show-scenario',
    'vba.list',
    'vba.view',
    'vba.import',
    'vba.update',
    'vba.delete',
    'vba.run'
)
$engineStatuses = @('accessible', 'unavailable', 'error', 'unknown')
$engineReasonCodes = @(
    'api_access_only',
    'no_collection_object',
    'no_model_object_observed',
    'no_model_objects_observed',
    'api_not_exposed',
    'probe_failed'
)
$requiredProvenMethods = @(
    'powerQueryList',
    'powerQueryCreate',
    'powerQueryUpdate',
    'powerQueryRename',
    'powerQueryDelete',
    'powerQueryRefresh',
    'powerQueryRefreshAll',
    'powerQueryLoadTo',
    'powerQueryUnload',
    'powerQueryEvaluate',
    'xmlXPathRead',
    'dataModelRead',
    'scenarioCreateShow',
    'vbaListView',
    'vbaMutation',
    'vbaRun'
)
$requiredEngineCapabilities = @(
    'xmlMapsApi',
    'rangeXPathApi',
    'workbookModelApi',
    'dataModelConnectionApi'
)

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

Assert-Confirmed $MacroApprovalConfirmed.IsPresent 'MacroApprovalConfirmed' `
    'Confirm that the exact reviewed helper has already received user-managed macro approval.'
Assert-Confirmed $VbaProjectTrustConfirmed.IsPresent 'VbaProjectTrustConfirmed' `
    'Confirm that the user has already enabled project-model trust for this controlled acceptance run.'
Assert-Confirmed $ExcelAuthoredWorkbookConfirmed.IsPresent 'ExcelAuthoredWorkbookConfirmed' `
    'Confirm that Excel created and saved the dedicated test .xlsm; synthetic vbaProject.bin is forbidden.'

$helperPath = Resolve-ExactInput $HelperPath 'HelperPath'
$workbookPath = Resolve-ExactInput $WorkbookPath 'WorkbookPath'
if ([IO.Path]::GetFileName($helperPath) -cne 'ExcelMcpHelper.xlam') {
    throw 'HelperPath must end with the exact case-sensitive name ExcelMcpHelper.xlam.'
}
if ([IO.Path]::GetExtension($workbookPath) -ine '.xlsm') {
    throw 'WorkbookPath must be a dedicated Excel-authored .xlsm test workbook.'
}
if ($helperPath -ceq $workbookPath) {
    throw 'The helper add-in cannot also be the acceptance target workbook.'
}

$validationReceipt = [ordered]@{
    schemaVersion = 1
    status = 'validation-only'
    acceptanceScope = 'direct-helper-engine'
    runtimeProof = $false
    publicCommandAcceptance = $false
    helperPath = $helperPath
    workbookPath = $workbookPath
    helperSourcePath = [IO.Path]::GetFullPath($helperSourcePath)
    helperSourceExists = [IO.File]::Exists($helperSourcePath)
    entryPoints = @('cli', 'mcp')
    phases = @('helper.capabilities', 'helper.inspect-engines', 'powerquery', 'vba')
}
if ($ValidateOnly) {
    $validationReceipt | ConvertTo-Json -Depth 8
    exit 0
}

if (-not $IsMacOS) {
    throw 'Real helper acceptance requires macOS desktop Excel.'
}
if (-not [IO.File]::Exists($helperSourcePath)) {
    throw "The original committed helper source is missing at '$helperSourcePath'. Integrate #915 before running acceptance."
}
$helperSource = [IO.File]::ReadAllText($helperSourcePath)
if ($helperSource -notmatch 'Public Function ExcelMcpDispatch\(ByVal requestJson As String\) As String') {
    throw 'The committed helper source does not expose the fixed ExcelMcpDispatch entry point.'
}
$helperSourceSha256 = (Get-FileHash -Algorithm SHA256 -LiteralPath $helperSourcePath).Hash.ToLowerInvariant()

. (Join-Path $PSScriptRoot 'spikes/macos/MacTestEnvironment.ps1')
Assert-MacAutomationAllowed

function Invoke-BoundedProcess {
    param(
        [string]$Executable,
        [string[]]$Arguments,
        [string]$StandardInput = '',
        [hashtable]$Environment = @{},
        [int]$TimeoutSeconds = $OperationTimeoutSeconds
    )
    $start = [Diagnostics.ProcessStartInfo]::new($Executable)
    $start.WorkingDirectory = $root
    $start.UseShellExecute = $false
    $start.RedirectStandardInput = $true
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
        if (-not [string]::IsNullOrEmpty($StandardInput)) {
            $process.StandardInput.Write($StandardInput)
        }
        $process.StandardInput.Close()
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

if (-not $SkipBuild) {
    $build = Invoke-BoundedProcess dotnet @(
        'build',
        'Sbroenne.ExcelMcp.sln',
        '-c',
        'Release',
        '-p:EnableWindowsTargeting=true',
        '--nologo',
        '-v',
        'minimal'
    ) '' @{} 300
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

$pipe = "em-helper-$([Guid]::NewGuid().ToString('N'))"
$environment = @{
    EXCELMCP_CLI_PIPE = $pipe
    EXCELMCP_MAC_VBA_HELPER_PATH = $helperPath
}
$sessionId = $null
$sessionInvalidated = $false
$uncertain = $false
$manualReconciliationRequired = $false
$cleanupAttempted = $false
$cleanupSucceeded = $false
$queryCreated = $false
$renamedQueryCreated = $false
$moduleCreated = $false
$results = [Collections.Generic.List[object]]::new()

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

function Invoke-HelperAction {
    param(
        [ValidateSet('cli', 'mcp')]
        [string]$EntryPoint,
        [string]$Action,
        [hashtable]$Arguments
    )
    $requestId = [Guid]::NewGuid().ToString('N')
    $request = [ordered]@{
        version = 1
        requestId = $requestId
        workbookPath = $workbookPath
        action = $Action
        arguments = $Arguments
    }
    $requestJson = $request | ConvertTo-Json -Compress -Depth 20
    if ([Text.Encoding]::UTF8.GetByteCount($requestJson) -gt $maximumPayloadBytes) {
        throw "Helper request for '$Action' exceeds $maximumPayloadBytes UTF-8 bytes."
    }
    $transportInput = @{
        helperPath = $helperPath
        requestJson = $requestJson
    } | ConvertTo-Json -Compress -Depth 20
    $assembly = if ($EntryPoint -eq 'cli') { $cliAssembly } else { $mcpAssembly }
    $process = Invoke-BoundedProcess dotnet @(
        $assembly,
        '--excelmcp-mac-automation',
        'helper.dispatch'
    ) $transportInput $environment
    if ($process.exitCode -ne 0) {
        throw "$EntryPoint helper transport exited $($process.exitCode): $($process.stderr)"
    }
    $outer = ConvertFrom-StrictJson $process.stdout "$EntryPoint helper transport"
    if ($outer.success -ne $true -or [string]::IsNullOrWhiteSpace([string]$outer.responseJson)) {
        throw "$EntryPoint helper transport failed: $($process.stdout)"
    }
    $responseJson = [string]$outer.responseJson
    if ([Text.Encoding]::UTF8.GetByteCount($responseJson) -gt $maximumPayloadBytes) {
        throw "$EntryPoint helper response exceeds $maximumPayloadBytes UTF-8 bytes."
    }
    $response = ConvertFrom-StrictJson $responseJson "$EntryPoint helper response"
    if ($response.version -ne 1 -or $response.requestId -cne $requestId) {
        throw "$EntryPoint helper response identity did not match the request."
    }
    if ($response.success -eq $true) {
        if ($null -ne $response.error -or $null -eq $response.result) {
            throw "$EntryPoint helper success response has an invalid result/error shape."
        }
        return $response.result
    }
    if (($null -ne $response.result) -or
        ($null -eq $response.error) -or
        [string]::IsNullOrWhiteSpace([string]$response.error.category) -or
        [string]::IsNullOrWhiteSpace([string]$response.error.code) -or
        [string]::IsNullOrWhiteSpace([string]$response.error.message)) {
        throw "$EntryPoint helper failure response has an invalid result/error shape."
    }
        if (($response.error.category -ceq 'RecoveryRequired') -or
            ($response.error.code -ceq 'rollback_failed')) {
            throw "RECOVERY_REQUIRED: $EntryPoint helper action '$Action' failed " +
                "[$($response.error.category)/$($response.error.code)]: $($response.error.message)"
        }
        throw "$EntryPoint helper action '$Action' failed [$($response.error.category)/$($response.error.code)]: $($response.error.message)"
}

function Assert-Capabilities {
    param([string]$EntryPoint, [hashtable]$Capabilities)
    if ($Capabilities.helperVersion -cne '1.3.0' -or $Capabilities.protocolVersion -ne 1) {
        throw "$EntryPoint helper version does not match protocol 1 / helper 1.3.0."
    }
    $supportedActions = @($Capabilities.supportedActions)
    if ($supportedActions.Count -ne $requiredActions.Count -or
        (Compare-Object $requiredActions $supportedActions -CaseSensitive).Count -ne 0) {
        throw "$EntryPoint helper supportedActions does not match the exact 1.3.0 contract."
    }
    if ($Capabilities.trustReadiness.powerQueryReadable -ne $true) {
        throw "$EntryPoint helper reports Power Query live access is not ready."
    }
    if ($Capabilities.trustReadiness.vbaProjectReadable -ne $true) {
        throw "$EntryPoint helper reports VBA project-model trust is not ready."
    }
    foreach ($capability in $requiredEngineCapabilities) {
        if (-not $Capabilities.engineCapabilities.ContainsKey($capability) -or
            $null -ne $Capabilities.engineCapabilities[$capability]) {
            throw "$EntryPoint helper engine capability '$capability' must remain null until separately proven."
        }
    }
    foreach ($method in $requiredProvenMethods) {
        if (-not $Capabilities.provenMethods.ContainsKey($method) -or
            $Capabilities.provenMethods[$method] -ne $false) {
            throw "$EntryPoint helper method '$method' must remain unproven before this direct-engine run."
        }
    }
}

function Assert-EngineInspection {
    param([string]$EntryPoint, [hashtable]$Inspection)
    $inspectionKeys = @($Inspection.Keys)
    if ($inspectionKeys.Count -ne 2 -or
        -not $Inspection.ContainsKey('xmlMaps') -or
        -not $Inspection.ContainsKey('workbookModel')) {
        throw "$EntryPoint helper engine inspection does not have the exact xmlMaps/workbookModel shape."
    }
    foreach ($engineName in @('xmlMaps', 'workbookModel')) {
        $engine = $Inspection[$engineName]
        $keys = @($engine.Keys)
        if ($keys.Count -ne 4 -or
            -not $engine.ContainsKey('status') -or
            -not $engine.ContainsKey('apiAccessible') -or
            -not $engine.ContainsKey('objectCount') -or
            -not $engine.ContainsKey('reasonCode')) {
            throw "$EntryPoint helper engine '$engineName' has an invalid result shape."
        }
        if ($engineStatuses -cnotcontains [string]$engine.status -or
            $engineReasonCodes -cnotcontains [string]$engine.reasonCode -or
            $engine.apiAccessible -isnot [bool] -or
            ($null -ne $engine.objectCount -and $engine.objectCount -isnot [long])) {
            throw "$EntryPoint helper engine '$engineName' returned an invalid observation."
        }
        if (($engine.status -ceq 'unknown' -and $engine.apiAccessible -ne $true) -or
            ($engine.status -in @('unavailable', 'error') -and $engine.apiAccessible -ne $false)) {
            throw "$EntryPoint helper engine '$engineName' returned inconsistent status/accessibility."
        }
    }
}

function Assert-FixtureNamesAbsent {
    param([string]$EntryPoint)
    $queries = Invoke-HelperAction $EntryPoint 'powerquery.list' @{}
    if ((@($queries.queries | ForEach-Object { $_.name }) -ccontains $queryName) -or
        (@($queries.queries | ForEach-Object { $_.name }) -ccontains $renamedQueryName)) {
        throw "$EntryPoint found a reserved fixture query before acceptance; refusing to delete pre-existing data."
    }
    $modules = Invoke-HelperAction $EntryPoint 'vba.list' @{}
    if (@($modules.modules | ForEach-Object { $_.name }) -ccontains $moduleName) {
        throw "$EntryPoint found the reserved fixture module before acceptance; refusing to delete pre-existing code."
    }
}

function Invoke-PowerQueryLifecycle {
    param([string]$EntryPoint)
    $null = Invoke-HelperAction $EntryPoint 'powerquery.create' @{
        name = $queryName
        formula = $queryFormula
        destination = 'connection-only'
        sheetName = $null
        cellAddress = $null
    }
    $script:queryCreated = $true
    $list = Invoke-HelperAction $EntryPoint 'powerquery.list' @{}
    if (@($list.queries | ForEach-Object { $_.name }) -cnotcontains $queryName) {
        throw "$EntryPoint did not list the created literal query."
    }
    $view = Invoke-HelperAction $EntryPoint 'powerquery.view' @{ name = $queryName }
    if ($view.name -cne $queryName -or $view.formula -cne $queryFormula) {
        throw "$EntryPoint did not return the exact created query identity and formula."
    }
    $null = Invoke-HelperAction $EntryPoint 'powerquery.update' @{
        name = $queryName
        formula = $updatedQueryFormula
        refresh = $false
    }
    $null = Invoke-HelperAction $EntryPoint 'powerquery.rename' @{
        name = $queryName
        newName = $renamedQueryName
    }
    $script:queryCreated = $false
    $script:renamedQueryCreated = $true
    $view = Invoke-HelperAction $EntryPoint 'powerquery.view' @{ name = $renamedQueryName }
    if ($view.name -cne $renamedQueryName -or $view.formula -cne $updatedQueryFormula) {
        throw "$EntryPoint did not return the exact renamed query identity and updated formula."
    }
    $null = Invoke-HelperAction $EntryPoint 'powerquery.delete' @{
        name = $renamedQueryName
        deleteConnection = $true
    }
    $script:renamedQueryCreated = $false
    $list = Invoke-HelperAction $EntryPoint 'powerquery.list' @{}
    if (@($list.queries | ForEach-Object { $_.name }) -ccontains $renamedQueryName) {
        throw "$EntryPoint still listed the deleted fixture query."
    }
}

function Invoke-VbaLifecycle {
    param([string]$EntryPoint)
    $null = Invoke-HelperAction $EntryPoint 'vba.import' @{
        moduleName = $moduleName
        source = $moduleSource
    }
    $script:moduleCreated = $true
    $list = Invoke-HelperAction $EntryPoint 'vba.list' @{}
    if (@($list.modules | ForEach-Object { $_.name }) -cnotcontains $moduleName) {
        throw "$EntryPoint did not list the imported fixture module."
    }
    $view = Invoke-HelperAction $EntryPoint 'vba.view' @{ moduleName = $moduleName }
    if ($view.moduleName -cne $moduleName -or $view.source -cne $moduleSource) {
        throw "$EntryPoint did not return the exact imported module source."
    }
    $null = Invoke-HelperAction $EntryPoint 'vba.update' @{
        moduleName = $moduleName
        source = $updatedModuleSource
    }
    $view = Invoke-HelperAction $EntryPoint 'vba.view' @{ moduleName = $moduleName }
    if ($view.moduleName -cne $moduleName -or $view.source -cne $updatedModuleSource) {
        throw "$EntryPoint did not return the exact updated module source."
    }
    $null = Invoke-HelperAction $EntryPoint 'vba.delete' @{ moduleName = $moduleName }
    $script:moduleCreated = $false
    $list = Invoke-HelperAction $EntryPoint 'vba.list' @{}
    if (@($list.modules | ForEach-Object { $_.name }) -ccontains $moduleName) {
        throw "$EntryPoint still listed the deleted fixture module."
    }
}

function Invoke-KnownStateCleanup {
    param([string]$EntryPoint)
    $script:cleanupAttempted = $true
    $ownedQueryNames = [Collections.Generic.List[string]]::new()
    if ($script:renamedQueryCreated) { $ownedQueryNames.Add($renamedQueryName) }
    if ($script:queryCreated) { $ownedQueryNames.Add($queryName) }
    foreach ($name in $ownedQueryNames) {
        try {
            $null = Invoke-HelperAction $EntryPoint 'powerquery.delete' @{
                name = $name
                deleteConnection = $true
            }
            if ($name -ceq $renamedQueryName) { $script:renamedQueryCreated = $false }
            if ($name -ceq $queryName) { $script:queryCreated = $false }
        }
        catch {
            if ($_.Exception.Message -match 'TIMEOUT_UNCERTAIN') { throw }
        }
    }
    if ($script:moduleCreated) {
        try {
            $null = Invoke-HelperAction $EntryPoint 'vba.delete' @{ moduleName = $moduleName }
            $script:moduleCreated = $false
        }
        catch {
            if ($_.Exception.Message -match 'TIMEOUT_UNCERTAIN') { throw }
        }
    }
}

function Invalidate-RunnerSession {
    try {
        $stop = Invoke-BoundedProcess dotnet @(
            $cliAssembly,
            '-q',
            'service',
            'stop'
        ) '' $environment 10
        if ($stop.exitCode -ne 0) { return $false }
        $stopResult = ConvertFrom-StrictJson $stop.stdout 'CLI private daemon stop'
        if ($stopResult.success -ne $true) { return $false }
        $script:sessionId = $null
        return $true
    }
    catch {
        return $false
    }
}

function Close-TestSessionWithoutSaving {
    if ([string]::IsNullOrWhiteSpace($script:sessionId)) { return }
    $close = Invoke-BoundedProcess dotnet @(
        $cliAssembly,
        '-q',
        'session',
        'close',
        '--session',
        $script:sessionId
    ) '' $environment
    if ($close.exitCode -ne 0) {
        throw "CLI session close failed: $($close.stdout) $($close.stderr)"
    }
    $closeResult = ConvertFrom-StrictJson $close.stdout 'CLI session close'
    if ($closeResult.success -ne $true) {
        throw "CLI session close failed: $($close.stdout)"
    }
    $script:sessionId = $null
}

$failureMessage = $null
try {
    $open = Invoke-BoundedProcess dotnet @(
        $cliAssembly,
        '-q',
        'session',
        'open',
        $workbookPath,
        '--timeout',
        '15'
    ) '' $environment 30
    if ($open.exitCode -ne 0) {
        throw "CLI session open failed: $($open.stdout) $($open.stderr)"
    }
    $openResult = ConvertFrom-StrictJson $open.stdout 'CLI session open'
    if ($openResult.success -ne $true) {
        throw "CLI session open failed: $($open.stdout)"
    }
    $sessionId = if ($openResult.ContainsKey('sessionId')) {
        [string]$openResult.sessionId
    } else {
        [string]$openResult.session_id
    }
    if ([string]::IsNullOrWhiteSpace($sessionId)) {
        throw 'CLI session open returned no session identity.'
    }

    foreach ($entryPoint in @('cli', 'mcp')) {
        $capabilities = Invoke-HelperAction $entryPoint 'helper.capabilities' @{}
        Assert-Capabilities $entryPoint $capabilities
        $engineInspection = Invoke-HelperAction $entryPoint 'helper.inspect-engines' @{}
        Assert-EngineInspection $entryPoint $engineInspection
        Assert-FixtureNamesAbsent $entryPoint
        Invoke-PowerQueryLifecycle $entryPoint
        Invoke-VbaLifecycle $entryPoint
        $results.Add([ordered]@{
            entryPoint = $entryPoint
            preflight = 'passed'
            capabilitySchemaAccepted = $true
            reportedProvenMethods = $capabilities.provenMethods
            engineInspection = $engineInspection
            powerQueryConnectionOnly = 'passed'
            vbaStandardModuleSource = 'passed'
            directEngineAccepted = $true
            publicCommandAcceptance = $false
        })
    }
    Close-TestSessionWithoutSaving
    $cleanupAttempted = $true
    $cleanupSucceeded = $true
}
catch {
    $failureMessage = $_.Exception.Message
    if ($failureMessage -match 'TIMEOUT_UNCERTAIN|RECOVERY_REQUIRED') {
        $uncertain = $true
        $manualReconciliationRequired = $true
        $cleanupAttempted = $true
        try {
            Close-TestSessionWithoutSaving
            $cleanupSucceeded = $true
        }
        catch {
            $failureMessage += " Exact workbook close-without-save failed: $($_.Exception.Message)"
        }
        $sessionInvalidated = Invalidate-RunnerSession
        if (-not $sessionInvalidated) {
            $failureMessage += ' The private runner session could not be invalidated automatically.'
        }
    }
    else {
        try {
            Invoke-KnownStateCleanup 'cli'
            Close-TestSessionWithoutSaving
            $cleanupSucceeded = $true
        }
        catch {
            $failureMessage += " Cleanup failed: $($_.Exception.Message)"
            $manualReconciliationRequired = $true
            if ($_.Exception.Message -match 'TIMEOUT_UNCERTAIN') {
                $uncertain = $true
            }
            $sessionInvalidated = Invalidate-RunnerSession
            if (-not $sessionInvalidated) {
                $failureMessage += ' The private runner session could not be invalidated automatically.'
            }
        }
    }
}

$receipt = [ordered]@{
    schemaVersion = 1
    status = if ($null -eq $failureMessage) { 'passed' } else { 'failed' }
    acceptanceScope = 'direct-helper-engine'
    runtimeProof = ($null -eq $failureMessage)
    publicCommandAcceptance = $false
    helperPath = $helperPath
    workbookPath = $workbookPath
    helperSourcePath = [IO.Path]::GetFullPath($helperSourcePath)
    helperSourceSha256 = $helperSourceSha256
    entryPoints = $results
    cleanup = [ordered]@{
        attempted = $cleanupAttempted
        succeeded = $cleanupSucceeded
        saved = $false
    }
    sessionInvalidated = $sessionInvalidated
    uncertain = $uncertain
    retainedWorkbookPath = if ($manualReconciliationRequired) { $workbookPath } else { $null }
    manualReconciliationRequired = $manualReconciliationRequired
    publicProofFlagsChanged = $false
    error = $failureMessage
}
$receipt | ConvertTo-Json -Depth 20
if ($null -ne $failureMessage) { exit 1 }
exit 0
