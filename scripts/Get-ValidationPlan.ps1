#Requires -Version 7.4

function Get-ValidationPlan {
    [CmdletBinding()]
    param(
        [AllowEmptyCollection()][string[]]$Paths = @(),
        [switch]$Full,
        [switch]$Staged,
        [string]$BaseRef,
        [string]$HeadRef = 'HEAD',
        [ValidateSet('MergeBase', 'Direct')][string]$Comparison = 'MergeBase'
    )
    if ($Staged) {
        $mergeHead = & git rev-parse --verify --quiet MERGE_HEAD 2>$null
        $parent = if ($LASTEXITCODE -eq 0 -and $mergeHead) { $mergeHead } else { 'HEAD' }
        $Paths = @(& git -c core.quotepath=false diff --cached --name-only --no-renames $parent)
        if ($LASTEXITCODE -ne 0) { throw 'Cannot determine staged validation inputs.' }
    } elseif ($BaseRef) {
        $range = if ($Comparison -eq 'MergeBase') { "$BaseRef...$HeadRef" } else { "$BaseRef..$HeadRef" }
        $Paths = @(& git -c core.quotepath=false diff --name-only --no-renames $range)
        if ($LASTEXITCODE -ne 0) { throw 'Cannot determine changed validation inputs.' }
    }
    $runtime = @('CLI', 'ComInterop', 'Core', 'McpServer', 'Service')
    $tooling = @('Packaging', 'ScriptSafety', 'SkillGeneration')
    $allOwners = @($runtime + $tooling + 'Diagnostics')
    $plan = [ordered]@{
        SchemaVersion = 2; Build = $false; Excel = $false; FullE2E = $false
        PackageComponents = @(); NpmTests = $false
        DocumentationCounts = $false; FullSolutionBuild = $false
        HookTests = $false; SkillTests = $false; PackagingTests = $false
        Docs = $false; InfrastructureDiagnostics = $false
        ChangedPaths = @(); Reasons = @(); BuildProjects = @(); CodeQlLanguages = @(); ExcelGroups = @()
        FastFilters = @{}; ProcessFilters = @{}; ToolingFilters = @{}; ToolingOnDemandFilters = @{}
        ExcelSelections = @()
    }
    $aliases = @{
        Range = 'Range,Ranges,FineFormatting,Protection,NativeDataCleanup'
        Sheet = 'Sheet,Worksheets,StructuredFilters,PageLayout,Protection'
        Table = 'Table,Tables,TableStyles'; Filtering = 'StructuredFilters,Range,Tables,Worksheets'
        NamedRange = 'Parameters,NamedRange'; Chart = 'Chart,Charts,ChartDepth'
        PivotTable = 'PivotTable,PivotTables,PivotDepth,PivotCalculation'; Slicer = 'Slicer,Timelines'
        Drawing = 'Drawing,DrawingLayout'; ConditionalFormat = 'ConditionalFormat,ConditionalRuleEditing'
        Calculation = 'CalculationMode'; Workbook = 'Workbook,WorkbookTheme,CellStyles,PageLayout'
        File = 'Files,Session,SessionLifecycle'; Vba = 'VBA,VBATrust'
        PowerQueryTrace = 'PowerQuery'; Settings = 'PowerQuery'
    }
    $features = @{
        CoreFree = ('Analysis,CalculationMode,Chart,ChartDepth,ConditionalFormat,Connection,DataModel,' +
            'DocumentInspector,Drawing,ErrorHandling,ExternalOlapSchema,GeneratedContracts,Parameters,' +
            'ParameterTransforms,PivotCalculation,PivotTables,PowerQuery,PythonInExcel,QueryTable,Range,' +
            'Results,Screenshot,ServiceRegistry,Slicer,Tables,VBA,Workbook') -split ','
        CoreExcel = 'ChartDepth,Connection,DataModel,PivotTables,PowerQuery,PythonInExcel,Slicer,VBA' -split ','
        ServiceFree = 'ExcelMcpService,PowerQuery,ServiceClient,ServiceDaemon,SessionLifecycle,StreamJsonRpc' -split ','
        ServiceExcel = ('Analysis,CalculationMode,CellStyles,ChartDepth,Charts,ConditionalFormat,' +
            'ConditionalRuleEditing,Connection,DataModel,DocumentInspector,Drawing,DrawingLayout,Files,' +
            'FineFormatting,NativeDataCleanup,PageLayout,Parameters,PivotCalculation,PivotDepth,PivotTables,' +
            'PowerQuery,Protection,PythonInExcel,QueryTable,Range,Ranges,Screenshot,Session,SessionLifecycle,' +
            'Sheet,Slicer,StructuredFilters,Table,Tables,TableStyles,Timelines,VBA,Window,Workbook,WorkbookTheme,XmlMap') -split ','
    }
    function Add-Filter([string]$Kind, [string]$Owner, [string]$Filter) {
        if ($Owner -notin $allOwners) { throw "No validation mapping for test owner $Owner." }
        $map = $plan["${Kind}Filters"]
        $map[$Owner] = (@($map[$Owner], $Filter) | Where-Object { $_ } |
            ForEach-Object { $_ -split '\|' } | Sort-Object -Unique) -join '|'
        $plan.BuildProjects += "tests/ExcelMcp.$Owner.Tests/ExcelMcp.$Owner.Tests.csproj"
    }
    function Add-Excel([string]$Owner, [string]$Filter, [string]$Area) {
        if (-not @($plan.ExcelSelections | Where-Object {
            $_.Project -eq $Owner -and $_.Filter -eq $Filter -and $_.Area -eq $Area
        }).Count) {
            $plan.ExcelSelections += @{ Project = $Owner; Filter = $Filter; Area = $Area }
        }
        $plan.ExcelGroups += $Area
        $plan.BuildProjects += "tests/ExcelMcp.$Owner.Tests/ExcelMcp.$Owner.Tests.csproj"
    }
    function Add-Owner([string]$Owner) {
        $kind = if ($Owner -in $tooling) { 'Tooling' } else { 'Fast' }
        Add-Filter $kind $Owner 'FullyQualifiedName~Sbroenne.ExcelMcp.'
        if ($Owner -eq 'CLI') { Add-Filter Process CLI 'FullyQualifiedName~Sbroenne.ExcelMcp.' }
        if ($Owner -in $runtime) { Add-Excel $Owner 'FullyQualifiedName~Sbroenne.ExcelMcp.' All }
    }
    function Add-Features([string]$Owner, [string[]]$Labels, [string]$Area) {
        $free = @($Labels | Where-Object { $_ -in $features["${Owner}Free"] })
        $excel = @($Labels | Where-Object { $_ -in $features["${Owner}Excel"] })
        if ($free.Count) { Add-Filter Fast $Owner (($free | ForEach-Object { "Feature=$_" }) -join '|') }
        if ($excel.Count) { Add-Excel $Owner (($excel | ForEach-Object { "Feature=$_" }) -join '|') $Area }
    }
    function Add-Contracts {
        foreach ($owner in @('Core', 'CLI', 'McpServer')) { Add-Filter Fast $owner 'Feature=GeneratedContracts' }
        $plan.DocumentationCounts = $true
        $plan.FullSolutionBuild = $true
    }
    function Add-Shared([string[]]$Owners) {
        foreach ($owner in $Owners) { Add-Owner $owner }
        $plan.FullE2E = $true
        $plan.InfrastructureDiagnostics = $true
        $plan.FullSolutionBuild = $true
    }
    if ($Full) {
        Add-Shared ($runtime + $tooling)
        $plan.DocumentationCounts = $true
        $plan.Docs = $true
        $plan.CodeQlLanguages = @('actions', 'csharp', 'javascript-typescript', 'python')
        $plan.Reasons += 'Complete validation explicitly selected.'
        $plan.PackageComponents = @('Cli', 'Extension', 'Mcp', 'Mcpb', 'Plugins')
        $plan.NpmTests = $true
    }
    foreach ($inputPath in $(if ($Full) { @() } else { $Paths | Sort-Object -Unique })) {
        if ([string]::IsNullOrWhiteSpace($inputPath)) { throw 'Validation paths cannot be empty.' }
        $path = $inputPath.Replace('\', '/')
        if ($path -match '^/|^[A-Za-z]:|(^|/)\.\.?(/|$)') { throw "Validation inputs must be repository-relative paths: $path." }
        $plan.ChangedPaths += $path
        $plan.CodeQlLanguages += Get-ValidationLanguages $path
        switch -Regex -CaseSensitive ($path) {
            '^(Directory\.|global\.json$|NuGet\.Config$|Sbroenne\.ExcelMcp\.sln$|package(-lock)?\.json$|\.npmrc$|scripts/(PackageHelpers|Build-ReleasePackages|Get-ValidationPlan|Get-CiValidationPlan|ValidationExecution)\.ps1$|\.github/workflows/(ci|release)\.yml$|README\.md$|LICENSE$|CHANGELOG\.md$)' {
                $plan.PackageComponents += @('Cli', 'Extension', 'Mcp', 'Mcpb', 'Plugins')
                $plan.NpmTests = $true; break
            }
            '^src/ExcelMcp\.CLI/(.*\.(csproj|props|targets)$|README\.md$)' {
                $plan.PackageComponents += 'Cli'; break
            }
            '^src/ExcelMcp\.McpServer/(.*\.(csproj|props|targets)$|README\.md$)' {
                $plan.PackageComponents += @('Extension', 'Mcp'); break
            }
            '^src/.*\.(csproj|props|targets)$' {
                $plan.PackageComponents += @('Cli', 'Extension', 'Mcp'); break
            }
            '^vscode-extension/|^scripts/ExtensionPackaging\.ps1$' {
                $plan.PackageComponents += 'Extension'; break
            }
            '^skills/excel-cli(?:-report-formatting)?/' {
                $plan.PackageComponents += 'Plugins'; break
            }
            '^skills/excel-mcp-report-formatting/|^scripts/Build-AgentSkills\.ps1$|^docs/reference/report-formatting\.md$' {
                $plan.PackageComponents += @('Extension', 'Plugins'); break
            }
            '^\.github/plugins/|^scripts/Build-Plugins\.ps1$' {
                $plan.PackageComponents += 'Plugins'; break
            }
            '^mcpb/' { $plan.PackageComponents += 'Mcpb'; break }
            '^npm-packages/shared/test/' { $plan.NpmTests = $true; break }
            '^npm-packages/shared/|^scripts/(Build|Test)-NpmPackages\.ps1$' {
                $plan.PackageComponents += @('Cli', 'Mcp'); $plan.NpmTests = $true; break
            }
            '^npm-packages/excelcli(?:-win32-(?:x64|arm64))?/' {
                $plan.PackageComponents += 'Cli'; $plan.NpmTests = $true; break
            }
            '^npm-packages/mcp-server-excel(?:-win32-(?:x64|arm64))?/' {
                $plan.PackageComponents += 'Mcp'; $plan.NpmTests = $true; break
            }
        }
        if ($path -match '^(docs/(?!agents/)|gh-pages/|samples/world-bank-dashboard/|README\.md$|FEATURES\.md$|CHANGELOG\.md$|LICENSE$|doc-counts\.json$)') { $plan.Docs = $true }
        switch -Regex -CaseSensitive ($path) {
            '^(skills/|scripts/Build-AgentSkills\.ps1$|docs/(AGENT-SKILLS\.md|reference/report-formatting\.md)$)' {
                Add-Filter Tooling SkillGeneration 'Feature=SkillGeneration'; break
            }
            '^(\.github/plugins/|scripts/(Build-Plugins|Sync-PublishedPluginRepo|Publish-PreparedPlugins)\.ps1$|\.github/workflows/publish-plugins\.yml$)' {
                Add-Filter Tooling Packaging 'Feature=PluginPublication|Feature=PluginBootstrap|Feature=PluginSkillVersion'; break
            }
            '^mcpb/' { Add-Filter Tooling Packaging 'Feature=McpbPackaging'; break }
            '^scripts/(PackageHelpers|ExtensionPackaging)\.ps1$' {
                Add-Owner Packaging; Add-Owner SkillGeneration; Add-Filter Tooling ScriptSafety 'Feature=PreCommit'; break
            }
            '^(doc-counts\.json$|scripts/check-doc-counts\.ps1$)' {
                Add-Filter Tooling ScriptSafety 'FullyQualifiedName~DocumentationAnalyzerTests'
                $plan.DocumentationCounts = $true; $plan.FullSolutionBuild = $true; break
            }
            '^scripts/Publish-GitHubRelease\.ps1$' {
                Add-Filter Tooling Packaging 'FullyQualifiedName~Sbroenne.ExcelMcp.Packaging.Tests.ReleaseMetadataScriptTests.GitHubRelease_'; break
            }
            '^(scripts/(Build-NpmPackages|Test-NpmPackages|Build-ReleasePackages|Build-Changelog|Format-ReleaseNotes|Update-(ReleaseVersion|McpRegistry)Metadata|Resolve-McpRegistryRelease|Test-McpRegistryPublication)\.ps1$|\.github/workflows/(release|publish-mcp-registry)\.yml$|package(-lock)?\.json$|\.npmrc$)' {
                Add-Owner Packaging; break
            }
            '^src/ExcelMcp\.Core/Commands/([^/]+)/' {
                $area = $Matches[1]
                $labels = if ($aliases.ContainsKey($area)) { $aliases[$area] -split ',' } else { @($area) }
                if (-not @($labels | Where-Object { $_ -in $features.CoreFree -or $_ -in $features.ServiceExcel }).Count -and $area -ne 'Diag') {
                    throw "No validation mapping for command area $area ($path)."
                }
                Add-Features Core $labels $area; Add-Features Service $labels $area
                if ($area -eq 'Diag') { Add-Filter Fast CLI 'Feature=Diag' }
                if ($area -eq 'DataModel') { Add-Features Service @('PowerQuery', 'PivotTables', 'Slicer', 'Tables') 'DataModel dependencies' }
                if ($area -in @('Connection', 'QueryTable')) { Add-Features Service @('PowerQuery') 'PowerQuery refresh dependency' }
                if ([IO.Path]::GetFileName($path).StartsWith('I') -or $path.Contains('/Attributes/')) {
                    Add-Contracts
                    foreach ($adapter in @('CLI', 'McpServer')) {
                        Add-Filter Fast $adapter (($labels | ForEach-Object { "Feature=$_" }) -join '|')
                    }
                }
                $plan.BuildProjects += 'src/ExcelMcp.Core/ExcelMcp.Core.csproj'
                break
            }
            '^tests/ExcelMcp\.([^/]+)\.Tests/(.+)' {
                $owner = $Matches[1]
                $file = $Matches[2]
                if ($owner -notin $allOwners) { throw "No validation mapping for $path." }
                $excelTransport = $owner -eq 'McpServer' -and
                    $file -match '^Integration/Tools/(ExcelFileToolProtocolRegressionTests|McpProcessOwnershipTests|McpServerSmokeTests|McpSessionLifetimeTests|SharePointWorkbookTests|McpIntegrationTestBase|McpOwnedExcelProcesses|McpProgramTransportFixture|McpResponseAssertions|ProgramTransportTestHost)(?:\.[^/]+)?\.cs$'
                if ($owner -eq 'Service' -and
                    $file -match '^PersistentServiceWindowTests(?:\.[A-Za-z0-9]+)*\.cs$') {
                    $class = 'PersistentServiceWindowTests'
                    $filter = "FullyQualifiedName~Sbroenne.ExcelMcp.Service.Tests.$class"
                    Add-Excel Service $filter $class
                } elseif ($file -match '^(Unit/|Integration/Tools/)' -and -not $excelTransport) {
                    Add-Filter Fast $owner 'FullyQualifiedName~Sbroenne.ExcelMcp.'
                    if ($owner -eq 'CLI') { Add-Filter Process CLI 'FullyQualifiedName~Sbroenne.ExcelMcp.' }
                } else {
                    # Shared helpers and other partial classes remain covered by the whole owning project.
                    Add-Owner $owner
                }
                break
            }
            '^tests/Shared/(GeneratedAssetsFixture|PackagingScriptTestHelper)\.cs$' {
                Add-Owner Packaging; Add-Owner SkillGeneration; break
            }
            '^tests/Shared/(TelemetryUserIdentityTests|UsageTelemetry)' {
                Add-Filter Fast CLI 'Feature=Telemetry'; Add-Filter Fast McpServer 'Feature=Telemetry'; break
            }
            '^tests/Shared/' { Add-Shared $allOwners; break }
            '^(tools/ExcelMcp\.Build/|build\.ps1$|tests/strict\.runsettings$|scripts/(Get-ValidationPlan|Get-CiValidationPlan|Build-CiInputs|Invoke-BuildTool|Invoke-NativeCommand|ValidationExecution|Invoke-ExcelFreeTests|Invoke-ExcelTests|Get-ExcelTestGroups|Invoke-TestStage|Test-CiCompletion|pre-commit)\.ps1$|\.github/workflows/ci\.yml$)' {
                Add-Filter Tooling ScriptSafety 'Feature=PreCommit'; break
            }
            '^\.editorconfig$' {
                foreach ($owner in $allOwners) { Add-Filter $(if ($owner -in $tooling) { 'Tooling' } else { 'Fast' }) $owner 'FullyQualifiedName~TestClassificationArchitectureTests' }
                $plan.FullSolutionBuild = $true; break
            }
            '^(Directory\.|global\.json$|NuGet\.Config$|Sbroenne\.ExcelMcp\.sln$|tests/Directory\.)' {
                Add-Shared $allOwners; break
            }
            '^src/ExcelMcp\.(Core|ComInterop|Service|Cleanup)/' { Add-Shared $runtime; break }
            '^(src/ExcelMcp\.Analyzers/|BannedSymbols[^/]*\.txt$|scripts/check-(com-leaks|success-flag|dynamic-casts|workbook-package-access)\.ps1$)' {
                Add-Filter Tooling ScriptSafety 'FullyQualifiedName~AnalyzerTests|Feature=PreCommit'
                $plan.FullSolutionBuild = $true; break
            }
            '^src/ExcelMcp\.Generators' { Add-Contracts; break }
            '^src/ExcelMcp\.Diagnostics/' { Add-Owner Diagnostics; break }
            '^src/Shared/(TelemetryUserIdentity|TelemetrySettings|UsageTelemetry|UsageRecord|LocalTelemetry)\.cs$' {
                Add-Filter Fast CLI 'Feature=Telemetry|FullyQualifiedName~EntryPointTelemetryTests|FullyQualifiedName~TelemetryUserIdentityTests'
                Add-Filter Process CLI 'Feature=Telemetry'
                Add-Filter Fast McpServer 'Feature=Telemetry|FullyQualifiedName~TelemetryUserIdentityTests'; break
            }
            '^src/ExcelMcp\.(CLI|McpServer)/Telemetry/' {
                $owner = $Matches[1]; Add-Filter Fast $owner 'Feature=Telemetry'
                if ($owner -eq 'CLI') { Add-Filter Process CLI 'Feature=Telemetry' }; break
            }
            '^src/ExcelMcp\.CLI/Commands/BatchCommand\.cs$' { Add-Filter Fast CLI 'Feature=Batch'; break }
            '^src/ExcelMcp\.CLI/Infrastructure/CliServiceTray\.cs$' {
                Add-Filter Process CLI 'FullyQualifiedName~CliDaemonLifecycleTests|FullyQualifiedName~CliDaemonTests'; break
            }
            '^src/ExcelMcp\.(CLI|McpServer)/.*(?<!\.md)$' { Add-Owner $Matches[1]; break }
            '^scripts/IssueTriage\.mjs$|^\.github/workflows/issue-triage\.(md|lock\.yml)$' {
                Add-Filter Tooling ScriptSafety 'FullyQualifiedName~IssueTriageScriptTests'; break
            }
            '^scripts/[^/]+\.mjs$' {
                Add-Filter Tooling Packaging 'Feature=PluginPublication'; break
            }
            '^scripts/Install-CopilotPonytailReview\.ps1$|^\.github/workflows/copilot-setup-steps\.yml$' { break }
            '^scripts/(AzureRunnerHost|ExcelRunner(Host|Policy)|Invoke-(ExcelRunner(Control|Maintenance)|CopilotSetupNpm)|(Deploy|Register)-ExcelAgentRunner|Initialize-ExcelAgentDesktop|Install-ExcelAgent(Office|Toolchain)|Open-ExcelAgentActivation)\.ps1$' {
                if (Test-Path -LiteralPath (Join-Path $PSScriptRoot "..\$path")) {
                    throw "Retired runner file $path must not be restored without an owning validation mapping."
                }
                break
            }
            '^infrastructure/azure/(configure-analytics-oidc|deploy-appinsights)\.ps1$' {
                Add-Filter Tooling ScriptSafety 'Feature=AutomationSafety'
                $plan.ToolingOnDemandFilters.ScriptSafety = 'FullyQualifiedName=Sbroenne.ExcelMcp.ScriptSafety.Tests.AutomationSafetyTests.AzurePreview_NeverMutatesAndRejectsFailedReads|FullyQualifiedName=Sbroenne.ExcelMcp.ScriptSafety.Tests.AutomationSafetyTests.AnalyticsOidc_SwitchesToSelectedTenantAndRestoresOriginalSubscription'
                break
            }
            '^(scripts/check-|scripts/tests/|scripts/Start-OlapTestCube\.ps1$|videos/excel-mcp-intro/Capture-Evidence\.ps1$)' {
                Add-Filter Tooling ScriptSafety 'Feature=AutomationSafety'; break
            }
            '^scripts/(Test-E2E|Test-CliWorkflow|Stop-ExcelCliService)\.ps1$' {
                $plan.FullE2E = $true
                Add-Excel CLI 'FullyQualifiedName~CliWorkflowAcceptanceTests' Acceptance
                Add-Excel McpServer 'FullyQualifiedName~McpServerSmokeTests' Acceptance
                break
            }
            '^(README\.md|CHANGELOG\.md|LICENSE)$|^src/ExcelMcp\.(CLI|McpServer)/README\.md$|^vscode-extension/(README\.md|CHANGELOG\.md|LICENSE)$|^npm-packages/(excelcli|mcp-server-excel)[^/]*/README\.md$' {
                Add-Filter Tooling Packaging 'Feature=Packaging'; break
            }
            '\.md$|^(docs/|gh-pages/|\.github/|\.changeset/|\.squad/|\.vscode/|infrastructure/|videos/|llm-tests/|npm-packages/|vscode-extension/|samples/world-bank-dashboard/|scripts/(AnalyticsStorage\.ps1$|.*(UsageAnalytics|StarHistory))|\.(gitignore|gitattributes|mcp\.json)$)' { break }
            default { throw "No validation mapping for $path. Add an owning area; a full-suite fallback is not allowed." }
        }
        $plan.Reasons += "$path -> explicit owning validation groups"
    }
    $plan.Excel = $plan.ExcelSelections.Count -gt 0 -or $plan.FullE2E
    $plan.HookTests = $plan.ToolingFilters.ContainsKey('ScriptSafety')
    $plan.SkillTests = $plan.ToolingFilters.ContainsKey('SkillGeneration')
    $plan.PackagingTests = $plan.ToolingFilters.ContainsKey('Packaging')
    foreach ($name in @('ChangedPaths', 'BuildProjects', 'CodeQlLanguages', 'ExcelGroups', 'PackageComponents')) {
        $plan[$name] = @($plan[$name] | Sort-Object -Unique)
    }
    $plan.Build = $plan.BuildProjects.Count -gt 0 -or $plan.FullSolutionBuild
    $plan.FastProjects = @($plan.FastFilters.Keys | Sort-Object)
    $plan.ProcessProjects = @($plan.ProcessFilters.Keys | Sort-Object)
    $plan.ToolingProjects = @($plan.ToolingFilters.Keys | Sort-Object)
    $plan.CiTestGroups = @('Fast', 'Process', 'Tooling' | Where-Object { $plan["${_}Filters"].Count })
    $plan.FullSolutionBuildGroup = if ($plan.FullSolutionBuild -and $plan.CiTestGroups.Count) { $plan.CiTestGroups[0] } else { '' }
    $plan.DocumentationChecksGroup = if ($plan.DocumentationCounts) { $plan.FullSolutionBuildGroup } else { '' }
    $coveredTestProjects = @($plan.CiTestGroups | ForEach-Object {
        $plan["${_}Projects"] | ForEach-Object { "tests/ExcelMcp.$_.Tests/ExcelMcp.$_.Tests.csproj" }
    })
    $uncoveredTestProjects = @($plan.BuildProjects | Where-Object {
        $_ -match '^tests/ExcelMcp\.[^/]+\.Tests/ExcelMcp\.[^/]+\.Tests\.csproj$' -and
        $_ -notin $coveredTestProjects
    })
    $plan.BuildOnly = $plan.Build -and (
        -not $plan.CiTestGroups.Count -or
        (-not $plan.FullSolutionBuild -and $uncoveredTestProjects.Count -gt 0)
    )
    return $plan
}

function Get-ValidationLanguages([string]$Path) {
    $languages = @()
    if ($Path -match '\.(cs|csproj|sln|slnf|props|targets)$|(^|/)(global\.json|NuGet\.Config)$') { $languages += 'csharp' }
    if ($Path -match '\.(js|jsx|mjs|cjs|ts|tsx|mts|cts)$|(^|/)(package(-lock)?\.json|[jt]sconfig[^/]*\.json|\.npmrc)$') { $languages += 'javascript-typescript' }
    if ($Path -match '\.py$|(^|/)(requirements[^/]*\.txt|pyproject\.toml|poetry\.lock|uv\.lock|Pipfile(\.lock)?|setup\.cfg)$') { $languages += 'python' }
    if ($Path -match '^\.github/(workflows/[^/]+\.ya?ml$|actions/.+/action\.ya?ml$)') { $languages += 'actions' }
    if ($Path -match '^\.github/(workflows/codeql\.yml|codeql/)|^(tools/ExcelMcp\.Build/|scripts/Get-(Ci)?ValidationPlan\.ps1$)') {
        $languages += @('actions', 'csharp', 'javascript-typescript', 'python')
    }
    $languages | Sort-Object -Unique
}
