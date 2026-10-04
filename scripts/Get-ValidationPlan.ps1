function Get-ValidationPlan {
    [CmdletBinding()]
    param(
        [AllowEmptyCollection()][string[]]$Paths = @(),
        [switch]$Full
    )

    $plan = [ordered]@{
        Build = $false
        Excel = $false
        SourceChecks = $false
        HookTests = $false
        Cli = $false
        Mcp = $false
        Extension = $false
        Mcpb = $false
        Skills = $false
        SkillTests = $false
        PackagingTests = $false
        Plugins = $false
        Reasons = [Collections.Generic.List[string]]::new()
    }
    $fast = [Collections.Generic.HashSet[string]]::new([StringComparer]::Ordinal)
    $process = [Collections.Generic.HashSet[string]]::new([StringComparer]::Ordinal)
    $tooling = [Collections.Generic.HashSet[string]]::new([StringComparer]::Ordinal)
    $toolingOwners = [Collections.Generic.HashSet[string]]::new([StringComparer]::Ordinal)
    $excel = [Collections.Generic.HashSet[string]]::new([StringComparer]::Ordinal)
    $allProjects = @('CLI', 'ComInterop', 'Core', 'McpServer', 'Service')
    $allExcel = @('Acceptance', 'Data', 'Desktop', 'Editing', 'Infrastructure', 'Lifecycle', 'Reporting', 'VBA')
    $fullTooling = $false
    $npmTests = $false
    $lockfileTests = $false
    $documentationCounts = $false
    $infrastructureDiagnostics = $false
    foreach ($original in $Paths) {
        $path = $original.Replace('\', '/')
        $kind = switch -Regex ($path) {
            '^(Directory\.Build\..*|Directory\.Packages\.props|global\.json|NuGet\.Config|Sbroenne\.ExcelMcp\.sln)$' { 'runtime'; break }
            '^src/ExcelMcp\.(Core|ComInterop|Service|Cleanup|Generators[^/]*)/' { 'runtime'; break }
            '^src/ExcelMcp\.CLI/' { 'cli'; break }
            '^src/ExcelMcp\.McpServer/' { 'mcp'; break }
            '^src/ExcelMcp\.Build\.Tasks/|^skills/|^docs/reference/report-formatting\.md$' { 'skills'; break }
            '^src/ExcelMcp\.Diagnostics/|^\.editorconfig$' { 'build'; break }
            '^scripts/(Test-E2E|Test-CliWorkflow|Test-CliApiCoverage|Stop-ExcelMcpProcesses)\.ps1$|^tests/.*/(PreBuildGracefulSaveAcceptanceTests|McpServerSmokeTests|CliWorkflowAcceptanceTests)\.cs$' { 'runtime'; break }
            '^tests/' { 'tests'; break }
            '^vscode-extension/' { 'extension'; break }
            '^mcpb/' { 'mcpb'; break }
            '^npm-packages/excelcli' { 'cli-package'; break }
            '^npm-packages/mcp-server-excel' { 'mcp-package'; break }
            '^npm-packages/shared/' { 'npm-packages'; break }
            '^\.github/plugins/|^\.github/workflows/(publish-plugins\.yml|update-awesome-copilot\.(md|lock\.yml))$' { 'plugins'; break }
            '^scripts/Build-AgentSkills\.ps1$' { 'skills'; break }
            '^scripts/(Build-Plugins|Sync-PublishedPluginRepo|Publish-PreparedPlugins)\.ps1$|^scripts/(PluginContent|AwesomeCopilotPolicy|Update-AwesomeCopilot)\.mjs$' { 'plugins'; break }
            '^scripts/(Build-NpmPackages|Test-NpmPackages|Build-ReleasePackages|PackageHelpers)\.ps1$|^\.github/workflows/(release|publish-mcp-registry)\.yml$' { 'packages'; break }
            '^doc-counts\.json$|^scripts/check-doc-counts\.ps1$' { 'doc-counts'; break }
            '^scripts/(pre-commit|Get-ValidationPlan|Get-CiValidationPlan|Get-ExcelTestGroups|Invoke-ExcelFreeTests|Invoke-ExcelTests|Invoke-TestStage|Test-CiCompletion|check-|Test-NpmLockfiles)' { 'tests'; break }
            '^\.github/workflows/ci\.yml$' { 'pipeline'; break }
            '^scripts/(Build-Changelog|Update-(ReleaseVersion|McpRegistry)Metadata|Resolve-McpRegistryRelease|Test-McpRegistryPublication)\.ps1$' { 'packages'; break }
            '^scripts/(Update|Restore|Persist|Test)-StarHistory\.ps1$|^scripts/.*UsageAnalytics.*\.ps1$' { 'maintenance'; break }
            '^scripts/Invoke-CopilotSetupNpm\.ps1$|^\.github/workflows/copilot-setup-steps\.yml$' { 'safety'; break }
            '^scripts/(AzureRunnerHost|ExcelRunner(Host|Policy)|Invoke-ExcelRunner(Control|Maintenance))\.ps1$|^scripts/(Deploy|Initialize|Install|Open|Register)-ExcelAgent(Runner|Desktop|Office|Toolchain|Activation)\.ps1$|^scripts/tests/[^/]+\.tests\.ps1$|^infrastructure/azure/[^/]+\.ps1$|^infrastructure/azure/excel[^/]*\.bicep$|^\.github/workflows/excel-runner[^/]*\.yml$' { 'safety'; break }
            '^infrastructure/azure/(configure-analytics-oidc|deploy-appinsights)\.ps1$|^videos/excel-mcp-intro/Capture-Evidence\.ps1$' { 'safety'; break }
            '^docs/|^gh-pages/|^videos/|^infrastructure/|^\.changeset/|^\.github/|\.md$|^\.(gitignore|gitattributes)$' { 'documentation'; break }
            '^(package(-lock)?\.json|\.npmrc)$' { 'packages'; break }
            default { 'unknown' }
        }
        $plan.Reasons.Add("$path -> $kind")
        if ($kind -in @('runtime', 'cli', 'mcp', 'unknown')) {
            $plan.Build = $true
            $plan.Excel = $true
            $plan.SourceChecks = $true
        }
        if ($kind -in @('runtime', 'cli', 'cli-package', 'npm-packages', 'packages', 'pipeline', 'unknown')) { $plan.Cli = $true }
        if ($kind -in @('runtime', 'mcp', 'mcp-package', 'npm-packages', 'packages', 'pipeline', 'unknown')) { $plan.Mcp = $true }
        if ($kind -in @('runtime', 'cli', 'mcp', 'skills', 'packages', 'pipeline', 'unknown')) { $plan.Skills = $true }
        if ($kind -in @('extension', 'skills', 'packages', 'pipeline', 'runtime', 'mcp', 'unknown')) { $plan.Extension = $true }
        if ($kind -in @('mcpb', 'packages', 'pipeline', 'runtime', 'mcp', 'unknown')) { $plan.Mcpb = $true }
        if ($kind -in @('plugins', 'skills', 'packages', 'pipeline', 'runtime', 'cli', 'mcp', 'unknown')) { $plan.Plugins = $true }
        if ($kind -in @('build', 'tests', 'skills', 'plugins', 'packages', 'safety', 'pipeline', 'doc-counts')) { $plan.Build = $true }
        if ($kind -in @('safety', 'pipeline') -or
            ($kind -eq 'tests' -and $path -notmatch '^tests/(ExcelMcp\.(SkillGeneration|Packaging)\.Tests/|Shared/(GeneratedAssetsFixture|PackagingScriptTestHelper)\.cs$)')) {
            $plan.HookTests = $true
        }
        if ($kind -in @('skills', 'pipeline') -or
            $path -match '^tests/(ExcelMcp\.SkillGeneration\.Tests/|Shared/(GeneratedAssetsFixture|PackagingScriptTestHelper)\.cs$)') {
            $plan.SkillTests = $true
        }
        if ($kind -in @('plugins', 'packages', 'pipeline', 'doc-counts') -or
            $path -match '^tests/(ExcelMcp\.Packaging\.Tests/|Shared/(GeneratedAssetsFixture|PackagingScriptTestHelper)\.cs$)' -or
            $path -match '^mcpb/(Build-McpBundle|McpbPackaging)\.ps1$') {
            $plan.PackagingTests = $true
            $plan.Build = $true
        }

        if ($kind -in @('runtime', 'cli', 'mcp', 'unknown', 'pipeline', 'build')) {
            foreach ($project in $allProjects) { [void]$fast.Add($project) }
            [void]$process.Add('CLI')
            $fullTooling = $true
        }
        if ($kind -in @('runtime', 'cli', 'mcp', 'unknown')) {
            [void]$excel.Add('Acceptance')
            $feature = switch -Regex ($path) {
                '^src/ExcelMcp\.Core/Commands/(Range|Sheet|NamedRange|Table|XmlMap)/' { 'Editing'; break }
                '^src/ExcelMcp\.Core/Commands/(Chart|PivotTable|Slicer|Drawing|ConditionalFormat|Analysis|Calculation)/' { 'Reporting'; break }
                '^src/ExcelMcp\.Core/Commands/(Connection|QueryTable|PowerQuery|DataModel|PythonInExcel)/' { 'Data'; break }
                '^src/ExcelMcp\.Core/Commands/Vba/' { 'VBA'; break }
                '^src/ExcelMcp\.Core/Commands/(Window|Screenshot)/' { 'Desktop'; break }
                '^src/ExcelMcp\.Core/Commands/(File|Workbook)/|^src/ExcelMcp\.(CLI|McpServer)/' { 'Lifecycle'; break }
                default { 'All' }
            }
            if ($feature -eq 'All') {
                foreach ($group in $allExcel) { [void]$excel.Add($group) }
            } else { [void]$excel.Add($feature) }
            if ($path -match '^src/ExcelMcp\.Core/Commands/(Table|PivotTable)/') {
                [void]$excel.Add('Data')
            }
            if ($path -match '^src/ExcelMcp\.Core/Commands/DataModel/') {
                [void]$excel.Add('Editing')
                [void]$excel.Add('Reporting')
            }
            if ($path -match '^src/ExcelMcp\.(ComInterop|Service|Cleanup)/|^(Directory\.|global\.json|NuGet\.Config|Sbroenne\.ExcelMcp\.sln)' -or $kind -eq 'unknown') {
                $infrastructureDiagnostics = $true
            }
        }
        if ($path -match '^tests/ExcelMcp\.(CLI|ComInterop|Core|McpServer|Service)\.Tests/') {
            $project = $Matches[1]
            [void]$fast.Add($project)
            if ($project -eq 'CLI') { [void]$process.Add($project) }
        }
        if (($path -match '^tests/Shared/' -and
             $path -notmatch '^tests/Shared/(GeneratedAssetsFixture|PackagingScriptTestHelper)\.cs$') -or
            $path -match '^tests/.*\.csproj$|^tests/.*xunit\.runner\.json$') {
            foreach ($project in $allProjects) { [void]$fast.Add($project) }
            [void]$process.Add('CLI')
            $fullTooling = $true
            foreach ($group in $allExcel) { [void]$excel.Add($group) }
            $infrastructureDiagnostics = $true
        }
        if ($kind -eq 'tests' -and $path -match '^scripts/') {
            [void]$toolingOwners.Add('ScriptSafety')
            foreach ($feature in @('PreCommit', 'AutomationSafety')) { [void]$tooling.Add("Feature=$feature") }
        }
        if ($path -match '^scripts/(Invoke-TestStage|Get-CiValidationPlan|Test-CiCompletion|Invoke-ExcelTests|Get-ExcelTestGroups)\.ps1$') {
            foreach ($project in $allProjects) { [void]$fast.Add($project) }
            [void]$process.Add('CLI')
            $fullTooling = $true
        }
        if ($kind -eq 'skills') {
            [void]$toolingOwners.Add('SkillGeneration')
            [void]$tooling.Add('Feature=SkillGeneration')
        }
        if ($path -match '^tests/Shared/(GeneratedAssetsFixture|PackagingScriptTestHelper)\.cs$') {
            [void]$toolingOwners.Add('SkillGeneration')
            [void]$toolingOwners.Add('Packaging')
            [void]$tooling.Add('RequiresExcel=false')
        }
        if ($kind -eq 'safety') {
            [void]$toolingOwners.Add('ScriptSafety')
            [void]$tooling.Add('Feature=AutomationSafety')
        }
        if ($kind -eq 'plugins') {
            [void]$toolingOwners.Add('Packaging')
            foreach ($feature in @('PluginPublication', 'PluginBootstrap', 'PluginSkillVersion')) {
                [void]$tooling.Add("Feature=$feature")
            }
        }
        if ($kind -eq 'doc-counts') {
            [void]$toolingOwners.Add('Packaging')
            $documentationCounts = $true
            [void]$tooling.Add('FullyQualifiedName~DocumentationCounts')
        }
        if ($kind -in @('packages', 'mcpb')) {
            [void]$toolingOwners.Add('Packaging')
            foreach ($feature in @('Packaging', 'McpbPackaging', 'PluginSkillVersion', 'ReleaseMetadata', 'PluginPublication')) {
                [void]$tooling.Add("Feature=$feature")
            }
        }
        if ($path -match '^tests/ExcelMcp\.(SkillGeneration|Packaging|ScriptSafety)\.Tests/') {
            [void]$toolingOwners.Add($Matches[1])
            if ($path -match 'PluginPublication') { [void]$tooling.Add('Feature=PluginPublication') }
            else { [void]$tooling.Add('RequiresExcel=false') }
        }
        if ($kind -in @('runtime', 'cli', 'mcp', 'unknown', 'pipeline', 'packages', 'npm-packages', 'cli-package', 'mcp-package')) {
            $npmTests = $true
        }
        if ($path -match '(^|/)(package-lock\.json|\.npmrc)$|^scripts/(Test-NpmLockfiles|check-npm-lockfiles)\.ps1$|^\.github/workflows/ci\.yml$') {
            $lockfileTests = $true
        }
    }
    if ($Full) {
        $plan.Reasons.Add('Complete validation requested.')
        foreach ($flag in @('Build', 'Excel', 'SourceChecks', 'Cli', 'Mcp', 'Extension', 'Mcpb', 'Skills', 'Plugins')) { $plan[$flag] = $true }
        foreach ($project in $allProjects) { [void]$fast.Add($project) }
        [void]$process.Add('CLI')
        foreach ($group in $allExcel) { [void]$excel.Add($group) }
        $fullTooling = $true
        $npmTests = $true
        $lockfileTests = $true
        $infrastructureDiagnostics = $true
    }
    $plan.FastProjects = @($fast | Sort-Object)
    $plan.ProcessProjects = @($process | Sort-Object)
    $plan.ToolingFilter = if ($fullTooling) { 'RequiresExcel=false' } else { @($tooling | Sort-Object) -join '|' }
    $plan.ToolingProjects = if ($fullTooling) {
        @('SkillGeneration', 'Packaging', 'ScriptSafety')
    } else { @($toolingOwners | Sort-Object) }
    if ($plan.ToolingProjects.Count -and -not $plan.ToolingFilter) { $plan.ToolingFilter = 'RequiresExcel=false' }
    $plan.CiTestGroups = @(
        if ($fast.Count) { 'Fast' }
        if ($process.Count) { 'Process' }
        if ($plan.ToolingProjects.Count) { 'Tooling' }
    )
    $plan.ExcelGroups = @($excel | Sort-Object)
    $plan.SourceChecksGroup = if ($fast.Count) { 'Fast' } elseif ($documentationCounts) { 'Tooling' } else { '' }
    $plan.InfrastructureDiagnostics = $infrastructureDiagnostics
    $plan.Packages = [bool](@('Cli', 'Mcp', 'Extension', 'Mcpb', 'Skills', 'Plugins') | Where-Object { $plan[$_] }).Count
    $plan.NpmTests = $npmTests
    $plan.LockfileTests = $lockfileTests
    [pscustomobject]$plan
}
