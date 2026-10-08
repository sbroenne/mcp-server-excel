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
        AzureInfrastructureTests = $false
        Plugins = $false
        Reasons = [Collections.Generic.List[string]]::new()
    }
    $fast = [Collections.Generic.HashSet[string]]::new([StringComparer]::Ordinal)
    $process = [Collections.Generic.HashSet[string]]::new([StringComparer]::Ordinal)
    $toolingSelections = @{}
    $codeQl = [Collections.Generic.HashSet[string]]::new([StringComparer]::Ordinal)
    $excel = [Collections.Generic.HashSet[string]]::new([StringComparer]::Ordinal)
    $allProjects = @('CLI', 'ComInterop', 'Core', 'McpServer', 'Portable', 'Service')
    $allExcel = @('Acceptance', 'Data', 'Desktop', 'Editing', 'Infrastructure', 'Lifecycle', 'Reporting', 'VBA')
    $fullTooling = $false
    $npmTests = $false
    $lockfileTests = $false
    $documentationCounts = $false
    $infrastructureDiagnostics = $false
    foreach ($original in $Paths) {
        $path = $original.Replace('\', '/')
        if ($path -match '^infrastructure/azure/(configure-analytics-oidc|deploy-appinsights)\.ps1$') {
            $plan.AzureInfrastructureTests = $true
        }
        $tooling = [Collections.Generic.HashSet[string]]::new([StringComparer]::Ordinal)
        $toolingOwners = [Collections.Generic.HashSet[string]]::new([StringComparer]::Ordinal)
        if ($path -match '\.cs$|\.(csproj|sln|slnf|props|targets)$|(^|/)(global\.json|NuGet\.Config)$') {
            [void]$codeQl.Add('csharp')
        }
        if ($path -match '\.(js|jsx|mjs|cjs|ts|tsx|mts|cts)$|(^|/)(package(-lock)?\.json|[jt]sconfig[^/]*\.json|\.npmrc)$') {
            [void]$codeQl.Add('javascript-typescript')
        }
        if ($path -match '\.py$|(^|/)(requirements[^/]*\.txt|pyproject\.toml|poetry\.lock|uv\.lock|Pipfile(\.lock)?|setup\.cfg)$') {
            [void]$codeQl.Add('python')
        }
        if ($path -match '^\.github/workflows/[^/]+\.ya?ml$|^\.github/actions/.+/action\.ya?ml$') {
            [void]$codeQl.Add('actions')
        }
        if ($path -match '^\.github/(workflows/codeql\.yml|codeql/)|^scripts/(Get-ValidationPlan|Get-CiValidationPlan)\.ps1$') {
            foreach ($language in @('actions', 'csharp', 'javascript-typescript', 'python')) { [void]$codeQl.Add($language) }
        }
        $kind = switch -Regex ($path) {
            '^skills/|^docs/reference/report-formatting\.md$' { 'skills'; break }
            '^\.github/plugins/|^\.github/workflows/(publish-plugins\.yml|update-awesome-copilot\.(md|lock\.yml))$' { 'plugins'; break }
            '(^|/)(AGENTS|CLAUDE)\.md$|^\.github/copilot-instructions\.md$' { 'documentation'; break }
            '^src/ExcelMcp\.CLI/README\.md$' { 'cli-docs'; break }
            '^src/ExcelMcp\.McpServer/README\.md$' { 'mcp-docs'; break }
            '^npm-packages/excelcli[^/]*/README\.md$' { 'cli-docs'; break }
            '^npm-packages/mcp-server-excel[^/]*/README\.md$' { 'mcp-docs'; break }
            '^README\.md$' { 'runtime-docs'; break }
            '^(LICENSE|CHANGELOG\.md)$' { 'distribution-docs'; break }
            '^docs/AGENT-SKILLS\.md$' { 'skill-docs'; break }
            '^mcpb/' { 'mcpb'; break }
            '^vscode-extension/(README\.md|LICENSE|CHANGELOG\.md)$' { 'extension-docs'; break }
            '\.md$' { 'documentation'; break }
            '^(Directory\.Build\..*|Directory\.Packages\.props|global\.json|NuGet\.Config|Sbroenne\.ExcelMcp\.sln)$' { 'runtime'; break }
            '^src/ExcelMcp\.(Core|ComInterop|Service|Cleanup|Generators[^/]*)/' { 'runtime'; break }
            '^src/ExcelMcp\.CLI/' { 'cli'; break }
            '^src/ExcelMcp\.McpServer/' { 'mcp'; break }
            '^src/ExcelMcp\.Build\.Tasks/' { 'skills'; break }
            '^src/ExcelMcp\.Diagnostics/|^\.editorconfig$' { 'build'; break }
            '^scripts/(Test-E2E|Test-CliWorkflow|Test-CliApiCoverage|Stop-ExcelMcpProcesses)\.ps1$|^tests/.*/(PreBuildGracefulSaveAcceptanceTests|McpServerSmokeTests|CliWorkflowAcceptanceTests)\.cs$' { 'runtime'; break }
            '^tests/' { 'tests'; break }
            '^llm-tests/' { 'evaluation'; break }
            '^vscode-extension/' { 'extension'; break }
            '^npm-packages/excelcli' { 'cli-package'; break }
            '^npm-packages/mcp-server-excel' { 'mcp-package'; break }
            '^npm-packages/shared/' { 'npm-packages'; break }
            '^scripts/Build-AgentSkills\.ps1$' { 'skills'; break }
            '^scripts/(Build-Plugins|Sync-PublishedPluginRepo|Publish-PreparedPlugins)\.ps1$|^scripts/(PluginContent|AwesomeCopilotPolicy|Update-AwesomeCopilot)\.mjs$' { 'plugins'; break }
            '^scripts/(Build-NpmPackages|Test-NpmPackages|Build-ReleasePackages|PackageHelpers)\.ps1$|^\.github/workflows/(release|publish-mcp-registry)\.yml$' { 'packages'; break }
            '^doc-counts\.json$|^scripts/check-doc-counts\.ps1$' { 'doc-counts'; break }
            '^scripts/(pre-commit|Get-ValidationPlan|Get-CiValidationPlan|Build-CiInputs|Get-ExcelTestGroups|Invoke-ExcelFreeTests|Invoke-ExcelTests|Invoke-TestStage|Test-CiCompletion|check-|Test-NpmLockfiles)' { 'tests'; break }
            '^\.github/workflows/ci\.yml$' { 'pipeline'; break }
            '^scripts/(Build-Changelog|Update-(ReleaseVersion|McpRegistry)Metadata|Resolve-McpRegistryRelease|Test-McpRegistryPublication)\.ps1$' { 'packages'; break }
            '^scripts/(Update|Restore|Persist|Test)-StarHistory\.ps1$|^scripts/.*UsageAnalytics.*\.ps1$' { 'maintenance'; break }
            '^scripts/(Invoke-CopilotSetupNpm|Install-CopilotPonytailReview)\.ps1$|^\.github/workflows/copilot-setup-steps\.yml$' { 'maintenance'; break }
            '^scripts/(AzureRunnerHost|ExcelRunner(Host|Policy)|Invoke-ExcelRunner(Control|Maintenance))\.ps1$|^scripts/(Deploy|Initialize|Install|Open|Register)-ExcelAgent(Runner|Desktop|Office|Toolchain|Activation)\.ps1$|^infrastructure/azure/[^/]*excel[^/]*\.(ps1|bicep)$|^\.github/workflows/excel-runner[^/]*\.yml$' { 'maintenance'; break }
            '^scripts/tests/[^/]+\.tests\.ps1$' { 'tests'; break }
            '^infrastructure/azure/(configure-analytics-oidc|deploy-appinsights)\.ps1$|^videos/excel-mcp-intro/Capture-Evidence\.ps1$' { 'safety'; break }
            '^docs/|^gh-pages/|^videos/|^infrastructure/|^\.changeset/|^\.github/|\.md$|^\.(gitignore|gitattributes)$' { 'documentation'; break }
            '^(package(-lock)?\.json|\.npmrc)$' { 'packages'; break }
            default { 'unknown' }
        }
        $plan.Reasons.Add("$path -> $kind")
        if ($kind -eq 'documentation') { continue }
        if ($kind -in @('runtime', 'cli', 'mcp', 'unknown')) {
            $plan.Build = $true
            $plan.Excel = $true
            $plan.SourceChecks = $true
        }
        if ($kind -in @('runtime', 'cli', 'cli-package', 'npm-packages', 'packages', 'pipeline', 'unknown', 'cli-docs', 'runtime-docs', 'distribution-docs')) { $plan.Cli = $true }
        if ($kind -in @('runtime', 'mcp', 'mcp-package', 'npm-packages', 'packages', 'pipeline', 'unknown', 'mcp-docs', 'runtime-docs', 'distribution-docs')) { $plan.Mcp = $true }
        if ($kind -in @('skills', 'skill-docs', 'packages', 'pipeline', 'unknown')) { $plan.Skills = $true }
        if ($kind -in @('extension', 'extension-docs', 'skills', 'packages', 'pipeline', 'runtime', 'mcp', 'unknown', 'distribution-docs')) { $plan.Extension = $true }
        if ($kind -in @('mcpb', 'packages', 'pipeline', 'unknown', 'distribution-docs')) { $plan.Mcpb = $true }
        if ($kind -in @('plugins', 'skills', 'packages', 'pipeline', 'unknown')) { $plan.Plugins = $true }
        if ($kind -in @('build', 'tests', 'skills', 'skill-docs', 'plugins', 'packages', 'safety', 'pipeline', 'doc-counts')) { $plan.Build = $true }
        if ($kind -in @('safety', 'pipeline') -or
            ($kind -eq 'tests' -and $path -notmatch '^tests/(ExcelMcp\.(SkillGeneration|Packaging)\.Tests/|Shared/(GeneratedAssetsFixture|PackagingScriptTestHelper)\.cs$)')) {
            $plan.HookTests = $true
        }
        if ($kind -in @('skills', 'skill-docs', 'pipeline') -or
            $path -match '^tests/(ExcelMcp\.SkillGeneration\.Tests/|Shared/(GeneratedAssetsFixture|PackagingScriptTestHelper)\.cs$)') {
            $plan.SkillTests = $true
        }
        if ($kind -in @('plugins', 'packages', 'mcpb', 'pipeline', 'doc-counts', 'cli-docs', 'mcp-docs', 'runtime-docs', 'distribution-docs', 'extension-docs') -or
            $path -match '^tests/(ExcelMcp\.Packaging\.Tests/|Shared/(GeneratedAssetsFixture|PackagingScriptTestHelper)\.cs$)' -or
            $path -match '^mcpb/(Build-McpBundle|McpbPackaging)\.ps1$') {
            $plan.PackagingTests = $true
            $plan.Build = $true
        }

        if ($kind -in @('runtime', 'cli', 'mcp', 'unknown', 'pipeline', 'build')) {
            foreach ($project in $allProjects) { [void]$fast.Add($project) }
            [void]$process.Add('CLI')
            if ($kind -in @('unknown', 'pipeline', 'build') -or
                $path -match '^(Directory\.|global\.json|NuGet\.Config|Sbroenne\.ExcelMcp\.sln)') { $fullTooling = $true }
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
        if ($path -match '^tests/ExcelMcp\.(CLI|ComInterop|Core|McpServer|Portable|Service)\.Tests/') {
            $project = $Matches[1]
            [void]$fast.Add($project)
            if ($project -eq 'CLI') { [void]$process.Add($project) }
        }
        if (($path -match '^tests/Shared/' -and
             $path -notmatch '^tests/Shared/(GeneratedAssetsFixture|PackagingScriptTestHelper)\.cs$') -or
            $path -match '^tests/Directory\.Build\.(props|targets)$') {
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
        if ($path -match '^scripts/(Invoke-TestStage|Get-ValidationPlan|Get-CiValidationPlan|Build-CiInputs|Invoke-ExcelFreeTests|Test-CiCompletion|Invoke-ExcelTests|Get-ExcelTestGroups)\.ps1$') {
            foreach ($project in $allProjects) { [void]$fast.Add($project) }
            [void]$process.Add('CLI')
            $fullTooling = $true
        }
        if ($kind -in @('skills', 'skill-docs')) {
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
        if ($kind -in @('cli-docs', 'mcp-docs', 'runtime-docs', 'distribution-docs', 'extension-docs')) {
            [void]$toolingOwners.Add('Packaging')
            [void]$tooling.Add('Feature=Packaging')
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
        foreach ($owner in $toolingOwners) {
            if (-not $toolingSelections.ContainsKey($owner)) {
                $toolingSelections[$owner] = [Collections.Generic.HashSet[string]]::new([StringComparer]::Ordinal)
            }
            foreach ($filter in $tooling) { [void]$toolingSelections[$owner].Add($filter) }
        }
        if ($kind -eq 'unknown') {
            foreach ($language in @('actions', 'csharp', 'javascript-typescript', 'python')) { [void]$codeQl.Add($language) }
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
        foreach ($language in @('actions', 'csharp', 'javascript-typescript', 'python')) { [void]$codeQl.Add($language) }
    }
    $plan.FastProjects = @($fast | Sort-Object)
    $plan.ProcessProjects = @($process | Sort-Object)
    $plan.ToolingProjects = if ($fullTooling) {
        @('SkillGeneration', 'Packaging', 'ScriptSafety')
    } else { @($toolingSelections.Keys | Sort-Object) }
    $plan.ToolingFilters = [ordered]@{}
    foreach ($owner in $plan.ToolingProjects) {
        $filters = @($toolingSelections[$owner] | Sort-Object)
        $plan.ToolingFilters[$owner] = if ($fullTooling -or 'RequiresExcel=false' -in $filters -or -not $filters.Count) {
            'RequiresExcel=false'
        } else { $filters -join '|' }
    }
    $plan.CodeQlLanguages = @($codeQl | Sort-Object)
    $plan.PackageBuild = [bool]($plan.Cli -or $plan.Mcp -or $plan.Extension)
    $plan.CiTestGroups = @(
        if ($fast.Count) { 'Fast' }
        if ($process.Count) { 'Process' }
        if ($plan.ToolingProjects.Count) { 'Tooling' }
    )
    $plan.ExcelGroups = @($excel | Sort-Object)
    $plan.SourceChecksGroup = if ($fast.Count -and ($plan.SourceChecks -or $fullTooling)) { 'Fast' } elseif ($documentationCounts) { 'Tooling' } else { '' }
    $plan.InfrastructureDiagnostics = $infrastructureDiagnostics
    $plan.Packages = [bool](@('Cli', 'Mcp', 'Extension', 'Mcpb', 'Skills', 'Plugins') | Where-Object { $plan[$_] }).Count
    $plan.NpmTests = $npmTests
    $plan.LockfileTests = $lockfileTests
    [pscustomobject]$plan
}
