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
        Plugins = $false
        Reasons = [Collections.Generic.List[string]]::new()
    }
    $fast = [Collections.Generic.HashSet[string]]::new([StringComparer]::Ordinal)
    $process = [Collections.Generic.HashSet[string]]::new([StringComparer]::Ordinal)
    $tooling = [Collections.Generic.HashSet[string]]::new([StringComparer]::Ordinal)
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
            '^scripts/(Test-E2E|Test-CliWorkflow|Stop-ExcelMcpProcesses)\.ps1$|^tests/.*/(PreBuildGracefulSaveAcceptanceTests|McpServerSmokeTests|CliWorkflowAcceptanceTests)\.cs$' { 'runtime'; break }
            '^tests/' { 'tests'; break }
            '^vscode-extension/' { 'extension'; break }
            '^mcpb/' { 'mcpb'; break }
            '^npm-packages/excelcli' { 'cli-package'; break }
            '^npm-packages/mcp-server-excel' { 'mcp-package'; break }
            '^npm-packages/shared/' { 'npm-packages'; break }
            '^\.github/plugins/|^\.github/workflows/(publish-plugins\.yml|update-awesome-copilot\.(md|lock\.yml))$' { 'plugins'; break }
            '^scripts/Build-AgentSkills\.ps1$' { 'skills'; break }
            '^scripts/(Build-Plugins|Sync-PublishedPluginRepo|Publish-PreparedPlugins)\.ps1$|^scripts/(PluginContent|AwesomeCopilotPolicy|Update-AwesomeCopilot)\.mjs$' { 'plugins'; break }
            '^scripts/(Build-NpmPackages|Test-NpmPackages|Build-ReleasePackages|PackageHelpers)\.ps1$|^\.github/workflows/release\.yml$' { 'packages'; break }
            '^doc-counts\.json$|^scripts/check-doc-counts\.ps1$' { 'doc-counts'; break }
            '^scripts/(pre-commit|Get-ValidationPlan|Get-CiValidationPlan|Get-ExcelTestGroups|Invoke-ExcelFreeTests|Invoke-ExcelTests|Invoke-TestStage|Test-CiCompletion|check-|Test-NpmLockfiles)' { 'tests'; break }
            '^\.github/workflows/ci\.yml$' { 'pipeline'; break }
            '^scripts/(Build-Changelog|Update-(ReleaseVersion|McpRegistry)Metadata)\.ps1$' { 'packages'; break }
            '^scripts/(Update|Restore|Persist|Test)-StarHistory\.ps1$|^scripts/.*UsageAnalytics.*\.ps1$' { 'maintenance'; break }
            '^docs/|^gh-pages/|^videos/|^infrastructure/|^specs/|^\.changeset/|^\.github/|\.md$|^\.(gitignore|gitattributes)$' { 'documentation'; break }
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
        if ($kind -in @('build', 'tests', 'skills', 'plugins', 'pipeline', 'doc-counts')) { $plan.Build = $true }
        if ($kind -in @('tests', 'pipeline')) { $plan.HookTests = $true }
        if ($kind -eq 'skills' -or $path -match '^tests/ExcelMcp\.SkillGeneration\.Tests/') { $plan.SkillTests = $true }

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
        if ($path -match '^tests/Shared/|^tests/.*\.csproj$|^tests/.*xunit\.runner\.json$') {
            foreach ($project in $allProjects) { [void]$fast.Add($project) }
            [void]$process.Add('CLI')
            $fullTooling = $true
            foreach ($group in $allExcel) { [void]$excel.Add($group) }
            $infrastructureDiagnostics = $true
        }
        if ($kind -eq 'tests' -and $path -match '^scripts/') {
            foreach ($feature in @('PreCommit', 'AutomationSafety')) { [void]$tooling.Add("Feature=$feature") }
        }
        if ($path -match '^scripts/(Invoke-TestStage|Get-CiValidationPlan|Test-CiCompletion|Invoke-ExcelTests|Get-ExcelTestGroups)\.ps1$') {
            foreach ($project in $allProjects) { [void]$fast.Add($project) }
            [void]$process.Add('CLI')
            $fullTooling = $true
        }
        if ($kind -eq 'skills') { [void]$tooling.Add('Feature=SkillGeneration') }
        if ($kind -eq 'plugins') { [void]$tooling.Add('Feature=PluginPublication') }
        if ($kind -eq 'doc-counts') {
            $documentationCounts = $true
            [void]$tooling.Add('FullyQualifiedName~DocumentationCounts')
        }
        if ($kind -in @('packages', 'mcpb')) {
            foreach ($feature in @('Packaging', 'McpbPackaging', 'PluginSkillVersion', 'ReleaseMetadata', 'PluginPublication')) {
                [void]$tooling.Add("Feature=$feature")
            }
        }
        if ($path -match '^tests/ExcelMcp\.SkillGeneration\.Tests/') {
            if ($path -match 'PluginPublication') { [void]$tooling.Add('Feature=PluginPublication') }
            else { $fullTooling = $true }
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
    $plan.CiTestGroups = @(
        if ($fast.Count) { 'Fast' }
        if ($process.Count) { 'Process' }
        if ($fullTooling -or $tooling.Count) { 'Tooling' }
    )
    $plan.ExcelGroups = @($excel | Sort-Object)
    $plan.SourceChecksGroup = if ($fast.Count) { 'Fast' } elseif ($documentationCounts) { 'Tooling' } else { '' }
    $plan.InfrastructureDiagnostics = $infrastructureDiagnostics
    $plan.Packages = [bool](@('Cli', 'Mcp', 'Extension', 'Mcpb', 'Skills', 'Plugins') | Where-Object { $plan[$_] }).Count
    $plan.NpmTests = $npmTests
    $plan.LockfileTests = $lockfileTests
    [pscustomobject]$plan
}
