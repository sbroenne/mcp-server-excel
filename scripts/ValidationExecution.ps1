#Requires -Version 7.4
. (Join-Path $PSScriptRoot 'Get-ValidationPlan.ps1')
. (Join-Path $PSScriptRoot 'Invoke-TestStage.ps1')

function Get-TestProject([string]$Owner) {
    if ($Owner -notin @('CLI', 'ComInterop', 'Core', 'McpServer', 'Service', 'Diagnostics', 'Packaging', 'ScriptSafety', 'SkillGeneration')) {
        throw "Unexpected test project: $Owner."
    }
    Join-Path (Split-Path -Parent $PSScriptRoot) "tests\ExcelMcp.$Owner.Tests\ExcelMcp.$Owner.Tests.csproj"
}

function Read-ValidationPlan([string]$Path) {
    $plan = Get-Content -LiteralPath $Path -Raw | ConvertFrom-Json -AsHashtable
    if ($plan.SchemaVersion -ne 1) { throw 'Unsupported validation plan schema.' }
    if ($plan.Excel -and -not $plan.ExcelSelections.Count -and -not $plan.FullE2E) {
        throw 'Excel was selected but the plan contains no test filters.'
    }
    foreach ($selection in $plan.ExcelSelections) {
        if ($selection.Project -notin @('CLI', 'ComInterop', 'Core', 'McpServer', 'Service') -or
            [string]::IsNullOrWhiteSpace($selection.Filter)) {
            throw 'Invalid saved Excel selection.'
        }
    }
    foreach ($kind in @('Fast', 'Process', 'Tooling')) {
        $allowed = switch ($kind) {
            Fast { @('CLI', 'ComInterop', 'Core', 'McpServer', 'Service', 'Diagnostics') }
            Process { @('CLI') }
            Tooling { @('Packaging', 'ScriptSafety', 'SkillGeneration') }
        }
        $filters = $plan["${kind}Filters"]
        if ($filters -isnot [Collections.IDictionary]) { throw "Missing $kind project filters." }
        if (@(Compare-Object @($plan["${kind}Projects"] | Sort-Object) @($filters.Keys | Sort-Object)).Count) {
            throw "Missing or unexpected $kind project filters."
        }
        foreach ($owner in $filters.Keys) {
            if ($owner -notin $allowed -or [string]::IsNullOrWhiteSpace($filters[$owner])) { throw "Invalid $kind selection for $owner." }
        }
    }
    foreach ($owner in $plan.ToolingOnDemandFilters.Keys) {
        if (-not $plan.ToolingFilters.ContainsKey($owner) -or [string]::IsNullOrWhiteSpace($plan.ToolingOnDemandFilters[$owner])) {
            throw 'Selected OnDemand cases require an owning Tooling project and a nonempty filter.'
        }
    }
    $groups = @('Fast', 'Process', 'Tooling' | Where-Object { $plan["${_}Filters"].Count })
    if (@(Compare-Object @($plan.CiTestGroups) $groups).Count) { throw 'Inconsistent saved test groups.' }
    return $plan
}

function Get-SelectedBuildProjects([Collections.IDictionary]$Plan, [string]$Group) {
    if ($Group -and $Group -notin @('Fast', 'Process', 'Tooling', 'Excel')) { throw "Unexpected test group: $Group." }
    if ($Group -and $Group -ne 'Excel' -and $Group -notin $Plan.CiTestGroups) { throw "The requested $Group group was not selected." }
    $projects = @(
        if ($Plan.FullSolutionBuild -and (-not $Group -or $Group -eq 'Excel' -or $Group -eq $Plan.FullSolutionBuildGroup)) {
            'Sbroenne.ExcelMcp.sln'
        } elseif (-not $Group) { $Plan.BuildProjects }
        elseif ($Group -eq 'Excel') { $Plan.ExcelSelections.Project | Sort-Object -Unique | ForEach-Object { Get-TestProject $_ } }
        else { $Plan["${Group}Projects"] | ForEach-Object { "tests\ExcelMcp.$_.Tests\ExcelMcp.$_.Tests.csproj" } }
    )
    if ($Group -and -not $projects.Count) { throw "$Group has no build projects." }
    return $projects
}

function Invoke-SelectedBuild([Collections.IDictionary]$Plan, [string]$Group) {
    $root = Split-Path -Parent $PSScriptRoot
    $paths = @(Get-SelectedBuildProjects $Plan $Group | ForEach-Object {
        $path = [IO.Path]::GetFullPath($_, $root)
        if (-not $path.StartsWith("$root\", [StringComparison]::OrdinalIgnoreCase) -or
            ([IO.Path]::GetExtension($path) -ne '.csproj' -and [IO.Path]::GetFileName($path) -ne 'Sbroenne.ExcelMcp.sln') -or
            -not (Test-Path -LiteralPath $path -PathType Leaf)) {
            throw "Build input must be an existing repository project: $_."
        }
        $path
    })
    if (-not $paths.Count) { return }
    if ($paths.Count -gt 1 -and @($paths | Where-Object { [IO.Path]::GetExtension($_) -ne '.csproj' }).Count) {
        throw 'Select either the complete solution or a set of projects, not both.'
    }
    $filter = $null
    $primary = $null
    try {
        $target = $paths[0]
        if ($paths.Count -gt 1) {
            $filter = Join-Path ([IO.Path]::GetTempPath()) "ExcelMcp.SelectedBuild.$([Guid]::NewGuid().ToString('N')).slnf"
            @{ solution = @{ path = Join-Path $root 'Sbroenne.ExcelMcp.sln'
                projects = @($paths | ForEach-Object { [IO.Path]::GetRelativePath($root, $_) }) } } |
                ConvertTo-Json -Depth 5 | Set-Content -LiteralPath $filter
            $target = $filter
        }
        foreach ($arguments in @(@('restore', $target), @('build', $target, '-c', 'Release', '--no-restore',
                    '-p:UseSharedCompilation=false', '--verbosity', 'minimal'))) {
            $result = Invoke-NativeCommand -Executable dotnet -Arguments $arguments -DeadlineSeconds 1200 `
                -Environment @{ MSBUILDDISABLENODEREUSE = '1' }
            Write-Host $result.Output
            if ($result.Error) { [Console]::Error.WriteLine($result.Error) }
        }
    } catch { $primary = $_.Exception; throw }
    finally {
        if ($filter) {
            try { Remove-Item -LiteralPath $filter -ErrorAction Stop }
            catch {
                if ($primary) { throw [AggregateException]::new('Selected build and temporary-input cleanup failed.', [Exception[]]@($primary, $_.Exception)) }
                throw
            }
        }
    }
}

function Get-FreeTestSelection {
    [CmdletBinding()]
    param(
        [Collections.IDictionary]$Plan,
        [string]$Group,
        [switch]$Local,
        [switch]$Contracts,
        [switch]$HookTests,
        [switch]$SkillTests,
        [switch]$PackagingTests
    )
    $filters = [ordered]@{}
    $onDemand = @{}
    if ($Group) {
        $commit = $Group -eq 'Commit'
        $kind = if ($commit) { 'Fast' } else { $Group }
        if ($kind -notin @('Fast', 'Process', 'Tooling') -or $kind -notin $Plan.CiTestGroups) {
            throw "The requested $kind group was not selected."
        }
        foreach ($owner in $Plan["${kind}Projects"]) {
            if ($commit -and $owner -notin @('CLI', 'Core', 'McpServer', 'Service')) { continue }
            $filter = $Plan["${kind}Filters"][$owner]
            if ([string]::IsNullOrWhiteSpace($filter)) { throw "Missing filter for $owner." }
            $filters[$owner] = switch ($kind) {
                Fast { "AdapterTestKind!=System&$(if ($commit) { 'Speed!=Slow&' })($filter)" }
                Process { "AdapterTestKind=System&($filter)" }
                Tooling { $filter }
            }
            if ($kind -eq 'Tooling' -and $Plan.ToolingOnDemandFilters.ContainsKey($owner)) {
                $onDemand[$owner] = $Plan.ToolingOnDemandFilters[$owner]
            }
        }
    } elseif ($Local) {
        foreach ($kind in @('Fast', 'Process', 'Tooling')) {
            foreach ($owner in $Plan["${kind}Projects"]) {
                $selected = $Plan["${kind}Filters"][$owner]
                if ($filters.Contains($owner)) { $filters[$owner] += "|$selected" } else { $filters[$owner] = $selected }
            }
        }
        foreach ($entry in @(@('ScriptSafety', $HookTests), @('SkillGeneration', $SkillTests), @('Packaging', $PackagingTests))) {
            if ($entry[1]) { $filters[$entry[0]] = 'RequiresExcel=false' }
        }
        if ($Contracts) { foreach ($owner in @('Core', 'CLI', 'McpServer')) {
            if ($filters.Contains($owner)) { $filters[$owner] += '|Feature=GeneratedContracts' } else { $filters[$owner] = 'Feature=GeneratedContracts' }
        } }
        $onDemand = $Plan.ToolingOnDemandFilters
    } else {
        foreach ($owner in @('CLI', 'ComInterop', 'Core', 'McpServer', 'Service', 'SkillGeneration', 'Packaging', 'ScriptSafety')) {
            $filters[$owner] = 'RequiresExcel=false'
        }
    }
    foreach ($owner in $filters.Keys) {
        $filter = $filters[$owner]
        if ($onDemand.ContainsKey($owner)) {
            if ([string]::IsNullOrWhiteSpace($onDemand[$owner])) { throw 'Selected OnDemand cases require a nonempty filter.' }
            $filter = "RequiresExcel=false&((RunType!=OnDemand&($filter))|(RunType=OnDemand&($($onDemand[$owner]))))"
        } else { $filter = "RequiresExcel=false&RunType!=OnDemand&($filter)" }
        [pscustomobject]@{ Owner = $owner; Filter = $filter }
    }
}

function Invoke-SelectedFreeTests {
    [CmdletBinding()]
    param([Collections.IDictionary]$Plan, [string]$Group, [string]$ResultsDirectory,
        [switch]$ListTests, [int]$DeadlineSeconds = 1800)
    foreach ($item in @(Get-FreeTestSelection -Plan $Plan -Group $Group)) {
        Invoke-TestStage -Project (Get-TestProject $item.Owner) -Filter $item.Filter `
            -ResultsDirectory $ResultsDirectory -Name $item.Owner -ListTests:$ListTests -DeadlineSeconds $DeadlineSeconds
    }
}

function Invoke-ExcelPlan {
    [CmdletBinding()]
    param([Collections.IDictionary]$Plan, [string]$ResultsDirectory, [switch]$ListTests, [int]$DeadlineSeconds = 7200)
    if (-not $Plan.Excel) { throw 'The requested Excel group was not selected.' }
    $root = Split-Path -Parent $PSScriptRoot
    $pipe = "excelmcp-selected-$PID-$([Guid]::NewGuid().ToString('N'))"
    $primary = $null
    try {
        foreach ($project in @($Plan.ExcelSelections | Group-Object Project)) {
            $filter = ($project.Group.Filter | ForEach-Object { $_ -split '\|' } | Sort-Object -Unique) -join '|'
            $required = if ($Plan.FullE2E) { '&Acceptance!=Required' } else { '' }
            $ordinary = @($project.Group | Where-Object { -not ($Plan.FullE2E -and $_.Area -eq 'Acceptance') })
            if (-not $ordinary.Count) { continue }
            Invoke-TestStage -Project (Get-TestProject $project.Name) -Filter "RequiresExcel=true&RunType!=OnDemand$required&($filter)" `
                -ResultsDirectory (Join-Path $ResultsDirectory Excel) -Name $project.Name -DeadlineSeconds $DeadlineSeconds `
                -HangTimeout 10m -Environment @{ EXCELMCP_CLI_PIPE = $pipe } -ListTests:$ListTests -ReconcileCases
        }
        if ($Plan.InfrastructureDiagnostics) {
            Invoke-TestStage -Project (Get-TestProject ComInterop) `
                -Filter 'RequiresExcel=true&RunType=OnDemand&FullyQualifiedName!~BeginBatch_RealIrmWorkbook&Locale!=ja-JP' `
                -ResultsDirectory (Join-Path $ResultsDirectory Infrastructure) -Name ComInterop -DeadlineSeconds $DeadlineSeconds `
                -HangTimeout 10m -Environment @{ EXCELMCP_CLI_PIPE = $pipe } -ListTests:$ListTests -ReconcileCases
        }
        if ($Plan.FullE2E -and -not $ListTests) {
            $result = Invoke-NativeCommand -Executable pwsh -Arguments @('-NoProfile', '-File', (Join-Path $root 'scripts\Test-E2E.ps1'),
                '-SkipBuild', '-PipeName', $pipe, '-ResultsDirectory', (Join-Path $ResultsDirectory Acceptance)) -DeadlineSeconds 10800
            Write-Host $result.Output
            if ($result.Error) { [Console]::Error.WriteLine($result.Error) }
        }
    } catch { $primary = $_.Exception; throw }
    finally {
        if (-not $ListTests -and (@($Plan.ExcelSelections | Where-Object Project -eq CLI).Count -or $Plan.FullE2E)) {
            try { & (Join-Path $root 'scripts\Stop-ExcelCliService.ps1') -PipeName $pipe }
            catch {
                if ($primary) { throw [AggregateException]::new('Selected tests and owned cleanup failed.', [Exception[]]@($primary, $_.Exception)) }
                throw
            }
        }
    }
}
