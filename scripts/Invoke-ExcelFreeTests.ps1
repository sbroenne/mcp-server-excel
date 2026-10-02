[CmdletBinding()]
param(
    [switch]$Local,
    [switch]$HookTests,
    [switch]$Contracts,
    [switch]$SkillTests,
    [switch]$PackagingTests,
    [string[]]$ChangedPaths = @(),
    [ValidateSet('Fast', 'Process', 'Tooling')][string]$Group,
    [string]$PlanFile,
    [string]$ResultsDirectory
)
$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $PSScriptRoot
. (Join-Path $PSScriptRoot 'Invoke-TestStage.ps1')
. (Join-Path $PSScriptRoot 'Get-ValidationPlan.ps1')
$selections = [ordered]@{}
if ($Group) {
    $plan = if ($PlanFile) {
        Get-Content -LiteralPath $PlanFile -Raw | ConvertFrom-Json
    } else {
        . (Join-Path $PSScriptRoot 'Get-ValidationPlan.ps1')
        Get-ValidationPlan -Full
    }
    if ($Group -notin $plan.CiTestGroups) { throw "The requested $Group group was not selected." }
    switch ($Group) {
        'Fast' {
            foreach ($project in $plan.FastProjects) { $selections[$project] = 'AdapterTestKind!=System' }
        }
        'Process' {
            foreach ($project in $plan.ProcessProjects) {
                $selections[$project] = 'AdapterTestKind=System'
            }
        }
        'Tooling' {
            foreach ($project in $plan.ToolingProjects) { $selections[$project] = $plan.ToolingFilter }
        }
    }
}
elseif ($PlanFile) { throw 'PlanFile requires an explicit Group.' }
elseif ($Local) {
    $plan = Get-ValidationPlan -Paths $ChangedPaths
    if ($HookTests -or $plan.HookTests) { $selections['ScriptSafety'] = 'RequiresExcel=false' }
    if ($SkillTests -or $plan.SkillTests) { $selections['SkillGeneration'] = 'Feature=SkillGeneration' }
    if ($PackagingTests -or $plan.PackagingTests) { $selections['Packaging'] = 'RequiresExcel=false' }
    if ($Contracts) {
        $selections['Core'] = 'Feature=GeneratedContracts'
        $selections['CLI'] = 'FullyQualifiedName~GeneratedActionContractCliTests'
        $selections['McpServer'] = 'FullyQualifiedName~McpToolSurfaceTests|FullyQualifiedName~CalculationGuidanceContractTests|FullyQualifiedName~GeneratedMcpParameterTests'
    }
    foreach ($path in $ChangedPaths) {
        if ($path -match '^tests[/\\]ExcelMcp\.(Core|CLI|ComInterop|McpServer|Service|SkillGeneration|Packaging|ScriptSafety)\.Tests[/\\]') {
            $selections[$Matches[1]] = 'RequiresExcel=false'
        }
        if ($path -match '^tests[/\\]Shared[/\\]' -and $path -notmatch '[/\\](GeneratedAssetsFixture|PackagingScriptTestHelper)\.cs$') {
            foreach ($project in @('CLI', 'ComInterop', 'Core', 'McpServer', 'Service', 'SkillGeneration', 'Packaging', 'ScriptSafety')) {
                $selections[$project] = 'RequiresExcel=false'
            }
        }
    }
}
else {
    foreach ($project in @('CLI', 'ComInterop', 'Core', 'McpServer', 'Service', 'SkillGeneration', 'Packaging', 'ScriptSafety')) {
        $selections[$project] = 'RequiresExcel=false'
    }
}
if (-not $ResultsDirectory) { $ResultsDirectory = Join-Path $root "TestResults\excel-free-$([Guid]::NewGuid().ToString('N'))" }
if ($selections.Count -eq 0) {
    if ($Group) { throw "$Group has no selected test projects." }
    Write-Host 'No local Excel-free tests selected.'
}
foreach ($entry in $selections.GetEnumerator()) {
    $project = Join-Path $root "tests\ExcelMcp.$($entry.Key).Tests\ExcelMcp.$($entry.Key).Tests.csproj"
    $filter = "RequiresExcel=false&RunType!=OnDemand&($($entry.Value))"
    Invoke-TestStage -Project $project -Filter $filter -ResultsDirectory $ResultsDirectory -Name $entry.Key
}
$global:LASTEXITCODE = 0
