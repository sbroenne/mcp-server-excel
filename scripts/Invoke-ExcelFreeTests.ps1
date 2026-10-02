[CmdletBinding()]
param(
    [switch]$Local,
    [switch]$HookTests,
    [switch]$Contracts,
    [switch]$SkillTests,
    [string[]]$ChangedPaths = @(),
    [ValidateSet('Fast', 'Process', 'Tooling')][string]$Group,
    [string]$PlanFile,
    [string]$ResultsDirectory
)
$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $PSScriptRoot
. (Join-Path $PSScriptRoot 'Invoke-TestStage.ps1')
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
        'Tooling' { $selections['SkillGeneration'] = $plan.ToolingFilter }
    }
}
elseif ($PlanFile) { throw 'PlanFile requires an explicit Group.' }
elseif ($Local) {
    if ($HookTests) { $selections['SkillGeneration'] = 'Feature=PreCommit|Feature=AutomationSafety' }
    if ($SkillTests) {
        $selections['SkillGeneration'] = if ($selections['SkillGeneration']) {
            "$($selections['SkillGeneration'])|Feature=SkillGeneration"
        } else { 'Feature=SkillGeneration' }
    }
    if ($Contracts) {
        $selections['Core'] = 'Feature=GeneratedContracts'
        $selections['CLI'] = 'FullyQualifiedName~GeneratedActionContractCliTests'
        $selections['McpServer'] = 'FullyQualifiedName~McpToolSurfaceTests|FullyQualifiedName~CalculationGuidanceContractTests|FullyQualifiedName~GeneratedMcpParameterTests'
    }
    foreach ($path in $ChangedPaths) {
        if ($path -match '(PluginPublication|Publish-PreparedPlugins|PluginContent|AwesomeCopilotPolicy|Update-AwesomeCopilot|update-awesome-copilot|publish-plugins)') {
            $selections['SkillGeneration'] = if ($selections['SkillGeneration']) {
                "$($selections['SkillGeneration'])|Feature=PluginPublication"
            } else { 'Feature=PluginPublication' }
        }
        if ($path -match '^tests[/\\]ExcelMcp\.(Core|CLI|ComInterop|McpServer|Service)\.Tests[/\\]') {
            $selections[$Matches[1]] = 'RequiresExcel=false'
        }
    }
}
else {
    foreach ($project in @('CLI', 'ComInterop', 'Core', 'McpServer', 'Service', 'SkillGeneration')) {
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
