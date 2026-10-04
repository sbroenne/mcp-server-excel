[CmdletBinding()]
param(
    [Parameter(Mandatory)][string]$PlanFile,
    [Parameter(Mandatory)][ValidateSet('Fast', 'Process', 'Tooling')][string]$Group,
    [switch]$ListProjects
)
$ErrorActionPreference = 'Stop'
$plan = Get-Content -LiteralPath $PlanFile -Raw | ConvertFrom-Json
if ($Group -notin $plan.CiTestGroups) { throw "The requested $Group build was not selected." }
if ($plan.SourceChecksGroup -eq $Group) {
    $projects = @('Sbroenne.ExcelMcp.sln')
} else {
    $owners = switch ($Group) {
        'Fast' { @($plan.FastProjects) }
        'Process' { @($plan.ProcessProjects) }
        'Tooling' { @($plan.ToolingProjects) }
    }
    $allowed = switch ($Group) {
        'Fast' { @('CLI', 'ComInterop', 'Core', 'McpServer', 'Service') }
        'Process' { @('CLI') }
        'Tooling' { @('Packaging', 'ScriptSafety', 'SkillGeneration') }
    }
    $projects = @(
        foreach ($owner in $owners | Sort-Object -Unique) {
            if ($owner -notin $allowed) { throw "Unexpected $Group project: $owner." }
            "tests\ExcelMcp.$owner.Tests\ExcelMcp.$owner.Tests.csproj"
        }
    )
}
if (-not $projects.Count) { throw "$Group has no build projects." }
if ($ListProjects) {
    $projects
} else {
    $root = Split-Path -Parent $PSScriptRoot
    foreach ($project in $projects) {
        $path = Join-Path $root $project
        & dotnet restore $path
        if ($LASTEXITCODE -ne 0) { throw "Restore failed: $project." }
        & dotnet build $path -c Release --no-restore --disable-build-servers -p:NuGetAudit=false
        if ($LASTEXITCODE -ne 0) { throw "Build failed: $project." }
    }
}
$global:LASTEXITCODE = 0
