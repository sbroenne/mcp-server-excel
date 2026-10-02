function Get-ExcelTestGroups {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$ResultsDirectory)

    $ResultsDirectory = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($ResultsDirectory)
    . (Join-Path $PSScriptRoot 'Invoke-TestStage.ps1')
    $root = Split-Path -Parent $PSScriptRoot
    foreach ($project in @('Core', 'Service', 'CLI', 'McpServer', 'ComInterop')) {
        $path = Join-Path $root "tests\ExcelMcp.$project.Tests\ExcelMcp.$project.Tests.csproj"
        $inventory = Join-Path $ResultsDirectory "$project-inventory.json"
        Invoke-TestStage -Project $path `
            -Filter 'FullyQualifiedName~ExcelValidationGroups_PreserveClassFixturesAndExportBuiltInventory' `
            -ResultsDirectory $ResultsDirectory -Name "$project-inventory" -DeadlineSeconds 120 `
            -Environment @{ EXCELMCP_TEST_SELECTION_OUTPUT = $inventory } | Out-Host
        if (-not (Test-Path -LiteralPath $inventory)) { throw "Missing built test inventory: $inventory" }
        $tests = @(Get-Content -LiteralPath $inventory -Raw | ConvertFrom-Json)
        foreach ($class in $tests | Group-Object Class) {
            $groups = @($class.Group.Group | Sort-Object -Unique)
            foreach ($group in $groups) {
                $cases = @($class.Group | Where-Object Group -eq $group)
                $filter = if ($groups.Count -eq 1) {
                    "FullyQualifiedName~$($class.Name)."
                } else {
                    ($cases | ForEach-Object { "FullyQualifiedName=$($_.Method)" }) -join '|'
                }
                [pscustomobject]@{
                    Project = $path
                    ProjectName = $project
                    Class = $class.Name
                    Group = $group
                    Filter = "($filter)"
                    Cases = $cases
                }
            }
        }
    }
}
