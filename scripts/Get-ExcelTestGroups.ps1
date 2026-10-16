#Requires -Version 7.4
. (Join-Path $PSScriptRoot 'ValidationExecution.ps1')

function Get-ExcelTestGroups {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$ResultsDirectory)
    $directory = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($ResultsDirectory)
    foreach ($owner in @('Core', 'Service', 'CLI', 'McpServer', 'ComInterop')) {
        $path = Join-Path $directory "$owner-inventory.json"
        if (Test-Path -LiteralPath $path) { throw "Use a new inventory path: $path." }
        Invoke-TestStage -Project (Get-TestProject $owner) -Filter 'FullyQualifiedName~ExcelValidationGroups_PreserveClassFixturesAndExportBuiltInventory' `
            -Name "$owner-inventory" -ResultsDirectory $directory -DeadlineSeconds 120 -Environment @{ EXCELMCP_TEST_SELECTION_OUTPUT = $path }
        if (-not (Test-Path -LiteralPath $path -PathType Leaf)) { throw "Missing built test inventory: $path." }
        $cases = @(Get-Content -LiteralPath $path -Raw | ConvertFrom-Json)
        if (-not $cases.Count) { throw "Empty built test inventory: $path." }
        foreach ($class in @($cases | Group-Object Class)) {
            $groups = @($class.Group | Group-Object Group)
            foreach ($group in $groups) {
                $filter = if ($groups.Count -eq 1) { "FullyQualifiedName~$($class.Name)." }
                else { ($group.Group.Method | ForEach-Object { "FullyQualifiedName=$_" }) -join '|' }
                [pscustomobject]@{
                    Project = Get-TestProject $owner; ProjectName = $owner; Class = $class.Name
                    Group = $group.Name; Filter = "($filter)"; Cases = @($group.Group)
                }
            }
        }
    }
}
