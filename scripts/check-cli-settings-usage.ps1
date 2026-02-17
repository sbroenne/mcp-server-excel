#!/usr/bin/env pwsh
<#
.SYNOPSIS
    Checks that all CLI Settings properties are used in the args switch statements.

.DESCRIPTION
    Detects the pattern where a developer adds a Setting property but forgets to include it
    in the args switch statement, causing user values to be silently dropped.

    Example bug detected:
    - Settings has: public string? ConnectionString { get; init; }
    - Switch case has: new { connectionName, refreshOnOpen } // connectionString missing!

.EXAMPLE
    .\check-cli-settings-usage.ps1

.NOTES
    Part of pre-commit checks. Fails if any CLI command has unused Settings properties.
#>

$ErrorActionPreference = "Stop"
$rootDir = Split-Path -Parent $PSScriptRoot
$cliCommandsDir = Join-Path $rootDir "src\ExcelMcp.CLI\Commands"

# Properties that are legitimately not passed to daemon (session management, meta properties)
$globalExclusions = @(
    "Action",
    "SessionId"
)

# Properties that are defined for future features but daemon doesn't support yet
# These should be reviewed periodically to see if they can be implemented
$futureFeatureExclusions = @(
    "SheetScope",      # NamedRange - daemon doesn't support worksheet-scoped names yet
    "ModuleType",      # VBA - daemon uses auto-detection for module type
    "LayoutStyle",     # PivotTable - uses LayoutType instead
    "TargetPivotTableName", # Slicer - not implemented in daemon
    "Position",        # Chart - parsed from position string, not passed directly
    "TargetSheet",     # Chart - uses SheetName instead for placement
    "EnableRefresh"    # Connection - daemon uses RefreshOnFileOpen instead
)

# Properties that are used indirectly (files read into other properties)
$indirectUsagePatterns = @{
    "ConnectionStringFile" = "ConnectionString"
    "CommandTextFile" = "CommandText"
    "MCodeFile" = "MCode"
    "CodeFile" = "Code"
    "CsvFile" = "CsvData"
    "DaxQueryFile" = "DaxQuery"
    "DmvQueryFile" = "DmvQuery"
    "ExpressionFile" = "Expression"
    "ValuesFile" = "Values"
    "FormulasFile" = "Formulas"
    "FormatsFile" = "Formats"
}

$issues = @()
$totalChecked = 0
$totalPassed = 0

function Get-SettingsProperties {
    param([string]$content)

    $properties = @()
    # Match properties in Settings class: public string? PropertyName { get; init; }
    $settingsMatch = $content -match '(?s)internal sealed class Settings[^{]*\{(.+)$'
    if ($settingsMatch) {
        $settingsBlock = $Matches[1]
        # Extract property names
        $propertyMatches = [regex]::Matches($settingsBlock, 'public\s+\w+\??\s+(\w+)\s*\{')
        foreach ($match in $propertyMatches) {
            $properties += $match.Groups[1].Value
        }
    }
    return $properties
}

function Get-UsedProperties {
    param([string]$content)

    $usedProps = @()
    # Find all settings.PropertyName usages
    $usageMatches = [regex]::Matches($content, 'settings\.(\w+)')
    foreach ($match in $usageMatches) {
        $usedProps += $match.Groups[1].Value
    }
    return $usedProps | Sort-Object -Unique
}

Write-Host "Checking CLI Settings property usage..." -ForegroundColor Cyan
Write-Host ""

$commandFiles = Get-ChildItem -Path $cliCommandsDir -Filter "*Command.cs" -File

foreach ($file in $commandFiles) {
    # Skip ListActionsCommand - it's a meta command
    if ($file.Name -eq "ListActionsCommand.cs") {
        continue
    }

    $content = Get-Content $file.FullName -Raw
    $fileName = $file.Name

    # Skip if no Settings class
    if (-not ($content -match 'internal sealed class Settings')) {
        continue
    }

    $totalChecked++

    $settingsProps = Get-SettingsProperties $content
    $usedProps = Get-UsedProperties $content

    $unusedProps = @()
    foreach ($prop in $settingsProps) {
        # Skip global exclusions
        if ($globalExclusions -contains $prop) {
            continue
        }

        # Skip future feature exclusions
        if ($futureFeatureExclusions -contains $prop) {
            continue
        }

        # Skip indirect usage (file properties that populate other properties)
        if ($indirectUsagePatterns.ContainsKey($prop)) {
            continue
        }

        # Check if property is used
        if ($usedProps -notcontains $prop) {
            $unusedProps += $prop
        }
    }

    if ($unusedProps.Count -gt 0) {
        $issues += [PSCustomObject]@{
            File = $fileName
            UnusedProperties = $unusedProps -join ", "
        }
    }
    else {
        $totalPassed++
    }
}

if ($issues.Count -gt 0) {
    Write-Host "Found CLI commands with unused Settings properties:" -ForegroundColor Red
    Write-Host ""

    foreach ($issue in $issues) {
        Write-Host "   $($issue.File)" -ForegroundColor Yellow
        Write-Host "      Unused: $($issue.UnusedProperties)" -ForegroundColor Gray
    }

    Write-Host ""
    Write-Host "   These Settings properties are defined but NOT passed to daemon in args." -ForegroundColor Red
    Write-Host "   User values will be silently ignored!" -ForegroundColor Red
    Write-Host ""
    Write-Host "   Fix: Add property to args switch statement:" -ForegroundColor Cyan
    Write-Host "        ""action"" => new { ..., propertyName = settings.PropertyName }," -ForegroundColor White
    Write-Host ""
    exit 1
}

Write-Host "CLI Settings usage check passed - $totalPassed/$totalChecked commands" -ForegroundColor Green
Write-Host "   All Settings properties are used in args switch statements" -ForegroundColor Gray
exit 0
