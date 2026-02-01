#!/usr/bin/env pwsh
<#
.SYNOPSIS
    Check that CLI command switch statements handle all action strings from ActionExtensions.cs

.DESCRIPTION
    This script verifies that each CLI Command's "action switch" statement handles all the
    action strings defined in ActionExtensions.cs for the corresponding enum. This prevents
    bugs like ChartCommand missing "create-from-table" case (found 2026-02-01).

.EXAMPLE
    .\scripts\check-cli-action-coverage.ps1
    .\scripts\check-cli-action-coverage.ps1 -Verbose

.NOTES
    Created to catch missing CLI switch cases that check-cli-coverage.ps1 misses.
    That script only checks if a CLI command EXISTS for each enum.
    This script checks if the CLI command HANDLES all actions for that enum.
#>

param(
    [switch]$Verbose
)

$ErrorActionPreference = "Stop"
$rootDir = Split-Path -Parent $PSScriptRoot

Write-Host "🔍 CLI Action Switch Coverage Check" -ForegroundColor Cyan
Write-Host "====================================" -ForegroundColor Cyan
Write-Host ""

$actionExtensionsPath = Join-Path $rootDir "src\ExcelMcp.Core\Models\Actions\ActionExtensions.cs"
$cliCommandsDir = Join-Path $rootDir "src\ExcelMcp.CLI\Commands"

if (-not (Test-Path $actionExtensionsPath)) {
    Write-Host "❌ ActionExtensions.cs not found: $actionExtensionsPath" -ForegroundColor Red
    exit 1
}

if (-not (Test-Path $cliCommandsDir)) {
    Write-Host "❌ CLI Commands directory not found: $cliCommandsDir" -ForegroundColor Red
    exit 1
}

# Map enum names to CLI command file names
# Pattern: ChartAction -> ChartCommand.cs
$enumToCliCommand = @{
    # Direct mappings (enum name matches command name)
    "ChartAction" = "ChartCommand.cs"
    "ChartConfigAction" = "ChartConfigCommand.cs"
    "ConnectionAction" = "ConnectionCommand.cs"
    "ConditionalFormatAction" = "ConditionalFormatCommand.cs"
    "DataModelAction" = "DataModelCommand.cs"
    "DataModelRelAction" = "DataModelRelCommand.cs"
    "NamedRangeAction" = "NamedRangeCommand.cs"
    "PowerQueryAction" = "PowerQueryCommand.cs"
    "PivotTableAction" = "PivotTableCommand.cs"
    "SlicerAction" = "SlicerCommand.cs"
    "TableAction" = "TableCommand.cs"
    "VbaAction" = "VbaCommand.cs"

    # Naming exceptions
    "WorksheetAction" = "SheetCommand.cs"
    "WorksheetStyleAction" = "SheetCommand.cs"  # Style actions handled via sheet command with --style flag

    # Range variants all use RangeCommand
    "RangeAction" = "RangeCommand.cs"
    "RangeEditAction" = "RangeCommand.cs"
    "RangeFormatAction" = "RangeCommand.cs"
    "RangeLinkAction" = "RangeCommand.cs"

    # PivotTable variants
    "PivotTableFieldAction" = "PivotTableCommand.cs"
    "PivotTableCalcAction" = "PivotTableCommand.cs"

    # Table variants
    "TableColumnAction" = "TableCommand.cs"
}

# Enums that don't have direct CLI switch coverage (handled differently)
$excludedEnums = @(
    "FileAction"  # CLI uses session subcommands, not action switch
)

# Actions that can safely fall through to default (no specific args required)
# ALL other actions MUST have explicit switch cases
$actionsAllowingDefault = @{
    "ChartAction" = @("list", "read")
    "ChartConfigAction" = @()  # All require specific args
    "PowerQueryAction" = @("list", "refresh-all")
    "WorksheetAction" = @("list")
    "WorksheetStyleAction" = @()  # All require specific args
    "TableAction" = @("list")
    "TableColumnAction" = @()  # All require specific args
    "VbaAction" = @("list")
    "ConnectionAction" = @("list")
    "DataModelAction" = @("list-tables", "read-info", "refresh")
    "DataModelRelAction" = @("list-relationships")
    "NamedRangeAction" = @("list")
    "SlicerAction" = @("list-slicers", "list-table-slicers")
    "ConditionalFormatAction" = @()  # All require specific args
    "PivotTableAction" = @("list")
    "PivotTableFieldAction" = @()  # All require specific args
    "PivotTableCalcAction" = @()  # All require specific args
    "RangeAction" = @("get-used-range")
    "RangeEditAction" = @()  # All require specific args
    "RangeFormatAction" = @()  # All require specific args
    "RangeLinkAction" = @()  # All require specific args
}

# Parse ActionExtensions.cs to extract action strings for each enum
function Get-ActionStringsFromExtensions {
    param([string]$Content)

    $result = @{}

    # Match each ToActionString method and extract enum name and action mappings
    # Pattern: public static string ToActionString(this EnumName action) => action switch { ... };
    $methodPattern = 'public\s+static\s+string\s+ToActionString\(this\s+(\w+Action)\s+action\)\s*=>\s*action\s+switch\s*\{([^}]+)\}'
    $matches = [regex]::Matches($Content, $methodPattern, [System.Text.RegularExpressions.RegexOptions]::Singleline)

    foreach ($match in $matches) {
        $enumName = $match.Groups[1].Value
        $switchBody = $match.Groups[2].Value

        # Extract action strings: EnumName.Value => "action-string"
        $actionPattern = '\w+\.\w+\s*=>\s*"([^"]+)"'
        $actionMatches = [regex]::Matches($switchBody, $actionPattern)

        $actions = @()
        foreach ($actionMatch in $actionMatches) {
            $actions += $actionMatch.Groups[1].Value
        }

        $result[$enumName] = $actions
    }

    return $result
}

# Parse CLI Command file to extract handled action strings from switch
function Get-HandledActionsFromCommand {
    param([string]$FilePath)

    if (-not (Test-Path $FilePath)) {
        return @()
    }

    $content = Get-Content $FilePath -Raw

    # Find the "action switch" block
    # Pattern: object? args = action switch { ... };
    $switchPattern = 'object\?\s+args\s*=\s*action\s+switch\s*\{([^;]+)\};'
    $match = [regex]::Match($content, $switchPattern, [System.Text.RegularExpressions.RegexOptions]::Singleline)

    if (-not $match.Success) {
        return @()
    }

    $switchBody = $match.Groups[1].Value

    # Extract action strings: "action-string" => ...
    $actionPattern = '"([^"]+)"\s*=>'
    $actionMatches = [regex]::Matches($switchBody, $actionPattern)

    $actions = @()
    foreach ($actionMatch in $actionMatches) {
        $actions += $actionMatch.Groups[1].Value
    }

    return $actions
}

# Read ActionExtensions.cs
$extensionsContent = Get-Content $actionExtensionsPath -Raw
$enumActionStrings = Get-ActionStringsFromExtensions -Content $extensionsContent

Write-Host "Found $($enumActionStrings.Count) action enums in ActionExtensions.cs" -ForegroundColor Gray
Write-Host ""

$totalMissing = 0
$totalChecked = 0

foreach ($enumName in $enumActionStrings.Keys | Sort-Object) {
    # Skip excluded enums
    if ($excludedEnums -contains $enumName) {
        if ($Verbose) {
            Write-Host "⏭️  $enumName → (excluded, CLI handles differently)" -ForegroundColor Gray
        }
        continue
    }

    # Get CLI command file
    if (-not $enumToCliCommand.ContainsKey($enumName)) {
        Write-Host "⚠️  $enumName → No CLI mapping defined (add to `$enumToCliCommand)" -ForegroundColor Yellow
        continue
    }

    $commandFile = $enumToCliCommand[$enumName]
    $commandPath = Join-Path $cliCommandsDir $commandFile

    if (-not (Test-Path $commandPath)) {
        Write-Host "⚠️  $enumName → CLI file not found: $commandFile" -ForegroundColor Yellow
        continue
    }

    # Get expected and actual actions
    $expectedActions = $enumActionStrings[$enumName]
    $handledActions = Get-HandledActionsFromCommand -FilePath $commandPath

    # Filter out actions that can safely fall through to default (list, read, etc.)
    if ($actionsAllowingDefault.ContainsKey($enumName)) {
        $allowedDefaults = $actionsAllowingDefault[$enumName]
        $expectedActions = $expectedActions | Where-Object { $_ -notin $allowedDefaults }
    }

    # Find missing actions
    $missingActions = $expectedActions | Where-Object { $_ -notin $handledActions }

    $totalChecked++

    if ($missingActions.Count -gt 0) {
        $totalMissing += $missingActions.Count
        Write-Host "❌ $enumName → $commandFile" -ForegroundColor Red
        foreach ($missing in $missingActions) {
            Write-Host "     Missing: `"$missing`"" -ForegroundColor Yellow
        }
    } else {
        if ($Verbose) {
            Write-Host "✅ $enumName → $commandFile ($($expectedActions.Count) actions)" -ForegroundColor Green
        }
    }
}

Write-Host ""

if ($totalMissing -gt 0) {
    Write-Host "❌ MISSING CLI SWITCH CASES DETECTED!" -ForegroundColor Red
    Write-Host "   Total missing: $totalMissing cases across $totalChecked enums checked" -ForegroundColor Yellow
    Write-Host ""
    Write-Host "Action Required:" -ForegroundColor Yellow
    Write-Host "  1. Add missing case to the CLI command's 'action switch' statement" -ForegroundColor Yellow
    Write-Host "  2. Include all required args for that action (match daemon's *Args class)" -ForegroundColor Yellow
    Write-Host "  3. Example: `"create-from-table`" => new { tableName = settings.TableName, ... }" -ForegroundColor Yellow
    Write-Host ""
    exit 1
}

Write-Host "✅ All $totalChecked CLI commands handle their action strings!" -ForegroundColor Green
exit 0
