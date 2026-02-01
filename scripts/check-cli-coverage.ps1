#!/usr/bin/env pwsh
<#
.SYNOPSIS
    Check that all action enums have corresponding CLI commands registered.

.DESCRIPTION
    Auto-discovers action enums from ToolActions.cs and verifies each has a CLI command
    registered in Program.cs. Fails if any enums are missing CLI commands.

.EXAMPLE
    .\scripts\check-cli-coverage.ps1

.NOTES
    Added to pre-commit to prevent CLI parity gaps (like DataModelRelAction missing).
#>

param(
    [switch]$Verbose
)

$ErrorActionPreference = "Stop"
$rootDir = Split-Path -Parent $PSScriptRoot

Write-Host "🔍 CLI Command Coverage Check" -ForegroundColor Cyan
Write-Host "==============================" -ForegroundColor Cyan
Write-Host ""

$toolActionsPath = Join-Path $rootDir "src\ExcelMcp.Core\Models\Actions\ToolActions.cs"
$programPath = Join-Path $rootDir "src\ExcelMcp.CLI\Program.cs"

if (-not (Test-Path $toolActionsPath)) {
    Write-Host "❌ ToolActions.cs not found: $toolActionsPath" -ForegroundColor Red
    exit 1
}

if (-not (Test-Path $programPath)) {
    Write-Host "❌ Program.cs not found: $programPath" -ForegroundColor Red
    exit 1
}

# Discover all action enums from ToolActions.cs
$toolActionsContent = Get-Content $toolActionsPath -Raw
$enumPattern = "public\s+enum\s+(\w+Action)\s*\{"
$enumMatches = [regex]::Matches($toolActionsContent, $enumPattern)

$allEnums = @()
foreach ($match in $enumMatches) {
    $allEnums += $match.Groups[1].Value
}

Write-Host "Found $($allEnums.Count) action enums in ToolActions.cs" -ForegroundColor Gray

# Map enum names to expected CLI command names
# Pattern: PowerQueryAction -> powerquery, DataModelRelAction -> datamodelrel
function Get-ExpectedCliCommand {
    param([string]$EnumName)

    # Remove 'Action' suffix and convert to lowercase
    $baseName = $EnumName -replace 'Action$', ''
    return $baseName.ToLowerInvariant()
}

# Sub-tool enums that share CLI commands with parent tools (no separate CLI command needed)
$subToolEnums = @{
    # Range sub-tools share 'range' command
    "RangeEditAction" = "range"
    "RangeFormatAction" = "range"
    "RangeLinkAction" = "range"

    # Worksheet sub-tools share 'sheet' command
    "WorksheetStyleAction" = "sheet"

    # PivotTable sub-tools share 'pivottable' command
    "PivotTableFieldAction" = "pivottable"
    "PivotTableCalcAction" = "pivottable"
    "SlicerAction" = "pivottable"  # Slicer is exposed via pivottable and table commands

    # Table sub-tools share 'table' command
    "TableColumnAction" = "table"
}

# Special naming exceptions (enum name doesn't match CLI command name)
$namingExceptions = @{
    "WorksheetAction" = "sheet"
    "ConditionalFormatAction" = "conditionalformat"
}

# MCP-specific enums that have no CLI equivalent (handled differently in CLI)
$mcpOnlyEnums = @(
    "FileAction"  # CLI uses session subcommands (session create/open/close/save) instead
)


# Read Program.cs and find registered commands
$programContent = Get-Content $programPath -Raw

# Extract registered CLI commands: config.AddCommand<*Command>("commandname")
$commandPattern = 'config\.AddCommand<\w+Command>\("(\w+)"\)'
$commandMatches = [regex]::Matches($programContent, $commandPattern)

$registeredCommands = @()
foreach ($match in $commandMatches) {
    $registeredCommands += $match.Groups[1].Value.ToLowerInvariant()
}

Write-Host "Found $($registeredCommands.Count) CLI commands registered in Program.cs" -ForegroundColor Gray
Write-Host ""

# Check each enum has a CLI command
$missingCommands = @()
$coveredEnums = @()

foreach ($enum in $allEnums) {
    # Skip MCP-only enums (no CLI equivalent by design)
    if ($mcpOnlyEnums -contains $enum) {
        if ($Verbose) {
            Write-Host "⏭️  $enum → (MCP-only, no CLI equivalent)" -ForegroundColor Gray
        }
        continue
    }

    # Check if this is a sub-tool enum
    if ($subToolEnums.ContainsKey($enum)) {
        $expectedCommand = $subToolEnums[$enum]
        if ($registeredCommands -contains $expectedCommand) {
            $coveredEnums += $enum
            if ($Verbose) {
                Write-Host "✅ $enum → $expectedCommand (sub-tool)" -ForegroundColor Green
            }
        } else {
            $missingCommands += @{ Enum = $enum; Expected = $expectedCommand; Reason = "sub-tool parent missing" }
        }
        continue
    }

    # Check naming exceptions
    if ($namingExceptions.ContainsKey($enum)) {
        $expectedCommand = $namingExceptions[$enum]
    } else {
        $expectedCommand = Get-ExpectedCliCommand -EnumName $enum
    }

    if ($registeredCommands -contains $expectedCommand) {
        $coveredEnums += $enum
        if ($Verbose) {
            Write-Host "✅ $enum → $expectedCommand" -ForegroundColor Green
        }
    } else {
        $missingCommands += @{ Enum = $enum; Expected = $expectedCommand; Reason = "no CLI command" }
    }
}

# Report results
if ($missingCommands.Count -gt 0) {
    Write-Host "❌ MISSING CLI COMMANDS DETECTED!" -ForegroundColor Red
    Write-Host ""

    foreach ($missing in $missingCommands) {
        Write-Host "   $($missing.Enum) → expected CLI command: '$($missing.Expected)'" -ForegroundColor Yellow
    }

    Write-Host ""
    Write-Host "Action Required:" -ForegroundColor Yellow
    Write-Host "  1. Create Command class: src/ExcelMcp.CLI/Commands/<Name>Command.cs" -ForegroundColor Yellow
    Write-Host "  2. Register in Program.cs: config.AddCommand<*Command>(""name"")" -ForegroundColor Yellow
    Write-Host "  3. Or add to subToolEnums/namingExceptions in this script if intentional" -ForegroundColor Yellow
    Write-Host ""

    exit 1
}

Write-Host "✅ All $($allEnums.Count) action enums have CLI commands!" -ForegroundColor Green
Write-Host "   Covered: $($coveredEnums.Count) (including $($subToolEnums.Count) sub-tool enums)" -ForegroundColor Gray
exit 0
