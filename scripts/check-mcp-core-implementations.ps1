#!/usr/bin/env pwsh
<#
.SYNOPSIS
    Validates that all MCP Tool enum actions have corresponding Core Command implementations

.DESCRIPTION
    Checks that every action defined in ToolActions.cs enums has a matching method
    in the corresponding Core Commands interface (IPowerQueryCommands, ISheetCommands, etc.)

    This prevents situations where MCP tools expose actions that don't exist in Core,
    which would cause runtime exceptions when called.

.EXAMPLE
    .\check-mcp-core-implementations.ps1

.NOTES
    Part of pre-commit validation to catch missing Core implementations early
#>

$ErrorActionPreference = "Stop"
$rootDir = Split-Path -Parent $PSScriptRoot

Write-Host "Checking MCP Tool actions have Core implementations..." -ForegroundColor Cyan

$errors = @()

function Get-InterfaceMethodNames {
    param(
        [string] $InterfaceContent
    )

    $pattern = '(?m)^\s*(?:\[.*\]\s*)*(?:public\s+)?(?:static\s+)?(?:async\s+)?[\w<>\[\],?\.]+\s+(\w+)\s*\('
    $methodMatches = [System.Text.RegularExpressions.Regex]::Matches($InterfaceContent, $pattern)
    $names = @()

    foreach ($match in $methodMatches) {
        $name = $match.Groups[1].Value
        if (-not [string]::IsNullOrWhiteSpace($name)) {
            $names += $name
        }
    }

    return $names
}

# Known intentional exceptions (documented in CORE-METHOD-RENAMING-SUMMARY.md)
$knownExceptions = @{
    "FileAction" = @("CloseWorkbook", "Open", "Save", "Close", "List", "CreateAndOpen")  # Session management actions (MCP-specific)
    "TableAction" = @("ApplyFilterValues", "SortMulti")  # Composite operations
    "TableColumnAction" = @("ApplyFilterValues", "SortMulti")  # Composite operations
    "RangeAction" = @("SetNumberFormatCustom")  # Maps to SetNumberFormat Core method (intentional name difference for LLM usability)
}

# Define mappings: Enum -> Core Interface File(s)
$mappings = @{
    "FileAction" = @("src/ExcelMcp.Core/Commands/IFileCommands.cs")
    "PowerQueryAction" = @("src/ExcelMcp.Core/Commands/PowerQuery/IPowerQueryCommands.cs")
    "WorksheetAction" = @("src/ExcelMcp.Core/Commands/Sheet/ISheetCommands.cs")
    "WorksheetStyleAction" = @("src/ExcelMcp.Core/Commands/Sheet/ISheetCommands.cs")
    "RangeAction" = @("src/ExcelMcp.Core/Commands/Range/IRangeCommands.cs")
    "RangeEditAction" = @("src/ExcelMcp.Core/Commands/Range/IRangeCommands.cs")
    "RangeFormatAction" = @("src/ExcelMcp.Core/Commands/Range/IRangeCommands.cs")
    "RangeLinkAction" = @("src/ExcelMcp.Core/Commands/Range/IRangeCommands.cs")
    "TableAction" = @("src/ExcelMcp.Core/Commands/Table/ITableCommands.cs")
    "TableColumnAction" = @("src/ExcelMcp.Core/Commands/Table/ITableCommands.cs")
    "DataModelAction" = @("src/ExcelMcp.Core/Commands/DataModel/IDataModelCommands.cs")
    "DataModelRelAction" = @("src/ExcelMcp.Core/Commands/DataModel/IDataModelCommands.cs")
    "VbaAction" = @("src/ExcelMcp.Core/Commands/Vba/IVbaCommands.cs")
    "ConnectionAction" = @("src/ExcelMcp.Core/Commands/Connection/IConnectionCommands.cs")
    "NamedRangeAction" = @("src/ExcelMcp.Core/Commands/NamedRange/INamedRangeCommands.cs")
    "PivotTableAction" = @("src/ExcelMcp.Core/Commands/PivotTable/IPivotTableCommands.cs")
    "PivotTableFieldAction" = @("src/ExcelMcp.Core/Commands/PivotTable/IPivotTableCommands.cs")
    "PivotTableCalcAction" = @("src/ExcelMcp.Core/Commands/PivotTable/IPivotTableCommands.cs")
    "ChartAction" = @("src/ExcelMcp.Core/Commands/Chart/IChartCommands.cs")
    "ChartConfigAction" = @("src/ExcelMcp.Core/Commands/Chart/IChartCommands.cs")
    "SlicerAction" = @(
        "src/ExcelMcp.Core/Commands/PivotTable/IPivotTableCommands.cs",
        "src/ExcelMcp.Core/Commands/Table/ITableCommands.cs"
    )
}

# Read ToolActions.cs to extract enum values
$toolActionsFile = Join-Path $rootDir "src/ExcelMcp.Core/Models/Actions/ToolActions.cs"
$toolActionsContent = Get-Content $toolActionsFile -Raw

foreach ($enumName in $mappings.Keys) {
    $interfaceFiles = $mappings[$enumName] | ForEach-Object { Join-Path $rootDir $_ }

    # Check if Core interface file exists
    $existingInterfaces = $interfaceFiles | Where-Object { Test-Path $_ }
    if ($existingInterfaces.Count -eq 0) {
        $missing = $interfaceFiles -join ", "
        Write-Host "  ⚠️  Warning: Core interface(s) not found: $missing" -ForegroundColor Yellow
        continue
    }

    # Extract enum values using regex
    $enumPattern = "enum $enumName\s*\{([^}]+)\}"
    if ($toolActionsContent -match $enumPattern) {
        $enumBody = $Matches[1]
        $enumValues = $enumBody -split ',' | ForEach-Object {
            $_.Trim() -replace '//.*$', '' | Where-Object { $_ -match '^\w+$' }
        }

        # Read Core interfaces and extract method names
        $methodSet = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
        foreach ($interfacePath in $existingInterfaces) {
            $interfaceContent = Get-Content $interfacePath -Raw
            $methodNames = Get-InterfaceMethodNames -InterfaceContent $interfaceContent
            foreach ($methodName in $methodNames) {
                [void]$methodSet.Add($methodName)
            }
        }

        # Check each enum value has a corresponding method (async or sync)
        $missingMethods = @()
        foreach ($value in $enumValues) {
            # Skip known exceptions
            if ($knownExceptions.ContainsKey($enumName) -and $knownExceptions[$enumName] -contains $value) {
                continue
            }

            $expectedNames = @("${value}Async", $value)
            $methodExists = $false

            foreach ($candidate in $expectedNames) {
                if ($methodSet.Contains($candidate)) {
                    $methodExists = $true
                    break
                }
            }

            if (-not $methodExists) {
                $missingMethods += $value
            }
        }

        if ($missingMethods.Count -gt 0) {
            $errors += "❌ $enumName has actions without Core implementations:"
            foreach ($missing in $missingMethods) {
                $errors += "   - $missing (expected ${missing}Async in $($mappings[$enumName] -join ', '))"
            }
        } else {
            Write-Host "  ✅ $enumName - all $($enumValues.Count) actions have Core implementations" -ForegroundColor Green
        }
    }
}

if ($errors.Count -gt 0) {
    Write-Host ""
    Write-Host "Validation Failed:" -ForegroundColor Red
    $errors | ForEach-Object { Write-Host $_ -ForegroundColor Red }
    Write-Host ""
    Write-Host "Action Required:" -ForegroundColor Yellow
    Write-Host "  1. Remove the enum value from ToolActions.cs if not needed" -ForegroundColor Yellow
    Write-Host "  2. Or implement the missing method in the Core interface" -ForegroundColor Yellow
    Write-Host "  3. Update ActionExtensions.cs to remove the mapping" -ForegroundColor Yellow
    Write-Host "  4. Update the MCP Tool switch statement to remove the case" -ForegroundColor Yellow
    exit 1
}

Write-Host ""
Write-Host "✅ All MCP Tool actions have Core implementations" -ForegroundColor Green
exit 0
