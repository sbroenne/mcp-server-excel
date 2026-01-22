#!/usr/bin/env pwsh
<#
.SYNOPSIS
    Audit script to verify Core Commands coverage in MCP Server

.DESCRIPTION
    Counts Core interface methods vs MCP Server enum values to detect gaps.
    Run quarterly or before major releases to ensure 100% coverage is maintained.

.EXAMPLE
    .\audit-core-coverage.ps1

.NOTES
    Author: ExcelMcp Team
    Created: 2025-01-28
    Purpose: Prevent Core Commands from being added without MCP Server exposure
#>

param(
    [switch]$Verbose,
    [switch]$FailOnGaps,
    [switch]$CheckNaming
)

$ErrorActionPreference = "Stop"
$rootDir = Split-Path -Parent $PSScriptRoot

Write-Host "🔍 Core Commands Coverage Audit" -ForegroundColor Cyan
Write-Host "=================================" -ForegroundColor Cyan
Write-Host ""

# Function to count unique async method names in Core interface files (handles overloads)
function Get-CoreMethodMatches {
    param([string]$InterfacePath)

    if (-not (Test-Path $InterfacePath)) {
        return @()
    }

    $content = Get-Content $InterfacePath -Raw

    # Match interface method signatures, e.g., "OperationResult Create(...)" or "Task<OperationResult> CreateAsync(...)"
    $pattern = '^[\s\t]*(?:[\w<>,\[\]\? ]+)\s+(?<name>\w+)\s*\([^;]*\)\s*;'
    $methodMatches = [regex]::Matches($content, $pattern, [System.Text.RegularExpressions.RegexOptions]::Multiline)

    $methodNames = @()
    foreach ($match in $methodMatches) {
        $name = $match.Groups['name'].Value
        if ($methodNames -notcontains $name) {
            $methodNames += $name
        }
    }

    return $methodNames
}

function Count-CoreMethods {
    param([string]$InterfacePath, [string]$InterfaceName)

    if (-not (Test-Path $InterfacePath)) {
        Write-Warning "Interface file not found: $InterfacePath"
        return 0
    }

    $methodNames = Get-CoreMethodMatches -InterfacePath $InterfacePath
    return $methodNames.Count
}

# Function to count enum values
function Count-EnumValues {
    param([string]$EnumName, [string]$ToolActionsPath)

    if (-not (Test-Path $ToolActionsPath)) {
        Write-Warning "ToolActions.cs not found: $ToolActionsPath"
        return 0
    }

    $content = Get-Content $ToolActionsPath -Raw
    # Find the enum definition
    $enumPattern = "public\s+enum\s+$EnumName\s*\{([^}]+)\}"
    if ($content -match $enumPattern) {
        $enumBody = $Matches[1]
        # Count non-empty, non-comment lines
        $lines = $enumBody -split "`n" | Where-Object {
            $_ -match '\S' -and $_ -notmatch '^\s*//'
        }
        return $lines.Count
    }

    return 0
}

# Function to count enum values for a specific interface (handles cross-interface enum splits)
function Count-EnumValuesForInterface {
    param(
        [string]$EnumName,
        [string]$InterfaceName,
        [string]$ToolActionsPath
    )

    # Check if this enum has a cross-interface split defined
    if ($Script:crossInterfaceEnumSplits -and $Script:crossInterfaceEnumSplits.ContainsKey($EnumName)) {
        $splits = $Script:crossInterfaceEnumSplits[$EnumName]
        if ($splits.ContainsKey($InterfaceName)) {
            # Return count of values specific to this interface
            return $splits[$InterfaceName].Count
        }
    }

    # No split defined - return full enum count
    return Count-EnumValues -EnumName $EnumName -ToolActionsPath $ToolActionsPath
}

# Function to extract unique method names from Core interface (without "Async" suffix, handles overloads)
function Get-CoreMethodNames {
    param([string]$InterfacePath)

    return Get-CoreMethodMatches -InterfacePath $InterfacePath
}

# Function to extract enum value names
function Get-EnumValueNames {
    param([string]$EnumName, [string]$ToolActionsPath)

    if (-not (Test-Path $ToolActionsPath)) {
        return @()
    }

    $content = Get-Content $ToolActionsPath -Raw
    $enumPattern = "public\s+enum\s+$EnumName\s*\{([^}]+)\}"
    if ($content -match $enumPattern) {
        $enumBody = $Matches[1]
        $enumValues = @()
        $lines = $enumBody -split "`n" | Where-Object {
            $_ -match '^\s*(\w+)' -and $_ -notmatch '^\s*//'
        }
        foreach ($line in $lines) {
            if ($line -match '^\s*(\w+)') {
                $enumValues += $Matches[1]
            }
        }
        return $enumValues
    }

    return @()
}

# Function to check naming consistency
function Check-NamingConsistency {
    param(
        [string]$InterfaceName,
        [string]$InterfacePath,
        [string]$EnumName,
        [string]$ToolActionsPath
    )

    $methodNames = Get-CoreMethodNames -InterfacePath $InterfacePath
    $enumValues = Get-EnumValueNames -EnumName $EnumName -ToolActionsPath $ToolActionsPath

    $mismatches = @()

    # Check each method has matching enum
    foreach ($method in $methodNames) {
        if ($enumValues -notcontains $method) {
            $mismatches += "Method '$method' has no matching enum value"
        }
    }

    # Check each enum has matching method
    foreach ($enum in $enumValues) {
        if ($methodNames -notcontains $enum) {
            $mismatches += "Enum '$enum' has no matching method"
        }
    }

    return $mismatches
}

# Discover all enum types from ToolActions.cs
function Get-AllEnumTypes {
    param([string]$ToolActionsPath)

    if (-not (Test-Path $ToolActionsPath)) {
        return @()
    }

    $content = Get-Content $ToolActionsPath -Raw
    $enumPattern = "public\s+enum\s+(\w+Action)\s*\{"
    $enumMatches = [regex]::Matches($content, $enumPattern)

    $enumTypes = @()
    foreach ($match in $enumMatches) {
        $enumTypes += $match.Groups[1].Value
    }

    return $enumTypes
}

# Discover interface files dynamically
function Find-InterfaceForEnum {
    param(
        [string]$EnumType,
        [string]$CommandsPath
    )

    # Map enum type to expected interface name
    # Pattern: PowerQueryAction -> IPowerQueryCommands
    # Special cases and sub-tool mappings

    $enumToInterface = @{
        # Known naming exceptions
        "WorksheetAction" = "ISheetCommands"
        "ConditionalFormatAction" = "IConditionalFormattingCommands"

        # Sub-tool enums that map to parent interfaces
        # Range sub-tools (all map to IRangeCommands)
        "RangeEditAction" = "IRangeCommands"
        "RangeFormatAction" = "IRangeCommands"
        "RangeLinkAction" = "IRangeCommands"

        # Worksheet sub-tools (all map to ISheetCommands)
        "WorksheetStyleAction" = "ISheetCommands"

        # DataModel sub-tools (all map to IDataModelCommands)
        "DataModelRelAction" = "IDataModelCommands"

        # Table sub-tools (all map to ITableCommands)
        "TableColumnAction" = "ITableCommands"

        # PivotTable sub-tools (all map to IPivotTableCommands)
        "PivotTableFieldAction" = "IPivotTableCommands"
        "PivotTableCalcAction" = "IPivotTableCommands"

        # Cross-interface enums (cover methods from multiple interfaces)
        # SlicerAction covers methods from BOTH IPivotTableCommands AND ITableCommands
        # We map to IPivotTableCommands as primary, and add ITableCommands below in additionalEnumMappings
        "SlicerAction" = "IPivotTableCommands"

        # Chart sub-tools (all map to IChartCommands)
        "ChartConfigAction" = "IChartCommands"
    }

    # Additional interface mappings for cross-interface enums
    # These enums have methods implemented in multiple Core interfaces
    # Format: "InterfaceName" = @("EnumName1", "EnumName2", ...)
    $Script:additionalEnumMappings = @{
        "ITableCommands" = @("SlicerAction")  # Table slicer methods exposed via SlicerAction
    }

    # Cross-interface enum value splits
    # When an enum covers methods from MULTIPLE interfaces, specify which values belong to each
    # Format: "EnumName" = @{ "InterfaceName" = @("Value1", "Value2", ...) }
    $Script:crossInterfaceEnumSplits = @{
        "SlicerAction" = @{
            # PivotTable slicer actions (4 values)
            "IPivotTableCommands" = @("CreateSlicer", "ListSlicers", "SetSlicerSelection", "DeleteSlicer")
            # Table slicer actions (4 values)
            "ITableCommands" = @("CreateTableSlicer", "ListTableSlicers", "SetTableSlicerSelection", "DeleteTableSlicer")
        }
    }

    if ($enumToInterface.ContainsKey($EnumType)) {
        $interfaceName = $enumToInterface[$EnumType]
    } else {
        # Standard pattern: {Name}Action -> I{Name}Commands
        $baseName = $EnumType -replace 'Action$', ''
        $interfaceName = "I${baseName}Commands"
    }

    # Search recursively for interface file
    $interfaceFiles = Get-ChildItem -Path $CommandsPath -Recurse -Filter "$interfaceName.cs"

    if ($interfaceFiles.Count -eq 0) {
        return $null
    }

    # Return the first match (should be only one)
    return @{
        Name = $interfaceName
        Path = $interfaceFiles[0].FullName
        Enum = $EnumType
    }
}

$toolActionsPath = "$rootDir/src/ExcelMcp.McpServer/Models/ToolActions.cs"

# Dynamically discover all interfaces to check
$commandsPath = Join-Path $rootDir "src\ExcelMcp.Core\Commands"
$enumTypes = Get-AllEnumTypes -ToolActionsPath $toolActionsPath

$interfaces = @()
foreach ($enumType in $enumTypes) {
    $interface = Find-InterfaceForEnum -EnumType $enumType -CommandsPath $commandsPath
    if ($interface) {
        $interfaces += $interface
    } else {
        Write-Warning "No interface found for enum type: $enumType"
    }
}

# Group interfaces by interface name (multiple enums can map to same interface)
$groupedInterfaces = @{}
foreach ($interface in $interfaces) {
    $key = $interface.Name
    if (-not $groupedInterfaces.ContainsKey($key)) {
        $groupedInterfaces[$key] = @{
            Name = $interface.Name
            Path = $interface.Path
            Enums = @()
        }
    }
    $groupedInterfaces[$key].Enums += $interface.Enum
}

# Add additional enum mappings for cross-interface enums
# This handles cases like SlicerAction which covers methods from both IPivotTableCommands and ITableCommands
if ($Script:additionalEnumMappings) {
    foreach ($interfaceName in $Script:additionalEnumMappings.Keys) {
        if ($groupedInterfaces.ContainsKey($interfaceName)) {
            $additionalEnums = $Script:additionalEnumMappings[$interfaceName]
            foreach ($enumName in $additionalEnums) {
                if ($groupedInterfaces[$interfaceName].Enums -notcontains $enumName) {
                    $groupedInterfaces[$interfaceName].Enums += $enumName
                }
            }
        }
    }
}

# Track results
$results = @()
$totalCoreMethods = 0
$totalEnumValues = 0
$hasGaps = $false

# Audit each interface (aggregating all related enums)
foreach ($key in $groupedInterfaces.Keys) {
    $interfaceGroup = $groupedInterfaces[$key]
    $coreMethods = Count-CoreMethods -InterfacePath $interfaceGroup.Path -InterfaceName $interfaceGroup.Name

    # Sum enum values across ALL enums that map to this interface
    $totalEnumValuesForInterface = 0
    $enumNames = @()
    foreach ($enumName in $interfaceGroup.Enums) {
        # Use interface-aware counting for cross-interface enums (e.g., SlicerAction)
        $enumCount = Count-EnumValuesForInterface -EnumName $enumName -InterfaceName $interfaceGroup.Name -ToolActionsPath $toolActionsPath
        $totalEnumValuesForInterface += $enumCount
        $enumNames += "$enumName($enumCount)"
    }

    $totalCoreMethods += $coreMethods
    $totalEnumValues += $totalEnumValuesForInterface

    $status = "✅"
    $statusText = "OK"

    if ($totalEnumValuesForInterface -lt $coreMethods) {
        $status = "❌"
        $statusText = "GAP"
        $hasGaps = $true
    } elseif ($totalEnumValuesForInterface -gt $coreMethods) {
        $status = "⚠️"
        $statusText = "EXTRA"
    }

    $result = [PSCustomObject]@{
        Interface = $interfaceGroup.Name
        CoreMethods = $coreMethods
        EnumValues = $totalEnumValuesForInterface
        Enums = ($interfaceGroup.Enums -join ", ")
        Gap = $coreMethods - $totalEnumValuesForInterface
        Status = $status
        StatusText = $statusText
    }

    $results += $result

    if ($Verbose) {
        Write-Host "Checking $($interfaceGroup.Name)..." -ForegroundColor Gray
        Write-Host "  Core Methods: $coreMethods" -ForegroundColor Gray
        Write-Host "  Enum Values: $totalEnumValuesForInterface (from: $($enumNames -join ', '))" -ForegroundColor Gray
        Write-Host "  Status: $status $statusText" -ForegroundColor $(if ($statusText -eq "OK") { "Green" } elseif ($statusText -eq "GAP") { "Red" } else { "Yellow" })
        Write-Host ""
    }
}

# Display results table
Write-Host ""
Write-Host "Audit Results:" -ForegroundColor Cyan
Write-Host ""
$results | Format-Table -Property Interface, CoreMethods, EnumValues, Enums, Gap, Status -AutoSize

# Summary
Write-Host ""
Write-Host "Summary:" -ForegroundColor Cyan
Write-Host "--------" -ForegroundColor Cyan
Write-Host "Total Core Methods: $totalCoreMethods" -ForegroundColor White
Write-Host "Total Enum Values:  $totalEnumValues" -ForegroundColor White

if ($totalCoreMethods -eq 0) {
    Write-Host "Coverage:           N/A (no core methods detected)" -ForegroundColor Yellow
} elseif ($totalEnumValues -eq $totalCoreMethods) {
    Write-Host "Coverage:           100% ✅" -ForegroundColor Green
} else {
    $coverage = [math]::Round(($totalEnumValues / $totalCoreMethods) * 100, 1)
    Write-Host "Coverage:           $coverage%" -ForegroundColor $(if ($coverage -ge 95) { "Yellow" } else { "Red" })
}

# Gaps detection
if ($hasGaps) {
    Write-Host ""
    Write-Host "⚠️  GAPS DETECTED!" -ForegroundColor Red
    Write-Host ""
    Write-Host "The following interfaces have fewer enum values than Core methods:" -ForegroundColor Red
    $results | Where-Object { $_.Gap -gt 0 } | ForEach-Object {
        Write-Host "  - $($_.Interface): Missing $($_.Gap) enum values" -ForegroundColor Red
    }
    Write-Host ""
    Write-Host "Action Required:" -ForegroundColor Yellow
    Write-Host "  1. Review Core interface for new methods" -ForegroundColor Yellow
    Write-Host "  2. Add missing enum values to ToolActions.cs" -ForegroundColor Yellow
    Write-Host "  3. Add ToActionString mappings to ActionExtensions.cs" -ForegroundColor Yellow
    Write-Host "  4. Add switch cases to appropriate MCP Tools" -ForegroundColor Yellow
    Write-Host "  5. See .github/instructions/coverage-prevention-strategy.instructions.md" -ForegroundColor Yellow

    if ($FailOnGaps) {
        exit 1
    }
} else {
    Write-Host ""
    Write-Host "✅ No gaps detected - 100% coverage maintained!" -ForegroundColor Green
}

# Extra enum values warning
$extraEnums = $results | Where-Object { $_.Gap -lt 0 }
if ($extraEnums.Count -gt 0) {
    Write-Host ""
    Write-Host "⚠️  Note: Some enums have more values than Core methods" -ForegroundColor Yellow
    Write-Host "This might be intentional (MCP-specific actions like 'close-workbook')" -ForegroundColor Gray
    $extraEnums | ForEach-Object {
        Write-Host "  - $($_.Interface): $([math]::Abs($_.Gap)) extra enum values" -ForegroundColor Yellow
    }
}

Write-Host ""
Write-Host "Audit completed at $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')" -ForegroundColor Gray

# Explicitly exit with success code (no gaps detected)
if ($FailOnGaps -and $hasGaps) {
    exit 1
}

# Naming consistency check (if requested)
if ($CheckNaming) {
    Write-Host ""
    Write-Host "🔤 Naming Consistency Check" -ForegroundColor Cyan
    Write-Host "===========================" -ForegroundColor Cyan
    Write-Host ""

    # Sub-tool enums are intentionally subsets of parent interface - skip naming check
    # These enums only contain a subset of the parent interface methods by design
    $subToolEnums = @(
        "RangeEditAction", "RangeFormatAction", "RangeLinkAction",  # IRangeCommands sub-tools
        "WorksheetStyleAction",  # ISheetCommands sub-tools
        "DataModelRelAction",  # IDataModelCommands sub-tools
        "TableColumnAction",  # ITableCommands sub-tools
        "PivotTableFieldAction", "PivotTableCalcAction", "SlicerAction",  # IPivotTableCommands sub-tools
        "ChartConfigAction"  # IChartCommands sub-tools
    )

    # Known intentional exceptions (documented in CORE-METHOD-RENAMING-SUMMARY.md)
    # Also includes methods that moved to sub-tool enums
    $knownExceptions = @{
        "TableAction" = @("ApplyFilterValues", "SortMulti", "ApplyFilter", "ClearFilters", "GetFilters",
                          "AddColumn", "RemoveColumn", "RenameColumn", "GetStructuredReference", "Sort",
                          "GetColumnNumberFormat", "SetColumnNumberFormat",
                          # Table slicer methods exposed via SlicerAction (cross-interface enum)
                          "CreateTableSlicer", "ListTableSlicers", "SetTableSlicerSelection", "DeleteTableSlicer")
        "FileAction" = @("CloseWorkbook", "Open", "Save", "Close", "List")  # MCP-specific session actions
        "RangeAction" = @("SetNumberFormatCustom", "InsertCells", "DeleteCells", "InsertRows", "DeleteRows",
                          "InsertColumns", "DeleteColumns", "Find", "Replace", "Sort",
                          "AddHyperlink", "RemoveHyperlink", "ListHyperlinks", "GetHyperlink",
                          "SetStyle", "GetStyle", "FormatRange", "ValidateRange", "GetValidation", "RemoveValidation",
                          "AutoFitColumns", "AutoFitRows", "MergeCells", "UnmergeCells", "GetMergeInfo",
                          "SetCellLock", "GetCellLock")  # Methods moved to RangeEdit/RangeFormat/RangeLink tools
        "WorksheetAction" = @("SetTabColor", "GetTabColor", "ClearTabColor", "SetVisibility", "GetVisibility",
                              "Show", "Hide", "VeryHide")  # Methods moved to WorksheetStyleAction
        "DataModelAction" = @("ListRelationships", "ReadRelationship", "DeleteRelationship",
                              "CreateRelationship", "UpdateRelationship")  # Methods moved to DataModelRelAction
        "PivotTableAction" = @("ListFields", "AddRowField", "AddColumnField", "AddValueField", "AddFilterField",
                               "RemoveField", "SetFieldFunction", "SetFieldName", "SetFieldFormat", "GetData",
                               "SetFieldFilter", "SortField", "GroupByDate", "GroupByNumeric",
                               "CreateCalculatedField", "ListCalculatedFields", "DeleteCalculatedField",
                               "SetLayout", "SetSubtotals", "SetGrandTotals",
                               "ListCalculatedMembers", "CreateCalculatedMember", "DeleteCalculatedMember",
                               "CreateSlicer", "ListSlicers", "SetSlicerSelection", "DeleteSlicer")  # Methods moved to PivotTableField/PivotTableCalc/Slicer
        "ChartAction" = @("SetSourceRange", "AddSeries", "RemoveSeries", "SetChartType", "SetTitle",
                          "SetAxisTitle", "GetAxisNumberFormat", "SetAxisNumberFormat", "ShowLegend", "SetStyle")  # Methods moved to ChartConfigAction
    }

    $hasNamingIssues = $false

    foreach ($interface in $interfaces) {
        # Skip sub-tool enums - they are intentionally subsets
        if ($subToolEnums -contains $interface.Enum) {
            Write-Host "⏭️  $($interface.Name) → $($interface.Enum): Skipped (sub-tool enum)" -ForegroundColor Gray
            continue
        }

        $mismatches = Check-NamingConsistency `
            -InterfaceName $interface.Name `
            -InterfacePath $interface.Path `
            -EnumName $interface.Enum `
            -ToolActionsPath $toolActionsPath

        # Filter out known exceptions
        if ($knownExceptions.ContainsKey($interface.Enum)) {
            $exceptions = $knownExceptions[$interface.Enum]
            $mismatches = $mismatches | Where-Object {
                $mismatch = $_
                # Match both "Method 'X' has no matching..." and "Enum 'X' has no matching..."
                -not ($exceptions | Where-Object { $mismatch -like "*'$_'*" })
            }
        }

        if ($mismatches.Count -gt 0) {
            $hasNamingIssues = $true
            Write-Host "❌ $($interface.Name) → $($interface.Enum):" -ForegroundColor Red
            foreach ($mismatch in $mismatches) {
                Write-Host "   $mismatch" -ForegroundColor Yellow
            }
            Write-Host ""
        } else {
            Write-Host "✅ $($interface.Name) → $($interface.Enum): All names match" -ForegroundColor Green
        }
    }

    # Report known exceptions
    $totalExceptions = 0
    foreach ($enumName in $knownExceptions.Keys) {
        $totalExceptions += $knownExceptions[$enumName].Count
    }

    if ($totalExceptions -gt 0) {
        Write-Host ""
        Write-Host "📝 Known Intentional Exceptions: $totalExceptions" -ForegroundColor Gray
        foreach ($enumName in $knownExceptions.Keys) {
            Write-Host "   $enumName`: " -NoNewline -ForegroundColor Gray
            Write-Host ($knownExceptions[$enumName] -join ", ") -ForegroundColor Gray
        }
        Write-Host "   (Documented in CORE-METHOD-RENAMING-SUMMARY.md)" -ForegroundColor Gray
    }

    if ($hasNamingIssues) {
        Write-Host ""
        Write-Host "⚠️  NAMING MISMATCHES DETECTED!" -ForegroundColor Red
        Write-Host ""
        Write-Host "Action Required:" -ForegroundColor Yellow
        Write-Host "  1. Review naming mismatches above" -ForegroundColor Yellow
        Write-Host "  2. Decide: Rename Core methods OR rename enum values" -ForegroundColor Yellow
        Write-Host "  3. Update all references (implementations, tools, tests, CLI)" -ForegroundColor Yellow
        Write-Host "  4. Run 'dotnet build' to verify" -ForegroundColor Yellow
        Write-Host "  5. If intentional, add to knownExceptions in audit script" -ForegroundColor Yellow
        Write-Host ""

        if ($FailOnGaps) {
            exit 1
        }
    } else {
        Write-Host ""
        Write-Host "✅ All naming consistent - enum values match Core method names!" -ForegroundColor Green
        Write-Host "   (Excluding $totalExceptions documented intentional exceptions)" -ForegroundColor Gray
    }
}

# Switch statement completeness check
Write-Host ""
Write-Host "🔀 Switch Statement Completeness Check" -ForegroundColor Cyan
Write-Host "=======================================" -ForegroundColor Cyan
Write-Host ""

# Function to extract handled enum values from switch statements
function Get-HandledEnumValues {
    param(
        [string]$ToolFilePath,
        [string]$EnumTypeName
    )

    if (-not (Test-Path $ToolFilePath)) {
        return @()
    }

    $content = Get-Content $ToolFilePath -Raw

    # Find switch statement on the enum type
    # Pattern: "action switch" or "return action switch" where action is the enum parameter
    # Match until we find the default case "_"
    $switchPattern = "(?s)return\s+action\s+switch\s*\{(.*?)\s+_\s*=>"

    if ($content -match $switchPattern) {
        $switchBody = $Matches[1]
        $handledValues = @()

        # Extract all case patterns: EnumType.Value =>
        $casePattern = "$EnumTypeName\.(\w+)\s*=>"
        $caseMatches = [regex]::Matches($switchBody, $casePattern)

        foreach ($match in $caseMatches) {
            $enumValue = $match.Groups[1].Value
            if ($handledValues -notcontains $enumValue) {
                $handledValues += $enumValue
            }
        }

        return $handledValues
    }

    return @()
}

# Check switch completeness for each tool
$toolsPath = Join-Path $rootDir "src\ExcelMcp.McpServer\Tools"
$switchIssues = @()
$hasSwitchIssues = $false

# Use the same discovered interfaces (already has Interface Name and EnumType)
$enumMappings = $interfaces

foreach ($mapping in $enumMappings) {
    $enumValues = Get-EnumValueNames -EnumName $mapping.Enum -ToolActionsPath $toolActionsPath

    # Dynamically find the tool file that uses this enum type as the first 'action' parameter
    # Look for: EnumType action, (as first parameter after method name)
    # This avoids false positives from references to other enum types in the same file
    $toolFiles = Get-ChildItem -Path $toolsPath -Filter "*.cs" | Where-Object {
        $content = Get-Content $_.FullName -Raw
        # Match the enum type as 'action' parameter in a method signature
        # Simplified pattern: look for the enum type followed by 'action' parameter
        # The method signature may span multiple lines and include 'partial' keyword
        $content -match "(?s)\b$($mapping.Enum)\s+action\s*,"
    }

    if ($toolFiles.Count -eq 0) {
        Write-Host "⚠️  No tool file found for $($mapping.Enum)" -ForegroundColor Yellow
        continue
    }

    if ($toolFiles.Count -gt 1) {
        # Multiple files use this enum - pick the one with matching name pattern
        # e.g., RangeAction -> RangeTool.cs or ExcelRangeTool.cs
        $enumBase = $mapping.Enum -replace 'Action$', ''
        $primaryTool = $toolFiles | Where-Object {
            $_.Name -match "$enumBase`Tool\.cs"
        } | Select-Object -First 1

        if (-not $primaryTool) {
            # Fallback to first file
            $primaryTool = $toolFiles[0]
        }
        $toolFile = $primaryTool
    } else {
        $toolFile = $toolFiles[0]
    }

    $handledValues = Get-HandledEnumValues -ToolFilePath $toolFile.FullName -EnumTypeName $mapping.Enum

    # Find unhandled enum values
    $unhandled = $enumValues | Where-Object { $handledValues -notcontains $_ }

    if ($unhandled.Count -gt 0) {
        $hasSwitchIssues = $true
        Write-Host "❌ $($toolFile.Name) ($($mapping.Enum)):" -ForegroundColor Red
        foreach ($value in $unhandled) {
            Write-Host "   Missing case: $($mapping.Enum).$value" -ForegroundColor Yellow
            $switchIssues += "Missing case: $($mapping.Enum).$value in $($toolFile.Name)"
        }
        Write-Host ""
    } else {
        Write-Host "✅ $($toolFile.Name): All $($enumValues.Count) enum values handled" -ForegroundColor Green
    }
}

if ($hasSwitchIssues) {
    Write-Host ""
    Write-Host "⚠️  UNHANDLED ENUM VALUES DETECTED!" -ForegroundColor Red
    Write-Host ""
    Write-Host "Action Required:" -ForegroundColor Yellow
    Write-Host "  1. Review missing case statements above" -ForegroundColor Yellow
    Write-Host "  2. Add missing cases to switch statements in tool files" -ForegroundColor Yellow
    Write-Host "  3. Implement the corresponding private methods" -ForegroundColor Yellow
    Write-Host "  4. Run 'dotnet build' to verify compilation" -ForegroundColor Yellow
    Write-Host "  5. Test the new actions work correctly" -ForegroundColor Yellow
    Write-Host ""
    Write-Host "Example fix for PowerQueryAction.LoadTo:" -ForegroundColor Gray
    Write-Host "  PowerQueryAction.LoadTo => await LoadToPowerQueryAsync(...)" -ForegroundColor Gray
    Write-Host ""

    if ($FailOnGaps) {
        exit 1
    }
} else {
    Write-Host ""
    Write-Host "✅ All switch statements complete - every enum value is handled!" -ForegroundColor Green
}

exit 0
