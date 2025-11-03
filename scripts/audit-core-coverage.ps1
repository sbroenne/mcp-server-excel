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

# Function to count async methods in Core interface files
function Count-CoreMethods {
    param([string]$InterfacePath, [string]$InterfaceName)

    if (-not (Test-Path $InterfacePath)) {
        Write-Warning "Interface file not found: $InterfacePath"
        return 0
    }

    $content = Get-Content $InterfacePath -Raw
    # Count lines like: Task<Something> MethodAsync(
    $matches = [regex]::Matches($content, 'Task<[^>]+>\s+\w+Async\s*\(')
    return $matches.Count
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

# Function to extract method names from Core interface (without "Async" suffix)
function Get-CoreMethodNames {
    param([string]$InterfacePath)

    if (-not (Test-Path $InterfacePath)) {
        return @()
    }

    $content = Get-Content $InterfacePath -Raw
    $matches = [regex]::Matches($content, 'Task<[^>]+>\s+(\w+)Async\s*\(')
    $methodNames = @()
    foreach ($match in $matches) {
        $methodNames += $match.Groups[1].Value
    }
    return $methodNames
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

# Define interfaces to check
$interfaces = @(
    @{
        Name = "IPowerQueryCommands"
        Path = "$rootDir/src/ExcelMcp.Core/Commands/PowerQuery/IPowerQueryCommands.cs"
        Enum = "PowerQueryAction"
    },
    @{
        Name = "ISheetCommands"
        Path = "$rootDir/src/ExcelMcp.Core/Commands/Sheet/ISheetCommands.cs"
        Enum = "WorksheetAction"
    },
    @{
        Name = "IRangeCommands"
        Path = "$rootDir/src/ExcelMcp.Core/Commands/Range/IRangeCommands.cs"
        Enum = "RangeAction"
    },
    @{
        Name = "ITableCommands"
        Path = "$rootDir/src/ExcelMcp.Core/Commands/Table/ITableCommands.cs"
        Enum = "TableAction"
    },
    @{
        Name = "IConnectionCommands"
        Path = "$rootDir/src/ExcelMcp.Core/Commands/Connection/IConnectionCommands.cs"
        Enum = "ConnectionAction"
    },
    @{
        Name = "IDataModelCommands"
        Path = "$rootDir/src/ExcelMcp.Core/Commands/DataModel/IDataModelCommands.cs"
        Enum = "DataModelAction"
    },
    @{
        Name = "IPivotTableCommands"
        Path = "$rootDir/src/ExcelMcp.Core/Commands/PivotTable/IPivotTableCommands.cs"
        Enum = "PivotTableAction"
    },
    @{
        Name = "INamedRangeCommands"
        Path = "$rootDir/src/ExcelMcp.Core/Commands/NamedRange/INamedRangeCommands.cs"
        Enum = "NamedRangeAction"
    },
    @{
        Name = "IVbaCommands"
        Path = "$rootDir/src/ExcelMcp.Core/Commands/Vba/IVbaCommands.cs"
        Enum = "VbaAction"
    },
    @{
        Name = "IFileCommands"
        Path = "$rootDir/src/ExcelMcp.Core/Commands/IFileCommands.cs"
        Enum = "FileAction"
    }
)

$toolActionsPath = "$rootDir/src/ExcelMcp.McpServer/Models/ToolActions.cs"

# Track results
$results = @()
$totalCoreMethods = 0
$totalEnumValues = 0
$hasGaps = $false

# Audit each interface
foreach ($interface in $interfaces) {
    $coreMethods = Count-CoreMethods -InterfacePath $interface.Path -InterfaceName $interface.Name
    $enumValues = Count-EnumValues -EnumName $interface.Enum -ToolActionsPath $toolActionsPath

    $totalCoreMethods += $coreMethods
    $totalEnumValues += $enumValues

    $status = "✅"
    $statusText = "OK"

    if ($enumValues -lt $coreMethods) {
        $status = "❌"
        $statusText = "GAP"
        $hasGaps = $true
    } elseif ($enumValues -gt $coreMethods) {
        $status = "⚠️"
        $statusText = "EXTRA"
    }

    $result = [PSCustomObject]@{
        Interface = $interface.Name
        CoreMethods = $coreMethods
        EnumValues = $enumValues
        Gap = $coreMethods - $enumValues
        Status = $status
        StatusText = $statusText
    }

    $results += $result

    if ($Verbose) {
        Write-Host "Checking $($interface.Name)..." -ForegroundColor Gray
        Write-Host "  Core Methods: $coreMethods" -ForegroundColor Gray
        Write-Host "  Enum Values: $enumValues" -ForegroundColor Gray
        Write-Host "  Status: $status $statusText" -ForegroundColor $(if ($statusText -eq "OK") { "Green" } elseif ($statusText -eq "GAP") { "Red" } else { "Yellow" })
        Write-Host ""
    }
}

# Display results table
Write-Host ""
Write-Host "Audit Results:" -ForegroundColor Cyan
Write-Host ""
$results | Format-Table -Property Interface, CoreMethods, EnumValues, Gap, Status -AutoSize

# Summary
Write-Host ""
Write-Host "Summary:" -ForegroundColor Cyan
Write-Host "--------" -ForegroundColor Cyan
Write-Host "Total Core Methods: $totalCoreMethods" -ForegroundColor White
Write-Host "Total Enum Values:  $totalEnumValues" -ForegroundColor White

if ($totalEnumValues -eq $totalCoreMethods) {
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
    
    $namingIssues = @()
    $hasNamingIssues = $false
    
    foreach ($interface in $interfaces) {
        $mismatches = Check-NamingConsistency `
            -InterfaceName $interface.Name `
            -InterfacePath $interface.Path `
            -EnumName $interface.Enum `
            -ToolActionsPath $toolActionsPath
        
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
    
    if ($hasNamingIssues) {
        Write-Host ""
        Write-Host "⚠️  NAMING MISMATCHES DETECTED!" -ForegroundColor Red
        Write-Host ""
        Write-Host "Action Required:" -ForegroundColor Yellow
        Write-Host "  1. Review naming mismatches above" -ForegroundColor Yellow
        Write-Host "  2. Decide: Rename Core methods OR rename enum values" -ForegroundColor Yellow
        Write-Host "  3. Update all references (implementations, tools, tests, CLI)" -ForegroundColor Yellow
        Write-Host "  4. Run 'dotnet build' to verify" -ForegroundColor Yellow
        Write-Host ""
        
        if ($FailOnGaps) {
            exit 1
        }
    } else {
        Write-Host ""
        Write-Host "✅ All naming consistent - enum values match Core method names!" -ForegroundColor Green
    }
}

exit 0
