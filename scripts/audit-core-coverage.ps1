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

Write-Host "Core Commands Coverage Audit" -ForegroundColor Cyan
Write-Host "=================================" -ForegroundColor Cyan
Write-Host ""

# Function to count unique async method names in Core interface files (handles overloads)
function Get-CoreMethodMatches {
    param([string]$InterfacePath, [string]$InterfaceName)

    if (-not (Test-Path $InterfacePath)) {
        return @()
    }

    $content = Get-Content $InterfacePath -Raw
    # Comments can contain example object literals with braces, so remove them
    # before isolating the interface body.
    $content = [regex]::Replace($content, '(?s)/\*.*?\*/', '')
    $content = [regex]::Replace($content, '(?m)//.*$', '')
    $interfacePattern = "public\s+interface\s+$([regex]::Escape($InterfaceName))\b[^\{]*\{(?<body>[^}]*)\}"
    $interfaceMatch = [regex]::Match($content, $interfacePattern)
    if (-not $interfaceMatch.Success) {
        throw "Could not parse interface '$InterfaceName' from $InterfacePath"
    }
    $content = $interfaceMatch.Groups['body'].Value

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

    $methodNames = Get-CoreMethodMatches -InterfacePath $InterfacePath -InterfaceName $InterfaceName
    return $methodNames.Count
}

# Locate the source file that declares an action enum.
function Find-EnumSourceFile {
    param([string]$EnumName, [string[]]$ActionEnumPaths)

    $enumPattern = "public\s+enum\s+$([regex]::Escape($EnumName))\b"
    $matches = @($ActionEnumPaths | Where-Object {
        (Get-Content $_ -Raw) -match $enumPattern
    })

    if ($matches.Count -gt 1) {
        throw "Action enum '$EnumName' is declared more than once: $($matches -join ', ')"
    }

    return $matches | Select-Object -First 1
}

# Function to count enum values for a specific interface (handles cross-interface enum splits)
function Count-EnumValuesForInterface {
    param(
        [string]$EnumName,
        [string]$InterfaceName,
        [string[]]$ActionEnumPaths
    )

    $enumValues = @(Get-EnumValueNames -EnumName $EnumName -ActionEnumPaths $ActionEnumPaths)
    if ($Script:enumCoverageExceptions.ContainsKey($EnumName)) {
        $exceptions = $Script:enumCoverageExceptions[$EnumName]
        $enumValues = @($enumValues | Where-Object { $exceptions -notcontains $_ })
    }

    return $enumValues.Count
}

# Function to extract unique method names from Core interface (without "Async" suffix, handles overloads)
function Get-CoreMethodNames {
    param([string]$InterfacePath, [string]$InterfaceName)

    return Get-CoreMethodMatches -InterfacePath $InterfacePath -InterfaceName $InterfaceName
}

# Function to extract enum value names
function Get-EnumValueNames {
    param([string]$EnumName, [string[]]$ActionEnumPaths)

    $enumSource = Find-EnumSourceFile -EnumName $EnumName -ActionEnumPaths $ActionEnumPaths
    if (-not $enumSource) {
        return @()
    }

    $content = Get-Content $enumSource -Raw
    $enumPattern = "public\s+enum\s+$([regex]::Escape($EnumName))\s*(?::[^\{]+)?\{(?<body>[^}]*)\}"
    $enumMatch = [regex]::Match($content, $enumPattern)
    if (-not $enumMatch.Success) {
        return @()
    }

    $identifierPattern = '(?m)^\s*(?<name>[A-Za-z_][A-Za-z0-9_]*)\s*(?:=\s*[^,\r\n]+)?\s*,?\s*(?://.*)?$'
    return @([regex]::Matches($enumMatch.Groups['body'].Value, $identifierPattern) |
        ForEach-Object { $_.Groups['name'].Value })
}

# Function to check naming consistency
function Check-NamingConsistency {
    param(
        [string]$InterfaceName,
        [string]$InterfacePath,
        [string]$EnumName,
        [string[]]$ActionEnumPaths
    )

    $methodNames = Get-CoreMethodNames -InterfacePath $InterfacePath -InterfaceName $InterfaceName
    $enumValues = Get-EnumValueNames -EnumName $EnumName -ActionEnumPaths $ActionEnumPaths

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
    param([string[]]$ActionEnumPaths)

    $enumTypes = foreach ($path in $ActionEnumPaths) {
        $content = Get-Content $path -Raw
        [regex]::Matches($content, 'public\s+enum\s+(\w+Action)\b') |
            ForEach-Object { $_.Groups[1].Value }
    }

    return @($enumTypes | Sort-Object -Unique)
}

# Discover interface files dynamically
function Find-InterfaceForEnum {
    param(
        [string]$EnumType,
        [string]$CommandsPath
    )

    # Map enum type to expected interface name.
    # Most generated enums follow {Name}Action -> I{Name}Commands.

    $enumToInterface = @{
        # Known naming exceptions.
        "CalculationAction" = "ICalculationModeCommands"
        "ConditionalFormatAction" = "IConditionalFormattingCommands"
    }

    if ($enumToInterface.ContainsKey($EnumType)) {
        $interfaceName = $enumToInterface[$EnumType]
    } else {
        # Standard pattern: {Name}Action -> I{Name}Commands
        $baseName = $EnumType -replace 'Action$', ''
        $interfaceName = "I${baseName}Commands"
    }

    # Prefer the conventional filename, then fall back to declaration search
    # for interfaces that share a file with their implementation.
    $interfaceFiles = Get-ChildItem -Path $CommandsPath -Recurse -Filter "$interfaceName.cs"
    if ($interfaceFiles.Count -eq 0) {
        $interfacePattern = "public\s+interface\s+$([regex]::Escape($interfaceName))\b"
        $interfaceFiles = @(Get-ChildItem -Path $CommandsPath -Recurse -Filter "*.cs" | Where-Object {
            (Get-Content $_.FullName -Raw) -match $interfacePattern
        })
    }

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

$toolActionsPath = Join-Path $rootDir "src\ExcelMcp.Core\Models\Actions\ToolActions.cs"
if (-not (Test-Path $toolActionsPath)) {
    Write-Error "Manual action enum source not found: $toolActionsPath"
    exit 1
}

$generatedActionsRoot = Join-Path $rootDir "src\ExcelMcp.Core\obj\GeneratedFiles"
if (-not (Test-Path $generatedActionsRoot)) {
    Write-Error "Generated action enums are absent. Run a build that emits compiler-generated files first."
    exit 1
}

$generatedActionFiles = @(Get-ChildItem -Path $generatedActionsRoot -Recurse -Filter "ServiceRegistry.*.g.cs" |
    Where-Object { (Get-Content $_.FullName -Raw) -match 'public\s+enum\s+\w+Action\b' })
if ($generatedActionFiles.Count -eq 0) {
    Write-Error "No emitted ServiceRegistry.*.g.cs action enums were found under $generatedActionsRoot. Run a build first."
    exit 1
}

$actionEnumPaths = @($toolActionsPath) + @($generatedActionFiles.FullName)

# FileAction contains ten service/session actions that intentionally do not map
# to IFileCommands. Test is the one FileAction backed by the Core interface.
$Script:enumCoverageExceptions = @{
    "FileAction" = @(
        "CloseWorkbook", "Open", "Close", "List", "Create",
        "Preflight", "ConfigureSafety", "Journal", "Recoveries", "Recover"
    )
}

# Dynamically discover all interfaces to check
$commandsPath = Join-Path $rootDir "src\ExcelMcp.Core\Commands"
$enumTypes = @(Get-AllEnumTypes -ActionEnumPaths $actionEnumPaths)
if ($enumTypes.Count -le 1) {
    Write-Error "Action enum discovery was vacuous: found only $($enumTypes.Count) enum(s)."
    exit 1
}
Write-Host "Loaded $($enumTypes.Count) action enums ($($generatedActionFiles.Count) generated, 1 manual)." -ForegroundColor DarkGray

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
        $enumCount = Count-EnumValuesForInterface -EnumName $enumName -InterfaceName $interfaceGroup.Name -ActionEnumPaths $actionEnumPaths
        $totalEnumValuesForInterface += $enumCount
        $enumNames += "$enumName($enumCount)"
    }

    $totalCoreMethods += $coreMethods
    $totalEnumValues += $totalEnumValuesForInterface

    $statusText = "OK"

    if ($totalEnumValuesForInterface -lt $coreMethods) {
        $statusText = "GAP"
        $hasGaps = $true
    } elseif ($totalEnumValuesForInterface -gt $coreMethods) {
        $statusText = "EXTRA"
    }

    $result = [PSCustomObject]@{
        Interface = $interfaceGroup.Name
        CoreMethods = $coreMethods
        EnumValues = $totalEnumValuesForInterface
        Enums = ($interfaceGroup.Enums -join ", ")
        Gap = $coreMethods - $totalEnumValuesForInterface
        Status = $statusText
    }

    $results += $result

    if ($Verbose) {
        Write-Host "Checking $($interfaceGroup.Name)..." -ForegroundColor Gray
        Write-Host "  Core Methods: $coreMethods" -ForegroundColor Gray
        Write-Host "  Enum Values: $totalEnumValuesForInterface (from: $($enumNames -join ', '))" -ForegroundColor Gray
        Write-Host "  Status: $statusText" -ForegroundColor $(if ($statusText -eq "OK") { "Green" } elseif ($statusText -eq "GAP") { "Red" } else { "Yellow" })
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
    Write-Host "Coverage:           100% " -ForegroundColor Green
} else {
    $coverage = [math]::Round(($totalEnumValues / $totalCoreMethods) * 100, 1)
    Write-Host "Coverage:           $coverage%" -ForegroundColor $(if ($coverage -ge 95) { "Yellow" } else { "Red" })
}

# Gaps detection
if ($hasGaps) {
    Write-Host ""
    Write-Host "GAPS DETECTED!" -ForegroundColor Red
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
    Write-Host "No gaps detected - 100% coverage maintained!" -ForegroundColor Green
}

# Extra enum values warning
$extraEnums = $results | Where-Object { $_.Gap -lt 0 }
if ($extraEnums.Count -gt 0) {
    Write-Host ""
    Write-Host "Note: Some enums have more values than Core methods" -ForegroundColor Yellow
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
    Write-Host "Naming Consistency Check" -ForegroundColor Cyan
    Write-Host "===========================" -ForegroundColor Cyan
    Write-Host ""

    $knownExceptions = $Script:enumCoverageExceptions

    $hasNamingIssues = $false

    foreach ($interface in $interfaces) {
        $mismatches = Check-NamingConsistency `
            -InterfaceName $interface.Name `
            -InterfacePath $interface.Path `
            -EnumName $interface.Enum `
            -ActionEnumPaths $actionEnumPaths

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
            Write-Host "$($interface.Name) -> $($interface.Enum):" -ForegroundColor Red
            foreach ($mismatch in $mismatches) {
                Write-Host "   $mismatch" -ForegroundColor Yellow
            }
            Write-Host ""
        } else {
            Write-Host "$($interface.Name) -> $($interface.Enum): All names match" -ForegroundColor Green
        }
    }

    # Report known exceptions
    $totalExceptions = 0
    foreach ($enumName in $knownExceptions.Keys) {
        $totalExceptions += $knownExceptions[$enumName].Count
    }

    if ($totalExceptions -gt 0) {
        Write-Host ""
        Write-Host "Known Intentional Exceptions: $totalExceptions" -ForegroundColor Gray
        foreach ($enumName in $knownExceptions.Keys) {
            Write-Host "   $enumName`: " -NoNewline -ForegroundColor Gray
            Write-Host ($knownExceptions[$enumName] -join ", ") -ForegroundColor Gray
        }
        Write-Host "   (Manual service/session actions without Core interface methods)" -ForegroundColor Gray
    }

    if ($hasNamingIssues) {
        Write-Host ""
        Write-Host "NAMING MISMATCHES DETECTED!" -ForegroundColor Red
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
        Write-Host "All naming consistent - enum values match Core method names!" -ForegroundColor Green
        Write-Host "   (Excluding $totalExceptions documented intentional exceptions)" -ForegroundColor Gray
    }
}

# Switch statement completeness check
Write-Host ""
Write-Host "Switch Statement Completeness Check" -ForegroundColor Cyan
Write-Host "=======================================" -ForegroundColor Cyan
Write-Host ""

# Function to extract handled enum values from switch statements. Hand-written
# tools use expression arms (Enum.Value =>); generated dispatchers use
# statement cases (case Enum.Value:).
function Get-HandledEnumValues {
    param(
        [string]$ToolFilePath,
        [string]$EnumTypeName
    )

    if (-not (Test-Path $ToolFilePath)) {
        return @()
    }

    $content = Get-Content $ToolFilePath -Raw

    $escapedEnum = [regex]::Escape($EnumTypeName)
    $casePattern = "(?:case\s+)?$escapedEnum\.(?<name>\w+)\s*(?::|=>)"
    $handledValues = @([regex]::Matches($content, $casePattern) |
        ForEach-Object { $_.Groups['name'].Value } |
        Sort-Object -Unique)

    if ($handledValues.Count -gt 0) {
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
    $enumValues = @(Get-EnumValueNames -EnumName $mapping.Enum -ActionEnumPaths $actionEnumPaths)

    if ($mapping.Enum -eq "FileAction") {
        $toolFile = Get-Item (Join-Path $toolsPath "ExcelFileTool.cs")
    } else {
        $enumSource = Find-EnumSourceFile -EnumName $mapping.Enum -ActionEnumPaths $actionEnumPaths
        $dispatchPath = $enumSource -replace '\.g\.cs$', '.Dispatch.g.cs'
        if (-not (Test-Path $dispatchPath)) {
            $hasSwitchIssues = $true
            Write-Host "Generated dispatcher missing for $($mapping.Enum): $dispatchPath" -ForegroundColor Red
            continue
        }
        $toolFile = Get-Item $dispatchPath
    }

    $handledValues = Get-HandledEnumValues -ToolFilePath $toolFile.FullName -EnumTypeName $mapping.Enum

    # Find unhandled enum values
    $unhandled = $enumValues | Where-Object { $handledValues -notcontains $_ }

    if ($unhandled.Count -gt 0) {
        $hasSwitchIssues = $true
        Write-Host "$($toolFile.Name) ($($mapping.Enum)):" -ForegroundColor Red
        foreach ($value in $unhandled) {
            Write-Host "   Missing case: $($mapping.Enum).$value" -ForegroundColor Yellow
            $switchIssues += "Missing case: $($mapping.Enum).$value in $($toolFile.Name)"
        }
        Write-Host ""
    } else {
        Write-Host "$($toolFile.Name): All $($enumValues.Count) enum values handled" -ForegroundColor Green
    }
}

if ($hasSwitchIssues) {
    Write-Host ""
    Write-Host "UNHANDLED ENUM VALUES DETECTED!" -ForegroundColor Red
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
    Write-Host "All switch statements complete - every enum value is handled!" -ForegroundColor Green
}

exit 0
