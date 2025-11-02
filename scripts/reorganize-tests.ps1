# Test Structure Reorganization Script
# This script reorganizes the test structure to match Core commands

Write-Host "╔════════════════════════════════════════════════════════════╗" -ForegroundColor Cyan
Write-Host "║  Test Structure Reorganization - Phase-by-Phase           ║" -ForegroundColor Cyan  
Write-Host "╚════════════════════════════════════════════════════════════╝`n" -ForegroundColor Cyan

$ErrorActionPreference = "Stop"
$baseDir = "tests\ExcelMcp.Core.Tests\Integration\Commands"

# Phase 1: Rename Directories
Write-Host "PHASE 1: Rename Directories to Match Core" -ForegroundColor Yellow
Write-Host "=" * 60 `n

$directoryRenames = @{
    "$baseDir\Parameter" = "$baseDir\NamedRange"
    "$baseDir\Script" = "$baseDir\Vba"  
    "$baseDir\FileOperations" = "$baseDir\File"
}

foreach ($rename in $directoryRenames.GetEnumerator()) {
    $old = $rename.Key
    $new = $rename.Value
    
    if (Test-Path $old) {
        Write-Host "  Renaming: $(Split-Path $old -Leaf) → $(Split-Path $new -Leaf)" -ForegroundColor White
        git mv $old $new
        Write-Host "    ✅ Complete" -ForegroundColor Green
    } else {
        Write-Host "    ⚠️  Directory not found: $old" -ForegroundColor Yellow
    }
}

# Merge VbaTrust into Vba
if (Test-Path "$baseDir\VbaTrust") {
    Write-Host "`n  Merging VbaTrust into Vba..." -ForegroundColor White
    Get-ChildItem "$baseDir\VbaTrust" -File | ForEach-Object {
        $newName = $_.Name -replace "VbaTrustDetectionTests", "VbaCommandsTests.Trust"
        git mv $_.FullName "$baseDir\Vba\$newName"
        Write-Host "    Moved: $($_.Name) → Vba/$newName" -ForegroundColor Gray
    }
    Remove-Item "$baseDir\VbaTrust" -Recurse -Force
    Write-Host "    ✅ VbaTrust merged into Vba" -ForegroundColor Green
}

Write-Host "`n" + ("=" * 60)
Write-Host "Phase 1 Complete!`n" -ForegroundColor Green

# Phase 2: Fix Namespaces  
Write-Host "PHASE 2: Fix Namespace Inconsistencies" -ForegroundColor Yellow
Write-Host "=" * 60 `n

function Update-Namespace {
    param([string]$filePath, [string]$oldNamespace, [string]$newNamespace)
    
    if (Test-Path $filePath) {
        $content = Get-Content $filePath -Raw
        $updated = $content -replace [regex]::Escape($oldNamespace), $newNamespace
        
        if ($content -ne $updated) {
            Set-Content $filePath -Value $updated -NoNewline
            Write-Host "    Updated: $(Split-Path $filePath -Leaf)" -ForegroundColor Gray
            return $true
        }
    }
    return $false
}

# Fix Range namespace (Integration.Range → Commands.Range)
Write-Host "  Fixing Range namespaces..." -ForegroundColor White
$rangeFiles = Get-ChildItem "$baseDir\Range" -Filter "*.cs"
$rangeCount = 0
foreach ($file in $rangeFiles) {
    if (Update-Namespace $file.FullName `
        "Sbroenne.ExcelMcp.Core.Tests.Integration.Range" `
        "Sbroenne.ExcelMcp.Core.Tests.Commands.Range") {
        $rangeCount++
    }
}
Write-Host "    ✅ Updated $rangeCount Range files" -ForegroundColor Green

# Fix PowerQuery regression test namespace
Write-Host "  Fixing PowerQuery regression test..." -ForegroundColor White
$pqFile = "$baseDir\PowerQuery\PowerQuerySuccessErrorRegressionTests.cs"
if (Update-Namespace $pqFile `
    "Sbroenne.ExcelMcp.Core.Tests.Integration.Commands.PowerQuery" `
    "Sbroenne.ExcelMcp.Core.Tests.Commands.PowerQuery") {
    Write-Host "    ✅ Updated PowerQuery regression test" -ForegroundColor Green
}

# Update namespaces after directory renames
Write-Host "  Updating namespaces for renamed directories..." -ForegroundColor White
$namespaceUpdates = @{
    "NamedRange" = @{
        Old = "Sbroenne.ExcelMcp.Core.Tests.Commands.Parameter"
        New = "Sbroenne.ExcelMcp.Core.Tests.Commands.NamedRange"
    }
    "Vba" = @{
        Old = "Sbroenne.ExcelMcp.Core.Tests.Commands.Script|Sbroenne.ExcelMcp.Core.Tests.Commands.VbaTrust"
        New = "Sbroenne.ExcelMcp.Core.Tests.Commands.Vba"
    }
    "File" = @{
        Old = "Sbroenne.ExcelMcp.Core.Tests.Commands.FileOperations"
        New = "Sbroenne.ExcelMcp.Core.Tests.Commands.File"
    }
}

foreach ($dir in $namespaceUpdates.Keys) {
    $dirPath = "$baseDir\$dir"
    if (Test-Path $dirPath) {
        $files = Get-ChildItem $dirPath -Filter "*.cs"
        foreach ($file in $files) {
            $content = Get-Content $file.FullName -Raw
            $oldPatterns = $namespaceUpdates[$dir].Old -split '\|'
            $updated = $content
            
            foreach ($oldPattern in $oldPatterns) {
                $updated = $updated -replace [regex]::Escape($oldPattern), $namespaceUpdates[$dir].New
            }
            
            if ($content -ne $updated) {
                Set-Content $file.FullName -Value $updated -NoNewline
                Write-Host "    Updated: $dir\$($file.Name)" -ForegroundColor Gray
            }
        }
    }
}

Write-Host "`n" + ("=" * 60)
Write-Host "Phase 2 Complete!`n" -ForegroundColor Green

# Phase 3: Rename Test Classes
Write-Host "PHASE 3: Rename Test Classes" -ForegroundColor Yellow
Write-Host "=" * 60 `n

function Update-ClassName {
    param([string]$filePath, [string]$oldClass, [string]$newClass)
    
    if (Test-Path $filePath) {
        $content = Get-Content $filePath -Raw
        
        # Update class declaration
        $updated = $content -replace "public\s+(partial\s+)?class\s+$oldClass\b", "public `$1class $newClass"
        
        # Update constructor
        $updated = $updated -replace "public\s+$oldClass\s*\(", "public $newClass("
        
        if ($content -ne $updated) {
            Set-Content $filePath -Value $updated -NoNewline
            return $true
        }
    }
    return $false
}

# Rename ParameterCommandsTests → NamedRangeCommandsTests
if (Test-Path "$baseDir\NamedRange") {
    Write-Host "  Renaming ParameterCommandsTests → NamedRangeCommandsTests..." -ForegroundColor White
    $files = Get-ChildItem "$baseDir\NamedRange" -Filter "*.cs"
    $count = 0
    foreach ($file in $files) {
        if (Update-ClassName $file.FullName "ParameterCommandsTests" "NamedRangeCommandsTests") {
            $count++
        }
    }
    Write-Host "    ✅ Updated $count NamedRange files" -ForegroundColor Green
}

# Rename ScriptCommandsTests → VbaCommandsTests
if (Test-Path "$baseDir\Vba") {
    Write-Host "  Renaming ScriptCommandsTests/VbaTrustDetectionTests → VbaCommandsTests..." -ForegroundColor White
    $files = Get-ChildItem "$baseDir\Vba" -Filter "*.cs"
    $count = 0
    foreach ($file in $files) {
        $updated = $false
        if (Update-ClassName $file.FullName "ScriptCommandsTests" "VbaCommandsTests") {
            $updated = $true
        }
        if (Update-ClassName $file.FullName "VbaTrustDetectionTests" "VbaCommandsTests") {
            $updated = $true
        }
        if ($updated) { $count++ }
    }
    Write-Host "    ✅ Updated $count Vba files" -ForegroundColor Green
}

Write-Host "`n" + ("=" * 60)
Write-Host "Phase 3 Complete!`n" -ForegroundColor Green

# Phase 4: Rename Sheet Test Files
Write-Host "PHASE 4: Consolidate Sheet Tests as Partials" -ForegroundColor Yellow
Write-Host "=" * 60 `n

if (Test-Path "$baseDir\Sheet") {
    Write-Host "  Converting Sheet tests to partial classes..." -ForegroundColor White
    
    # Rename SheetTabColorTests.cs → SheetCommandsTests.TabColor.cs
    $tabColorFile = "$baseDir\Sheet\SheetTabColorTests.cs"
    if (Test-Path $tabColorFile) {
        git mv $tabColorFile "$baseDir\Sheet\SheetCommandsTests.TabColor.cs"
        Write-Host "    Renamed: SheetTabColorTests.cs → SheetCommandsTests.TabColor.cs" -ForegroundColor Gray
        
        # Update class to partial
        $file = "$baseDir\Sheet\SheetCommandsTests.TabColor.cs"
        Update-ClassName $file "SheetTabColorTests" "SheetCommandsTests"
        $content = Get-Content $file -Raw
        $content = $content -replace "public\s+class\s+SheetCommandsTests", "public partial class SheetCommandsTests"
        Set-Content $file -Value $content -NoNewline
    }
    
    # Rename SheetVisibilityTests.cs → SheetCommandsTests.Visibility.cs
    $visFile = "$baseDir\Sheet\SheetVisibilityTests.cs"
    if (Test-Path $visFile) {
        git mv $visFile "$baseDir\Sheet\SheetCommandsTests.Visibility.cs"
        Write-Host "    Renamed: SheetVisibilityTests.cs → SheetCommandsTests.Visibility.cs" -ForegroundColor Gray
        
        # Update class to partial
        $file = "$baseDir\Sheet\SheetCommandsTests.Visibility.cs"
        Update-ClassName $file "SheetVisibilityTests" "SheetCommandsTests"
        $content = Get-Content $file -Raw
        $content = $content -replace "public\s+class\s+SheetCommandsTests", "public partial class SheetCommandsTests"
        Set-Content $file -Value $content -NoNewline
    }
    
    Write-Host "    ✅ Sheet tests consolidated" -ForegroundColor Green
}

Write-Host "`n" + ("=" * 60)
Write-Host "Phase 4 Complete!`n" -ForegroundColor Green

# Summary
Write-Host "`n╔════════════════════════════════════════════════════════════╗" -ForegroundColor Green
Write-Host "║  Reorganization Complete!                                  ║" -ForegroundColor Green
Write-Host "╚════════════════════════════════════════════════════════════╝`n" -ForegroundColor Green

Write-Host "✅ Directory renames complete" -ForegroundColor Green
Write-Host "✅ Namespace updates complete" -ForegroundColor Green  
Write-Host "✅ Class renames complete" -ForegroundColor Green
Write-Host "✅ File renames complete`n" -ForegroundColor Green

Write-Host "Next Steps:" -ForegroundColor Yellow
Write-Host "  1. Build the project: dotnet build tests\ExcelMcp.Core.Tests" -ForegroundColor White
Write-Host "  2. Run tests: dotnet test tests\ExcelMcp.Core.Tests" -ForegroundColor White
Write-Host "  3. Review changes: git status" -ForegroundColor White
Write-Host "  4. Commit: git commit -m 'test: reorganize test structure to match Core commands'" -ForegroundColor White
Write-Host ""
