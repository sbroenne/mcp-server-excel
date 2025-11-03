#!/usr/bin/env pwsh
<#
.SYNOPSIS
    Automatically fixes Success = true with ErrorMessage violations (Rule 0)

.DESCRIPTION
    Finds all violations where:
    - result.Success = true (set optimistically)
    - catch block sets result.ErrorMessage
    - but FORGETS to set result.Success = false
    
    Fixes by adding: result.Success = false; at the start of each catch block.
    Creates backup of all modified files.

.PARAMETER DryRun
    Show what would be fixed without making changes

.PARAMETER NoBackup
    Skip creating backup files

.EXAMPLE
    .\fix-success-flags.ps1
    Fixes all violations with backup
    
.EXAMPLE
    .\fix-success-flags.ps1 -DryRun
    Shows what would be fixed without making changes

.NOTES
    Part of Rule 0 enforcement. See CRITICAL-RULES.md Rule 0.
#>

param(
    [switch]$DryRun,
    [switch]$NoBackup
)

$ErrorActionPreference = "Stop"
$rootDir = Split-Path -Parent $PSScriptRoot

Write-Host "🔧 Fixing Success = true with ErrorMessage violations (Rule 0)..." -ForegroundColor Cyan
Write-Host ""

if ($DryRun) {
    Write-Host "⚠️  DRY RUN MODE - No files will be modified" -ForegroundColor Yellow
    Write-Host ""
}

# Find all violations
$violations = @()

Get-ChildItem -Path "$rootDir\src\ExcelMcp.Core\Commands" -Filter "*.cs" -Recurse | ForEach-Object {
    $lines = Get-Content $_.FullName
    
    for ($i = 0; $i -lt $lines.Count; $i++) {
        # Look for Success = true
        if ($lines[$i] -match '\.Success\s*=\s*true') {
            $successLine = $i
            
            # Check next 30 lines for ErrorMessage = "something" (not null/empty)
            for ($j = $i + 1; $j -lt [Math]::Min($i + 30, $lines.Count); $j++) {
                # Skip if we hit another Success assignment
                if ($lines[$j] -match '\.Success\s*=') {
                    break
                }
                
                # Found ErrorMessage being set to non-null value
                if ($lines[$j] -match '\.ErrorMessage\s*=\s*["\$]' -and 
                    $lines[$j] -notmatch '= null' -and 
                    $lines[$j] -notmatch '= string.Empty' -and
                    $lines[$j] -notmatch '= ""') {
                    
                    $errorLine = $j
                    
                    # Check if Success = false exists between these lines
                    $hasSuccessFalse = $false
                    for ($k = $i + 1; $k -lt $j; $k++) {
                        if ($lines[$k] -match '\.Success\s*=\s*false') {
                            $hasSuccessFalse = $true
                            break
                        }
                    }
                    
                    if (-not $hasSuccessFalse) {
                        # Find the catch block line (search backwards from error line)
                        $catchLine = -1
                        for ($c = $errorLine; $c -ge 0; $c--) {
                            if ($lines[$c] -match '^\s*catch\s*[\({]') {
                                $catchLine = $c
                                break
                            }
                        }
                        
                        $violations += [PSCustomObject]@{
                            File = $_.FullName
                            FileName = $_.Name
                            SuccessLine = $successLine
                            ErrorLine = $errorLine
                            CatchLine = $catchLine
                            Lines = $lines
                        }
                    }
                    break
                }
            }
        }
    }
}

if ($violations.Count -eq 0) {
    Write-Host "✅ No violations found - all Success flags match reality!" -ForegroundColor Green
    exit 0
}

Write-Host "Found $($violations.Count) violations to fix:" -ForegroundColor Yellow
Write-Host ""

# Group by file
$fileGroups = $violations | Group-Object -Property File

foreach ($fileGroup in $fileGroups) {
    $filePath = $fileGroup.Name
    $fileName = [System.IO.Path]::GetFileName($filePath)
    $fileViolations = $fileGroup.Group
    
    Write-Host "📝 $fileName - $($fileViolations.Count) violation(s)" -ForegroundColor Cyan
    
    if (-not $DryRun) {
        # Create backup
        if (-not $NoBackup) {
            $backupPath = "$filePath.bak"
            Copy-Item $filePath $backupPath -Force
            Write-Host "   Backup created: $fileName.bak" -ForegroundColor Gray
        }
        
        # Read file content
        $lines = Get-Content $filePath
        
        # Sort violations by line number (descending) to avoid line number shifts
        $sortedViolations = $fileViolations | Sort-Object -Property CatchLine -Descending
        
        foreach ($violation in $sortedViolations) {
            if ($violation.CatchLine -ge 0) {
                # Find the opening brace after catch
                $openBraceLine = -1
                for ($b = $violation.CatchLine; $b -lt $lines.Count; $b++) {
                    if ($lines[$b] -match '\{') {
                        $openBraceLine = $b
                        break
                    }
                }
                
                if ($openBraceLine -ge 0) {
                    # Get indentation of the line after the opening brace
                    $nextLine = $lines[$openBraceLine + 1]
                    $indent = if ($nextLine -match '^(\s+)') { $Matches[1] } else { "            " }
                    
                    # Insert: result.Success = false;
                    $fixLine = "${indent}result.Success = false;"
                    
                    # Insert after the opening brace
                    $lines = @(
                        $lines[0..$openBraceLine]
                        $fixLine
                        $lines[($openBraceLine + 1)..($lines.Count - 1)]
                    )
                    
                    Write-Host "   ✅ Fixed: Line $($violation.CatchLine + 1) (catch block)" -ForegroundColor Green
                }
            }
        }
        
        # Write fixed content back to file
        Set-Content -Path $filePath -Value $lines -NoNewline
        Write-Host "   💾 Saved: $fileName" -ForegroundColor Green
    } else {
        # Dry run - just show what would be fixed
        foreach ($violation in $fileViolations) {
            Write-Host "   Would fix: Line $($violation.SuccessLine + 1) → Line $($violation.ErrorLine + 1)" -ForegroundColor Yellow
            Write-Host "              Add: result.Success = false; at line $($violation.CatchLine + 1)" -ForegroundColor Gray
        }
    }
    
    Write-Host ""
}

# Summary
Write-Host "=" * 60 -ForegroundColor Cyan
Write-Host ""

if ($DryRun) {
    Write-Host "✅ Dry run complete - $($violations.Count) violations would be fixed" -ForegroundColor Yellow
    Write-Host ""
    Write-Host "Run without -DryRun to apply fixes:" -ForegroundColor White
    Write-Host "  .\scripts\fix-success-flags.ps1" -ForegroundColor Gray
} else {
    Write-Host "✅ Fixed $($violations.Count) violations in $($fileGroups.Count) files!" -ForegroundColor Green
    Write-Host ""
    
    if (-not $NoBackup) {
        Write-Host "Backups created with .bak extension" -ForegroundColor Gray
        Write-Host "To restore: Copy-Item *.bak -Destination (original name)" -ForegroundColor Gray
        Write-Host ""
    }
    
    Write-Host "Next steps:" -ForegroundColor Yellow
    Write-Host "  1. Verify fixes: .\scripts\check-success-flag.ps1" -ForegroundColor White
    Write-Host "  2. Build: dotnet build -c Release" -ForegroundColor White
    Write-Host "  3. Test: dotnet test --filter 'Category=Integration'" -ForegroundColor White
    Write-Host "  4. Commit: git add . && git commit" -ForegroundColor White
}

Write-Host ""
exit 0
