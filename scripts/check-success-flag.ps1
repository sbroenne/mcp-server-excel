#!/usr/bin/env pwsh
<#
.SYNOPSIS
    Checks for Success = true with ErrorMessage violations (Rule 0)

.DESCRIPTION
    Scans Core Commands for the critical bug pattern:
    - result.Success = true (set optimistically)
    - catch block sets result.ErrorMessage
    - but FORGETS to set result.Success = false
    
    This causes Success=true responses with error messages, confusing LLMs.

.EXAMPLE
    .\check-success-flag.ps1
    
.NOTES
    Part of pre-commit validation. See CRITICAL-RULES.md Rule 0.
#>

$ErrorActionPreference = "Stop"
$rootDir = Split-Path -Parent $PSScriptRoot

Write-Host "Checking for Success = true with ErrorMessage violations (Rule 0)..." -ForegroundColor Cyan
Write-Host ""

$violations = @()

$files = @(Get-ChildItem -LiteralPath "$rootDir\src\ExcelMcp.Core\Commands" -File -Filter "*.cs" -Recurse |
    Where-Object { $_.FullName -notmatch '[\\/](bin|obj)[\\/]' -and $_.Name -notmatch '\.g\.cs$' })
if ($files.Count -eq 0) { throw 'No command source files found for the success flag pattern guard.' }
$files | ForEach-Object {
    $lines = @(Get-Content -LiteralPath $_.FullName)
    
    for ($i = 0; $i -lt $lines.Count; $i++) {
        # Look for Success = true
        if ($lines[$i] -match '\.Success\s*=\s*true') {
            $successLine = $i + 1
            
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
                    
                    $errorLine = $j + 1
                    
                    # Check if Success is set to false between these lines
                    $hasSuccessFalse = $false
                    for ($k = $i + 1; $k -lt $j; $k++) {
                        if ($lines[$k] -match '\.Success\s*=\s*false') {
                            $hasSuccessFalse = $true
                            break
                        }
                    }
                    
                    if (-not $hasSuccessFalse) {
                        $violations += [PSCustomObject]@{
                            File = $_.FullName.Replace("$rootDir\", "")
                            SuccessLine = $successLine
                            ErrorMessageLine = $errorLine
                            SuccessCode = $lines[$i].Trim()
                            ErrorCode = $lines[$j].Trim()
                        }
                    }
                    break
                }
            }
        }
    }
}

# Report results
if ($violations.Count -eq 0) {
    Write-Host "No nearby conflicting Success/ErrorMessage assignments found; this pattern check is not control-flow analysis." -ForegroundColor Green
    exit 0
}

Write-Host "Found $($violations.Count) Rule 0 violations!" -ForegroundColor Red
Write-Host ""
Write-Host "Pattern: Success = true followed by ErrorMessage = ... without Success = false" -ForegroundColor Yellow
Write-Host ""

foreach ($v in $violations) {
    Write-Host "$($v.File)" -ForegroundColor Red
    Write-Host "   Line $($v.SuccessLine): $($v.SuccessCode)" -ForegroundColor Yellow
    Write-Host "   Line $($v.ErrorMessageLine): $($v.ErrorCode)" -ForegroundColor Yellow
    Write-Host "   -> Missing: result.Success = false; before ErrorMessage" -ForegroundColor Magenta
    Write-Host ""
}

Write-Host "Action Required:" -ForegroundColor Yellow
Write-Host "  1. In each catch block, add: result.Success = false;" -ForegroundColor Yellow
Write-Host "  2. Or move: result.Success = true; to INSIDE try block (after work succeeds)" -ForegroundColor Yellow
Write-Host "  3. See CRITICAL-RULES.md Rule 0 for pattern examples" -ForegroundColor Yellow
Write-Host ""
Write-Host "See: .github/instructions/CRITICAL-RULES.md Rule 0" -ForegroundColor Gray

exit 1