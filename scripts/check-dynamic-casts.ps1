#!/usr/bin/env pwsh
<#
.SYNOPSIS
    Checks that all ((dynamic)) casts in ExcelMcp.Core and ExcelMcp.ComInterop have justification comments.

.DESCRIPTION
    Every use of ((dynamic)) cast (explicit type coercion) must be preceded by a comment explaining
    why the PIA type cannot be used. Bare ((dynamic)) casts indicate potential PIA coverage gaps
    that weren't investigated.

    Valid comment prefixes (on the line immediately before the cast):
      // PIA gap: ...    — Type not in v16 Microsoft.Office.Interop.Excel PIA
      // TODO: ...       — Type IS in PIA but migration not yet done (tracked for removal)
      // Reason: ...     — Other documented reason for dynamic cast

    False positives are excluded:
      - ExcelBatch.cs / ExcelSession.cs / ExcelShutdownService.cs (infrastructure — uses `dynamic excel`)
      - Lines inside comments

.EXAMPLE
    .\check-dynamic-casts.ps1

.NOTES
    Run automatically as part of pre-commit hook.
    To add a new documented cast, place a comment ending in "// PIA gap:", "// TODO:", or "// Reason:"
    on the line immediately before the ((dynamic)) cast.
#>

param(
    [switch]$Verbose
)

$ErrorActionPreference = "Stop"
$rootDir = Split-Path -Parent $PSScriptRoot

$searchDirs = @(
    (Join-Path $rootDir "src\ExcelMcp.Core"),
    (Join-Path $rootDir "src\ExcelMcp.ComInterop")
)

# Files where bare dynamic casts are acceptable (infrastructure files)
$excludeFiles = @(
    "ExcelBatch.cs",
    "ExcelSession.cs",
    "ExcelShutdownService.cs",
    "ExcelShutdownHelper.cs"
)

$violations = @()
$checkedFiles = 0

foreach ($dir in $searchDirs) {
    $csFiles = Get-ChildItem -Path $dir -Filter "*.cs" -Recurse -ErrorAction SilentlyContinue
    foreach ($file in $csFiles) {
        if ($excludeFiles -contains $file.Name) {
            if ($Verbose) { Write-Host "   Skipped (infrastructure): $($file.Name)" -ForegroundColor Gray }
            continue
        }

        $checkedFiles++
        $lines = Get-Content $file.FullName
        for ($i = 0; $i -lt $lines.Count; $i++) {
            $line = $lines[$i]

            # Check for ((dynamic)) cast pattern
            if ($line -match '\(\(dynamic\)') {
                # Skip lines that are themselves comments
                $trimmed = $line.TrimStart()
                if ($trimmed.StartsWith("//")) { continue }

                # Check if any preceding comment line (within 5 lines) has a justification comment
                $hasJustification = $false
                for ($j = $i - 1; $j -ge 0 -and $j -ge ($i - 5); $j--) {
                    $prevLine = $lines[$j].TrimStart()
                    if ([string]::IsNullOrWhiteSpace($prevLine)) { continue }

                    # Once we hit a non-comment line, stop looking
                    if (-not $prevLine.StartsWith("//")) { break }

                    if ($prevLine.StartsWith("// PIA gap:") -or
                        $prevLine.StartsWith("// TODO:") -or
                        $prevLine.StartsWith("// Reason:") -or
                        $prevLine.StartsWith("// REASON:")) {
                        $hasJustification = $true
                        break
                    }
                }

                if (-not $hasJustification) {
                    $violations += [PSCustomObject]@{
                        File = $file.FullName.Replace($rootDir, "").TrimStart("\")
                        Line = $i + 1
                        Code = $line.Trim()
                    }
                }
            }
        }
    }
}

Write-Host "Checked $checkedFiles C# files for undocumented ((dynamic)) casts" -ForegroundColor Cyan

if ($violations.Count -eq 0) {
    Write-Host "All ((dynamic)) casts are documented" -ForegroundColor Green
    exit 0
}

Write-Host ""
Write-Host "UNDOCUMENTED ((dynamic)) CASTS FOUND: $($violations.Count)" -ForegroundColor Red
Write-Host ""
Write-Host "Every ((dynamic)) cast must have a comment on the preceding line explaining why:" -ForegroundColor Yellow
Write-Host "  // PIA gap: <type> not in Microsoft.Office.Interop.Excel v16 PIA because..." -ForegroundColor Gray
Write-Host "  // TODO: <type> IS in PIA, migration tracked — left as dynamic temporarily" -ForegroundColor Gray
Write-Host "  // Reason: <explanation>" -ForegroundColor Gray
Write-Host ""

foreach ($v in $violations) {
    Write-Host "  $($v.File):$($v.Line)" -ForegroundColor Yellow
    Write-Host "    $($v.Code)" -ForegroundColor Gray
}

Write-Host ""
Write-Host "Fix these before committing. See docs/PIA-COVERAGE.md for guidance." -ForegroundColor Red
exit 1
