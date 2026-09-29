#!/usr/bin/env pwsh
# Detect high-risk Excel COM access patterns that hide intermediate RCWs.

param(
    [string[]]$InputPath
)

$ErrorActionPreference = "Stop"
$rootDir = Split-Path -Parent $PSScriptRoot

$rules = @(
    @{
        Name = "chained COM property access"
        Pattern = '\.(Rows|Columns|ListColumns|ListRows|Range|TableRange1|TableRange2|ChartArea|Parent|Application|Cells)\s*(?:\[[^\r\n]*\])?\s*\.(Count|Address|AutoFilter|Item|Range|Row|Rows|Column|Columns|Parent|Application|SeriesCollection|Formula|Name)\b'
        Guidance = "Capture each COM object in a local variable and release it in finally."
    },
    @{
        Name = "discarded Names.Add result"
        Pattern = '^\s*namesCollection\.Add\s*\('
        Guidance = "Capture the returned Excel.Name and release it in finally."
    }
)

function Get-ScanFiles {
    param([string[]]$RequestedPaths)

    if ($RequestedPaths.Count -gt 0) {
        foreach ($requestedPath in $RequestedPaths) {
            $resolvedPath = (Resolve-Path -LiteralPath $requestedPath).Path
            if (Test-Path -LiteralPath $resolvedPath -PathType Container) {
                Get-ChildItem -LiteralPath $resolvedPath -Recurse -File -Filter "*.cs"
            }
            else {
                Get-Item -LiteralPath $resolvedPath
            }
        }
        return
    }

    $relativePaths = @(
        "src\ExcelMcp.Core\Commands\Connection",
        "src\ExcelMcp.Core\Commands\Table\TableCommands.Data.cs",
        "src\ExcelMcp.Core\Commands\Table\TableCommands.Lifecycle.cs",
        "src\ExcelMcp.Core\Commands\Table\TableCommands.Filters.cs",
        "src\ExcelMcp.Core\Commands\Table\TableCommands.StructuredReferences.cs",
        "src\ExcelMcp.Core\Commands\PivotTable\PivotTableCommands.Create.cs",
        "src\ExcelMcp.Core\Commands\Chart\RegularChartStrategy.cs",
        "src\ExcelMcp.Core\Commands\Chart\ChartCommands.Lifecycle.cs",
        "src\ExcelMcp.Core\Commands\Chart\ChartCommands.DataSource.cs",
        "src\ExcelMcp.Core\Commands\NamedRange\NamedRangeCommands.Operations.cs"
    )

    foreach ($relativePath in $relativePaths) {
        $fullPath = Join-Path $rootDir $relativePath
        if (Test-Path -LiteralPath $fullPath -PathType Container) {
            Get-ChildItem -LiteralPath $fullPath -Recurse -File -Filter "*.cs"
        }
        elseif (Test-Path -LiteralPath $fullPath -PathType Leaf) {
            Get-Item -LiteralPath $fullPath
        }
    }
}

Write-Host "Scanning high-risk Excel COM access patterns..." -ForegroundColor Yellow

$findings = @()
$files = @(Get-ScanFiles -RequestedPaths $InputPath |
    Where-Object { $_.FullName -notmatch '[\\/](bin|obj)[\\/]' -and $_.Name -notmatch '\.g\.cs$' } |
    Sort-Object -Property FullName -Unique)
if ($files.Count -eq 0) { throw 'No source files found for the high-risk COM access pattern guard.' }

foreach ($file in $files) {
    $lineNumber = 0
    foreach ($line in Get-Content -LiteralPath $file.FullName) {
        $lineNumber++
        $trimmed = $line.TrimStart()
        if ($trimmed.StartsWith("//") -or $trimmed.StartsWith("*")) {
            continue
        }

        foreach ($rule in $rules) {
            if ($line -match $rule.Pattern) {
                $displayPath = if ($file.FullName.StartsWith($rootDir, [StringComparison]::OrdinalIgnoreCase)) {
                    [IO.Path]::GetRelativePath($rootDir, $file.FullName)
                }
                else {
                    $file.FullName
                }

                $findings += [pscustomobject]@{
                    Path = $displayPath
                    Line = $lineNumber
                    Rule = $rule.Name
                    Guidance = $rule.Guidance
                    Text = $trimmed
                }
            }
        }
    }
}

if ($findings.Count -gt 0) {
    foreach ($finding in $findings) {
        Write-Host "$($finding.Path):$($finding.Line) - $($finding.Rule)" -ForegroundColor Red
        Write-Host "  $($finding.Text)" -ForegroundColor DarkRed
        Write-Host "  $($finding.Guidance)" -ForegroundColor Yellow
    }

    Write-Host ""
    Write-Host "$($findings.Count) high-risk COM access pattern(s) detected." -ForegroundColor Red
    exit 1
}

Write-Host "No high-risk COM access patterns detected in $($files.Count) file(s)." -ForegroundColor Green
exit 0
