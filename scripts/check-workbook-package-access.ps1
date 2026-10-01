#!/usr/bin/env pwsh
<#
.SYNOPSIS
    Rejects direct workbook package access in production, tests, and scripts.

.DESCRIPTION
    Workbook files are opaque. ExcelMcp may copy an intact Excel-authored
    workbook, but it must not create, inspect, parse, or mutate workbook ZIP,
    OOXML, relationship, custom XML, or DataMashup internals.

    ZIP handling used only for distribution/plugin packaging is explicitly
    excluded because those archives are not workbook files.
#>

param(
    [string]$RootPath
)

$ErrorActionPreference = 'Stop'
$rootDir = if ([string]::IsNullOrWhiteSpace($RootPath)) {
    Split-Path -Parent $PSScriptRoot
}
else {
    [IO.Path]::GetFullPath($RootPath)
}
$scanRoots = @('src', 'tests', 'scripts')
$sourceExtensions = @('.cs', '.js', '.mjs', '.ps1', '.psm1', '.sh')
$distributionZipAllowList = @(
    'scripts/Build-ReleasePackages.ps1',
    'scripts/Test-DistributionPackages.ps1'
)

$violations = [Collections.Generic.List[object]]::new()
foreach ($scanRoot in $scanRoots) {
    $absoluteRoot = Join-Path $rootDir $scanRoot
    if (-not (Test-Path -LiteralPath $absoluteRoot -PathType Container)) {
        continue
    }
    foreach ($file in Get-ChildItem -LiteralPath $absoluteRoot -Recurse -File) {
        if ($sourceExtensions -notcontains $file.Extension -or
            $file.FullName -match '[/\\](bin|obj|node_modules)[/\\]') {
            continue
        }

        $relativePath = [IO.Path]::GetRelativePath($rootDir, $file.FullName)
        $relativePath = $relativePath.Replace('\', '/')
        if ($relativePath -in @(
                'scripts/check-workbook-package-access.ps1',
                'tests/ExcelMcp.SkillGeneration.Tests/WorkbookPackageAccessGuardTests.cs')) {
            continue
        }
        $content = @(Get-Content -LiteralPath $file.FullName)
        for ($lineIndex = 0; $lineIndex -lt $content.Count; $lineIndex++) {
            $line = $content[$lineIndex]
            $reason = $null
            $normalizedLine = $line.Replace('\', '/')
            if ($normalizedLine.Contains(
                    'xl/workbook.xml',
                    [StringComparison]::OrdinalIgnoreCase) -or
                $line -match '(?i)(\[Content_Types\]\.xml|DataMashup|MacPowerQueryPackage|MacWorkbookPackage)') {
                $reason = 'workbook package internals'
            }
            elseif ($line -match '(?i)(System\.IO\.Packaging|DocumentFormat\.OpenXml)') {
                $reason = 'workbook package library'
            }
            elseif ($line -match '(?i)(System\.IO\.Compression|\bZipArchive\b|\bZipFile\b)') {
                $isDistributionArchive =
                    $relativePath.StartsWith(
                        'tests/ExcelMcp.SkillGeneration.Tests/',
                        [StringComparison]::Ordinal) -or
                    $distributionZipAllowList -contains $relativePath
                if (-not $isDistributionArchive) {
                    $reason = 'direct ZIP access outside distribution packaging'
                }
            }

            if ($reason) {
                $violations.Add([PSCustomObject]@{
                    File = $relativePath
                    Line = $lineIndex + 1
                    Reason = $reason
                    Code = $line.Trim()
                })
            }
        }
    }
}

if ($violations.Count -eq 0) {
    Write-Host 'Workbook package access check passed' -ForegroundColor Green
    exit 0
}

Write-Host 'Direct workbook package access is prohibited:' -ForegroundColor Red
foreach ($violation in $violations) {
    Write-Host (
        "  {0}:{1}: {2}: {3}" -f
        $violation.File,
        $violation.Line,
        $violation.Reason,
        $violation.Code) -ForegroundColor Yellow
}
exit 1
