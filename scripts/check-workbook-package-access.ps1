#!/usr/bin/env pwsh
<#
.SYNOPSIS
    Rejects direct Excel workbook package XML access from production source.

.DESCRIPTION
    ExcelMcp controls workbooks through desktop Excel COM. ZIP/OOXML inspection
    is allowed in tests for fixture construction and concrete verification, but
    production source must not parse or modify workbook package parts.
#>

param(
    [string]$RootPath
)

$ErrorActionPreference = "Stop"

if ([string]::IsNullOrWhiteSpace($RootPath)) {
    $RootPath = Split-Path -Parent $PSScriptRoot
}

$sourceRoot = Join-Path $RootPath "src"
if (-not (Test-Path -LiteralPath $sourceRoot -PathType Container)) {
    throw "Source directory not found: $sourceRoot"
}

$violations = [System.Collections.Generic.List[object]]::new()
$workbookPartPattern = '(?i)(?:\[Content_Types\]\.xml|(?:^|["''\\/])xl[\\/](?:workbook|worksheets|_rels|sharedStrings|styles|theme|connections|pivot|charts?)[^"''\r\n]*\.xml)'
$openXmlApiPattern = '(?i)(?:DocumentFormat\.OpenXml|SpreadsheetDocument|OpenXmlPackage|WorkbookPart|WorksheetPart|System\.IO\.Packaging)'
$zipApiPattern = '(?i)(?:System\.IO\.Compression|ZipArchive|ZipFile)'
$xmlApiPattern = '(?i)(?:System\.Xml|XDocument|XmlDocument|XmlReader|XElement)'

foreach ($file in Get-ChildItem -LiteralPath $sourceRoot -Recurse -File |
             Where-Object { $_.Extension -in @(".cs", ".csproj", ".props", ".targets") }) {
    $relativeSourcePath = [IO.Path]::GetRelativePath($sourceRoot, $file.FullName)
    $pathSegments = $relativeSourcePath -split '[\\/]'
    if ($pathSegments -contains "bin" -or $pathSegments -contains "obj") {
        continue
    }

    $content = Get-Content -LiteralPath $file.FullName -Raw
    $reasons = [System.Collections.Generic.List[string]]::new()

    if ($content -match $openXmlApiPattern) {
        $reasons.Add("Open XML or package API")
    }
    if ($content -match $workbookPartPattern) {
        $reasons.Add("Excel workbook package part path")
    }
    if ($file.Extension -eq ".cs" -and
        $content -match $zipApiPattern -and
        $content -match $xmlApiPattern) {
        $reasons.Add("combined ZIP and XML processing")
    }

    if ($reasons.Count -gt 0) {
        $relativePath = [IO.Path]::GetRelativePath($RootPath, $file.FullName)
        $violations.Add([PSCustomObject]@{
            Path = $relativePath
            Reasons = $reasons -join ", "
        })
    }
}

if ($violations.Count -eq 0) {
    Write-Host "No production Excel workbook package XML access found." -ForegroundColor Green
    exit 0
}

Write-Host "Production Excel workbook package XML access is forbidden:" -ForegroundColor Red
foreach ($violation in $violations) {
    Write-Host "  $($violation.Path): $($violation.Reasons)" -ForegroundColor Yellow
}
Write-Host "Use Excel COM in production. ZIP/OOXML access belongs only in tests." -ForegroundColor Red
exit 1
