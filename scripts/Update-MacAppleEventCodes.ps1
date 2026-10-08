#Requires -Version 7.0
<#
.SYNOPSIS
Regenerates the native Excel Apple Event constants from supported API dictionaries.
.DESCRIPTION
Maintainers run this on a Mac with Excel after adding a native primitive. The
checked-in output lets ordinary builds run without Excel or a macOS SDK.
DictionaryPath and StandardDictionaryPath allow Excel-free extraction tests.
#>
[CmdletBinding()]
param(
    [string]$ExcelApplication = '/Applications/Microsoft Excel.app',
    [string]$DictionaryPath,
    [string]$StandardDictionaryPath = '/System/Library/ScriptingDefinitions/CocoaStandard.sdef',
    [string]$OutputPath = (Join-Path $PSScriptRoot '../src/ExcelMcp.Service/Mac/MacExcelDictionary.g.cs')
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Read-ApiDictionary([string]$Text) {
    $settings = [Xml.XmlReaderSettings]::new()
    $settings.DtdProcessing = [Xml.DtdProcessing]::Ignore
    $settings.XmlResolver = $null
    $textReader = [IO.StringReader]::new($Text)
    $reader = [Xml.XmlReader]::Create($textReader, $settings)
    try {
        $document = [Xml.XmlDocument]::new()
        $document.XmlResolver = $null
        $document.Load($reader)
        return ,$document
    }
    finally {
        $reader.Dispose()
        $textReader.Dispose()
    }
}

if ($DictionaryPath) {
    $excelText = Get-Content -LiteralPath $DictionaryPath -Raw
} else {
    if (-not $IsMacOS) { throw 'Extracting the installed Excel dictionary requires macOS.' }
    $excelText = (& /usr/bin/sdef $ExcelApplication) -join "`n"
    if ($LASTEXITCODE -ne 0) { throw "Excel dictionary extraction failed with exit code $LASTEXITCODE." }
}
$excel = Read-ApiDictionary $excelText
# Excel explicitly imports CocoaStandard.sdef; resolve only this known SDK dictionary.
$standard = Read-ApiDictionary (Get-Content -LiteralPath $StandardDictionaryPath -Raw)
$symbols = [ordered]@{
    WorkbookClass = @($excel, "//class[@name='workbook']/@code", 4)
    WorksheetClass = @($excel, "//class[@name='worksheet']/@code", 4)
    ChartSheetClass = @($excel, "//class[@name='chart sheet']/@code", 4)
    Calculation = @($excel, "//class[@name='application']/property[@name='calculation']/@code", 4)
    CalculationAutomatic = @($excel, "//enumeration[@name='XlCalculation']/enumerator[@name='calculation automatic']/@code", 4)
    CalculationManual = @($excel, "//enumeration[@name='XlCalculation']/enumerator[@name='calculation manual']/@code", 4)
    CalculationSemiautomatic = @($excel, "//enumeration[@name='XlCalculation']/enumerator[@name='calculation semiautomatic']/@code", 4)
    CalculateSheet = @($excel, "//command[@name='calculate'][direct-parameter[@type='4004']]/@code", 8)
    CalculateRange = @($excel, "//command[@name='calculate'][direct-parameter[@type='range']]/@code", 8)
    RangeClass = @($excel, "//class[@name='range']/@code", 4)
    Formula2 = @($excel, "//class[@name='range']/property[@name='formula2']/@code", 4)
    Formula2R1C1 = @($excel, "//class[@name='range']/property[@name='formula2 r1c1']/@code", 4)
    Value2 = @($excel, "//class[@name='range']/property[@name='value2']/@code", 4)
    HasFormula = @($excel, "//class[@name='range']/property[@name='has formula']/@code", 4)
    MergeCells = @($excel, "//class[@name='range']/property[@name='merge cells']/@code", 4)
    MergeArea = @($excel, "//class[@name='range']/property[@name='merge area']/@code", 4)
    FirstRowIndex = @($excel, "//class[@name='range']/property[@name='first row index']/@code", 4)
    FirstColumnIndex = @($excel, "//class[@name='range']/property[@name='first column index']/@code", 4)
    RowClass = @($excel, "//class[@name='row']/@code", 4)
    ColumnClass = @($excel, "//class[@name='column']/@code", 4)
    GetAddress = @($excel, "//command[@name='get address']/@code", 8)
    AddressExternalParameter = @($excel, "//command[@name='get address']/parameter[@name='external']/@code", 4)
    Evaluate = @($excel, "//command[@name='evaluate']/@code", 8)
    EvaluateNameParameter = @($excel, "//command[@name='evaluate']/parameter[@name='name']/@code", 4)
    WindowClass = @($excel, "//class[@name='window']/@code", 4)
    Name = @($excel, "//class[@name='workbook']/property[@name='name']/@code", 4)
    FullName = @($excel, "//class[@name='workbook']/property[@name='full name']/@code", 4)
    ActiveSheet = @($excel, "//class[@name='workbook']/property[@name='active sheet']/@code", 4)
    Visible = @($excel, "//class[@name='window']/property[@name='visible']/@code", 4)
    SheetVisible = @($excel, "//enumeration[@name='XlSheetVisibility']/enumerator[@name='sheet visible']/@code", 4)
    SheetHidden = @($excel, "//enumeration[@name='XlSheetVisibility']/enumerator[@name='sheet hidden']/@code", 4)
    SheetVeryHidden = @($excel, "//enumeration[@name='XlSheetVisibility']/enumerator[@name='sheet very hidden']/@code", 4)
    Close = @($standard, "//command[@name='close']/@code", 8)
    Make = @($standard, "//command[@name='make']/@code", 8)
    MakeClassParameter = @($standard, "//command[@name='make']/parameter[@name='new']/@code", 4)
    MakeLocationParameter = @($standard, "//command[@name='make']/parameter[@name='at']/@code", 4)
    Count = @($standard, "//command[@name='count']/@code", 8)
    CountClassParameter = @($standard, "//command[@name='count']/parameter[@name='each']/@code", 4)
    RunMacro = @($excel, "//command[@name='run VB macro']/@code", 8)
    MacroArgument1 = @($excel, "//command[@name='run VB macro']/parameter[@name='arg1']/@code", 4)
}

function ConvertTo-FourCharCode([string]$Code) {
    if ($Code -match '^0x([0-9a-fA-F]{8})$') {
        return ('0x{0}u' -f $Matches[1].ToUpperInvariant())
    }
    $value = [uint32]0
    foreach ($character in $Code.ToCharArray()) {
        $value = ($value -shl 8) -bor [uint32][char]$character
    }
    return ('0x{0:X8}u' -f $value)
}

$lines = [Collections.Generic.List[string]]::new()
$lines.Add('// <auto-generated by scripts/Update-MacAppleEventCodes.ps1 />')
$lines.Add('namespace Sbroenne.ExcelMcp.Service.Mac;')
$lines.Add('')
$lines.Add('internal static class MacExcelDictionary')
$lines.Add('{')
foreach ($symbol in $symbols.GetEnumerator()) {
    $document, $query, $length = $symbol.Value
    $codes = @($document.SelectNodes($query) | ForEach-Object Value | Sort-Object -Unique)
    if ($codes.Count -ne 1 -or
        ($codes[0].Length -ne $length -and -not ($length -eq 4 -and $codes[0] -match '^0x[0-9a-fA-F]{8}$')) -or
        $codes[0] -match '[^\x20-\x7e]') {
        throw "Required Excel API symbol '$($symbol.Key)' is missing, ambiguous, or malformed."
    }
    if ($length -eq 8) {
        $lines.Add("    internal const uint $($symbol.Key)Class = $(ConvertTo-FourCharCode $codes[0].Substring(0, 4));")
        $lines.Add("    internal const uint $($symbol.Key)Id = $(ConvertTo-FourCharCode $codes[0].Substring(4, 4));")
    } else {
        $lines.Add("    internal const uint $($symbol.Key) = $(ConvertTo-FourCharCode $codes[0]);")
    }
}
$lines.Add('}')
$source = ($lines -join "`n") + "`n"
[IO.File]::WriteAllText([IO.Path]::GetFullPath($OutputPath), $source, [Text.UTF8Encoding]::new($false))
