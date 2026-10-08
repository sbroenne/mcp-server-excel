using Xunit;

namespace Sbroenne.ExcelMcp.ScriptSafety.Tests;

[Collection("Sequential")]
[Trait("RequiresExcel", "false")]
[Trait("Feature", "PreCommit")]
public sealed class MacAppleEventCodeGenerationTests
{
    [Fact]
    public async Task HelperSourcePreparationUsesIndependentVersionWithoutExcelOrWorkbookFiles()
    {
        var result = await ValidationSelectionTests.RunAsync("""
            $directory = Join-Path ([IO.Path]::GetTempPath()) ('excelmcp-helper-source-' + [guid]::NewGuid().ToString('N'))
            try {
                & ./scripts/Build-MacHelper.ps1 -PrepareOnly -OutputDirectory $directory
                $version = (Get-Content ./helper/mac/VERSION -Raw).Trim()
                $module = Get-Content (Join-Path $directory 'ExcelMcpHelper.bas') -Raw
                if (-not $module.Contains('Private Const HelperVersion As String = "' + $version + '"')) { throw 'Independent version was not injected.' }
                if ($module.Contains('@@HELPER_VERSION@@')) { throw 'Unresolved helper source placeholder.' }
                if (@(Get-ChildItem $directory).Count -ne 1) { throw 'Source preparation created an unexpected artifact.' }
            } finally {
                if ([IO.Directory]::Exists($directory)) { [IO.Directory]::Delete($directory, $true) }
            }
            """);
        Assert.True(result.ExitCode == 0, result.Output);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task DictionaryExtraction_ValidatesEveryRequiredSymbolBeforeWriting(bool missingWorkbook)
    {
        var result = await ValidationSelectionTests.RunAsync($$"""
            $directory = Join-Path ([IO.Path]::GetTempPath()) ('excelmcp-ae-codes-' + [guid]::NewGuid().ToString('N'))
            [void][IO.Directory]::CreateDirectory($directory)
            try {
                $excel = Join-Path $directory 'Excel.sdef'
                $standard = Join-Path $directory 'Standard.sdef'
                $output = Join-Path $directory 'MacExcelDictionary.g.cs'
                Set-Content $excel @'
            <dictionary><suite>
              <class name="application" code="capp"><property name="calculation" code="1208"/></class>
              <class name="{{(missingWorkbook ? "unrelated" : "workbook")}}" code="X141">
                <property name="name" code="pnam"/>
                <property name="full name" code="1773"/>
                <property name="active sheet" code="1107"/>
              </class>
              <class name="window" code="cwin"><property name="visible" code="pvis"/></class>
              <class name="worksheet" code="XwSH"/>
              <class name="chart sheet" code="XcSH"/>
              <class name="row" code="crow"/>
              <class name="column" code="ccol"/>
              <class name="range" code="X117">
                <property name="formula2" code="2122"/>
                <property name="formula2 r1c1" code="F2rc"/>
                <property name="value2" code="DPV2"/>
                <property name="has formula" code="1573"/>
                <property name="merge cells" code="1588"/>
                <property name="merge area" code="1587"/>
                <property name="first row index" code="XfrX"/>
                <property name="first column index" code="XfcX"/>
              </class>
              <class name="validation" code="X136">
                <property name="formula2" code="vFm2"/>
                <property name="formula2 r1c1" code="vFrc"/>
              </class>
              <command name="run VB macro" code="sTBL2620"><parameter name="arg1" code="5040"/></command>
              <command name="get address" code="sTBL1515"><parameter name="external" code="5123"/></command>
              <command name="evaluate" code="smXL2435"><parameter name="name" code="pnam"/></command>
              <command name="calculate" code="smXL1175"><direct-parameter type="4004"/></command>
              <command name="calculate" code="sTBL1175"><direct-parameter type="range"/></command>
              <enumeration name="XlSheetVisibility" code="e225">
                <enumerator name="sheet visible" code="0x0270ffff"/>
                <enumerator name="sheet hidden" code="0x02710000"/>
                <enumerator name="sheet very hidden" code="0x02710002"/>
              </enumeration>
              <enumeration name="XlCalculation" code="e174">
                <enumerator name="calculation automatic" code="0x023deff7"/>
                <enumerator name="calculation manual" code="0x023defd9"/>
                <enumerator name="calculation semiautomatic" code="0x023e0002"/>
              </enumeration>
            </suite></dictionary>
            '@
                Set-Content $standard @'
            <dictionary><suite>
              <command name="get" code="coregetd"/>
              <command name="set" code="coresetd"/>
              <command name="close" code="coreclos"/>
              <command name="make" code="corecrel"><parameter name="new" code="kocl"/><parameter name="at" code="insh"/></command>
              <command name="count" code="corecnte"><parameter name="each" code="kocl"/></command>
            </suite></dictionary>
            '@
                $failure = $null
                try {
                    & ./scripts/Update-MacAppleEventCodes.ps1 -DictionaryPath $excel -StandardDictionaryPath $standard -OutputPath $output
                } catch { $failure = $_ }
                if (${{(missingWorkbook ? "true" : "false")}}) {
                    if (-not $failure -or $failure -notmatch 'WorkbookClass') { throw 'Missing workbook symbol was not diagnosed.' }
                    if (Test-Path $output) { throw 'Invalid dictionary produced an output artifact.' }
                } else {
                    if ($failure) { throw $failure }
                    $source = Get-Content $output -Raw
                    foreach ($expected in @('WorkbookClass = 0x58313431u', 'FullName = 0x31373733u', 'CloseId = 0x636C6F73u', 'WorksheetClass = 0x58775348u', 'SheetVisible = 0x0270FFFFu', 'RangeClass = 0x58313137u', 'Formula2 = 0x32313232u', 'Formula2R1C1 = 0x46327263u', 'Value2 = 0x44505632u', 'HasFormula = 0x31353733u', 'MergeCells = 0x31353838u', 'MergeArea = 0x31353837u', 'FirstRowIndex = 0x58667258u', 'FirstColumnIndex = 0x58666358u')) {
                        if (-not $source.Contains($expected)) { throw "Missing native code: $expected" }
                    }
                    if ($source.Contains($directory)) { throw 'Generated code leaked a local dictionary path.' }
                }
            } finally { [IO.Directory]::Delete($directory, $true) }
            """);
        Assert.True(result.ExitCode == 0, result.Output);
    }
}
