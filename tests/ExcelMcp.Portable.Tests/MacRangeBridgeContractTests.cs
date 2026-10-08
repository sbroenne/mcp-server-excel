using System.Reflection;
using System.Text.Json.Nodes;
using Sbroenne.ExcelMcp.Service.Mac;
using Xunit;

namespace Sbroenne.ExcelMcp.Portable.Tests;

[Trait("RequiresExcel", "false")]
public sealed class MacRangeBridgeContractTests
{
    private const string ResourceName = "Sbroenne.ExcelMcp.Service.Mac.MacExcelBridge.js";

    [Fact]
    public void NumericReadsAndCopies_UseUnderlyingValue2()
    {
        var script = ReadBridgeScript();

        Assert.Contains("target.value2 = tileMatrix(", script, StringComparison.Ordinal);
        Assert.Contains("source.value2(), target.rows.length, target.columns.length", script, StringComparison.Ordinal);
        Assert.Contains(
            "formulaErrorCode(",
            script,
            StringComparison.Ordinal);
        Assert.DoesNotContain("formulaRange.value()", script, StringComparison.Ordinal);
        Assert.DoesNotContain("changingRange.value()", script, StringComparison.Ordinal);
    }

    [Fact]
    public void SheetList_NormalizesVisibilityInsteadOfComparingToBooleanFalse()
    {
        var result = MacNativeWorksheet.CreateList("owned.xlsx",
            new JsonArray("Visible", "Hidden", "VeryHidden"),
            new JsonArray(MacExcelDictionary.SheetVisible, MacExcelDictionary.SheetHidden, MacExcelDictionary.SheetVeryHidden));
        Assert.True(result.Success);
        Assert.True(string.IsNullOrEmpty(result.ErrorMessage));
        Assert.Equal("owned.xlsx", result.FilePath);
        Assert.Equal([1, 2, 3], result.Worksheets.Select(sheet => sheet.Index));
        Assert.Equal([true, false, false], result.Worksheets.Select(sheet => sheet.Visible));
    }

    [Theory]
    [InlineData("null", "[]")]
    [InlineData("[\"Sheet1\"]", "[]")]
    [InlineData("[\"\"]", "[4095]")]
    [InlineData("[\"Sheet1\"]", "[4095]")]
    public void SheetList_RejectsMalformedOrUnknownNativeState(string names, string visibility)
    {
        Assert.Throws<InvalidOperationException>(() => MacNativeWorksheet.CreateList(
            "owned.xlsx", JsonNode.Parse(names), JsonNode.Parse(visibility)));
    }

    [Fact]
    public void WorkbookIdentity_ResolvesSymlinksBeforeComparingPaths()
    {
        var script = ReadBridgeScript();

        Assert.Contains("function canonicalPath(filePath)", script, StringComparison.Ordinal);
        Assert.Contains("canonicalPath(fullName) === target", script, StringComparison.Ordinal);
    }

    [Fact]
    public void AutoFit_UsesDeclaredApplicationCommandWithRowAndColumnRanges()
    {
        var script = ReadBridgeScript();

        Assert.Contains("excel.autofit(column)", script, StringComparison.Ordinal);
        Assert.Contains("excel.autofit(row)", script, StringComparison.Ordinal);
        Assert.Contains(
            "sheet.ranges.byName(address).columnWidth = fittedWidth",
            script,
            StringComparison.Ordinal);
        Assert.Contains(
            "sheet.ranges.byName(address).rowHeight = fittedHeight",
            script,
            StringComparison.Ordinal);
        Assert.DoesNotContain("range.columns.autofit()", script, StringComparison.Ordinal);
        Assert.DoesNotContain("range.rows.autofit()", script, StringComparison.Ordinal);
    }

    [Fact]
    public void MergeMutations_UseDeclaredCommandsAndVerifyFreshRangeState()
    {
        var script = ReadBridgeScript();

        Assert.Contains("excel.merge(range)", script, StringComparison.Ordinal);
        Assert.Contains("excel.unmerge(range)", script, StringComparison.Ordinal);
        Assert.Contains(
            "sheet.ranges.byName(args.rangeAddress).mergeCells()",
            script,
            StringComparison.Ordinal);
        Assert.DoesNotContain("range.merge()", script, StringComparison.Ordinal);
        Assert.DoesNotContain("range.unmerge()", script, StringComparison.Ordinal);
    }

    [Fact]
    public void Calculate_HandlesWorksheetScopeWithoutFallingBackToWorkbook()
    {
        var script = ReadBridgeScript();

        Assert.Contains(
            "else if (args.scope === \"sheet\")",
            script,
            StringComparison.Ordinal);
        Assert.Contains(
            "worksheetByName(workbook, args.sheetName).calculate()",
            script,
            StringComparison.Ordinal);
    }

    [Theory]
    [InlineData("range.set-values", "args.values", "values")]
    [InlineData("range.set-formulas", "args.formulas", "formulas")]
    public void MatrixWrites_ValidateMergedTargetsShapeAndOverwritePolicy(
        string command,
        string matrixExpression,
        string parameterName)
    {
        var script = ReadBridgeScript();
        var commandStart = script.IndexOf(
            $"if (command === \"{command}\")",
            StringComparison.Ordinal);
        Assert.True(commandStart >= 0);
        var commandEnd = script.IndexOf("\n            }", commandStart, StringComparison.Ordinal);
        Assert.True(commandEnd > commandStart);
        var branch = script[commandStart..commandEnd];

        Assert.Contains("requireWritableMergeTarget(excel, range)", branch, StringComparison.Ordinal);
        Assert.Contains("requireMatrixShape(", branch, StringComparison.Ordinal);
        Assert.Contains(
            "requireWritableDestination(excel, range, args.overwritePolicy)",
            branch,
            StringComparison.Ordinal);
        Assert.Contains(
            $"{matrixExpression}, range.rows.length, range.columns.length, \"{parameterName}\")",
            branch,
            StringComparison.Ordinal);
    }

    [Fact]
    public void MergedWriteValidation_FailsClosedWithoutMergeAreaReadback()
    {
        var script = ReadBridgeScript();

        Assert.Contains(
            "cannot reliably identify the merged range's top-left cell",
            script,
            StringComparison.Ordinal);
        Assert.DoesNotContain("const mergeArea = range.mergeArea()", script, StringComparison.Ordinal);
    }

    [Fact]
    public void ProtectedCopies_KeepCompatibleExplicitTargetsAndRepeatMatrices()
    {
        var script = ReadBridgeScript();

        Assert.Contains("const target = copyTargetRange(", script, StringComparison.Ordinal);
        Assert.Contains("requested.rows.length % sourceRows !== 0", script, StringComparison.Ordinal);
        Assert.Contains("requested.columns.length % sourceColumns !== 0", script, StringComparison.Ordinal);
        Assert.Contains(
            "requireWritableDestination(excel, target, args.overwritePolicy)",
            script,
            StringComparison.Ordinal);
        Assert.Contains(
            "source.value2(), target.rows.length, target.columns.length",
            script,
            StringComparison.Ordinal);
        Assert.Contains(
            "source.formulaR1c1(), target.rows.length, target.columns.length",
            script,
            StringComparison.Ordinal);
    }

    [Fact]
    public void RangeReads_NormalizeFormulaErrorsAndReturnCellMetadata()
    {
        var script = ReadBridgeScript();

        Assert.Contains("function normalizedRangeRead(", script, StringComparison.Ordinal);
        Assert.Contains("\"-2146826265\": [\"#REF!\"", script, StringComparison.Ordinal);
        Assert.Contains("\"2023\": [\"#REF!\"", script, StringComparison.Ordinal);
        Assert.Contains("ERROR.TYPE(${reference})", script, StringComparison.Ordinal);
        Assert.Contains("cellAddress: `${columnName(cellColumn)}${cellRow}`", script, StringComparison.Ordinal);
        Assert.Contains("cellErrors: read.cellErrors", script, StringComparison.Ordinal);
    }

    private static string ReadBridgeScript()
    {
        using var stream = typeof(MacAutomationHost).Assembly.GetManifestResourceStream(ResourceName);
        Assert.NotNull(stream);
        using var reader = new StreamReader(stream);
        return reader.ReadToEnd();
    }
}
