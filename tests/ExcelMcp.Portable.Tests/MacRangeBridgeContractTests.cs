using System.Reflection;
using Sbroenne.ExcelMcp.Service.Mac;
using Xunit;

namespace Sbroenne.ExcelMcp.Portable.Tests;

public sealed class MacRangeBridgeContractTests
{
    private const string ResourceName = "Sbroenne.ExcelMcp.Service.Mac.MacExcelBridge.js";

    [Fact]
    public void NumericReadsAndCopies_UseUnderlyingValue2()
    {
        var script = ReadBridgeScript();

        Assert.Contains("target.value2 = normalizeMatrix(source.value2())", script, StringComparison.Ordinal);
        Assert.DoesNotContain("range.value()", script, StringComparison.Ordinal);
        Assert.DoesNotContain("formulaRange.value()", script, StringComparison.Ordinal);
        Assert.DoesNotContain("changingRange.value()", script, StringComparison.Ordinal);
    }

    [Fact]
    public void SheetList_NormalizesVisibilityInsteadOfComparingToBooleanFalse()
    {
        var script = ReadBridgeScript();

        Assert.Contains("visible: currentSheetVisibility(sheets[index]).value === -1", script, StringComparison.Ordinal);
        Assert.DoesNotContain("visible() !== false", script, StringComparison.Ordinal);
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
    public void MatrixWrites_ValidateMergedTargetsAndExactShape(
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

        Assert.Contains("requireUnmergedWriteTarget(range)", branch, StringComparison.Ordinal);
        Assert.Contains("requireMatrixShape(", branch, StringComparison.Ordinal);
        Assert.Contains(
            $"{matrixExpression}, range.rows.length, range.columns.length, \"{parameterName}\")",
            branch,
            StringComparison.Ordinal);
    }

    private static string ReadBridgeScript()
    {
        using var stream = typeof(MacAutomationHost).Assembly.GetManifestResourceStream(ResourceName);
        Assert.NotNull(stream);
        using var reader = new StreamReader(stream);
        return reader.ReadToEnd();
    }
}
