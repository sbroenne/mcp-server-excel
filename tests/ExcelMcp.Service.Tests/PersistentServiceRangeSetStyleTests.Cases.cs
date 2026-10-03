using System.Globalization;
using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Xunit;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sbroenne.ExcelMcp.Service.Tests;

public sealed partial class PersistentServiceRangeSetStyleTests
{
    [Fact]
    public void SetStyle_Heading1_AppliesSuccessfully()
    {
        // Arrange & Act
        var batch = _fixture.BatchToken;
        var sheetName = _fixture.CreateTestSheet(batch);
        RequireSuccess(_commands.SetStyle(batch, sheetName, "A1", "Heading 1"));

        AssertAppliedStyle(sheetName, "A1", "Heading 1");
        AssertAppliedStyle(sheetName, "B1", "Normal");
    }

    [Fact]
    public void SetStyle_GoodBadNeutral_AllApplySuccessfully()
    {
        // Arrange & Act & Assert
        var batch = _fixture.BatchToken;
        var sheetName = _fixture.CreateTestSheet(batch);

        RequireSuccess(_commands.SetStyle(batch, sheetName, "A1", "Good"));
        RequireSuccess(_commands.SetStyle(batch, sheetName, "A2", "Bad"));
        RequireSuccess(_commands.SetStyle(batch, sheetName, "A3", "Neutral"));
        AssertAppliedStyle(sheetName, "A1", "Good");
        AssertAppliedStyle(sheetName, "A2", "Bad");
        AssertAppliedStyle(sheetName, "A3", "Neutral");
        AssertAppliedStyle(sheetName, "B1", "Normal");
    }

    [Fact]
    public void SetStyle_Accent1_AppliesSuccessfully()
    {
        // Arrange & Act
        var batch = _fixture.BatchToken;
        var sheetName = _fixture.CreateTestSheet(batch);
        RequireSuccess(_commands.SetStyle(batch, sheetName, "A1:E1", "Accent1"));

        foreach (var column in new[] { "A", "B", "C", "D", "E" })
        {
            AssertAppliedStyle(sheetName, $"{column}1", "Accent1");
        }
        AssertAppliedStyle(sheetName, "F1", "Normal");
    }

    [Fact]
    public void SetStyle_TotalStyle_AppliesSuccessfully()
    {
        // Arrange & Act
        var batch = _fixture.BatchToken;
        var sheetName = _fixture.CreateTestSheet(batch);
        RequireSuccess(_commands.SetStyle(batch, sheetName, "A10:E10", "Total"));

        foreach (var column in new[] { "A", "B", "C", "D", "E" })
        {
            AssertAppliedStyle(sheetName, $"{column}10", "Total");
        }
        AssertAppliedStyle(sheetName, "A9", "Normal");
    }

    [Fact]
    public void SetStyle_CurrencyComma_AppliesSuccessfully()
    {
        // Arrange & Act & Assert
        var batch = _fixture.BatchToken;
        var sheetName = _fixture.CreateTestSheet(batch);

        RequireSuccess(_commands.SetStyle(batch, sheetName, "B5:B10", "Currency"));
        RequireSuccess(_commands.SetStyle(batch, sheetName, "C5:C10", "Comma"));
        for (var row = 5; row <= 10; row++)
        {
            AssertAppliedStyle(sheetName, $"B{row}", "Currency");
            AssertAppliedStyle(sheetName, $"C{row}", "Comma");
        }
        AssertAppliedStyle(sheetName, "D5", "Normal");
    }

    [Fact]
    public void SetStyle_InvalidStyleName_ThrowsException()
    {
        // Arrange & Act & Assert - Should throw when Excel COM rejects invalid style name
        var batch = _fixture.BatchToken;
        var sheetName = _fixture.CreateTestSheet(batch);
        RequireSuccess(_commands.SetValues(batch, sheetName, "A1:B1", [["guard", 7]]));
        RequireSuccess(_commands.SetFormulas(batch, sheetName, "C1", [["=B1*2"]]));
        RequireSuccess(_commands.SetStyle(batch, sheetName, "A1", "Good"));
        var before = RequireSuccess(_commands.GetStyle(batch, sheetName, "A1"));
        var beforeCells = CaptureStyleGuard(batch, sheetName);
        var exception = Assert.Throws<System.Reflection.TargetParameterCountException>(
            () => _commands.SetStyle(batch, sheetName, "A1", "NonExistentStyle"));

        Assert.Contains("Style", exception.Message);
        AssertAppliedStyle(sheetName, "A1", before.StyleName);
        AssertAppliedStyle(sheetName, "B1", "Normal");
        Assert.Equal(beforeCells, CaptureStyleGuard(batch, sheetName));
    }

    [Fact]
    public void SetStyle_ResetToNormal_ClearsFormatting()
    {
        // Arrange & Act
        var batch = _fixture.BatchToken;
        var sheetName = _fixture.CreateTestSheet(batch);

        // Apply fancy style
        RequireSuccess(_commands.SetStyle(batch, sheetName, "A1", "Accent1"));
        AssertAppliedStyle(sheetName, "A1", "Accent1");

        // Reset to normal
        RequireSuccess(_commands.SetStyle(batch, sheetName, "A1", "Normal"));
        AssertAppliedStyle(sheetName, "A1", "Normal");
    }

    /// <summary>
    /// Regression test: FormatRange with verticalAlignment='middle' must succeed
    /// (treated as alias for 'center').
    /// </summary>
    [Fact]
    public void FormatRange_VerticalAlignmentMiddle_AcceptedAsAlias()
    {
        // Arrange
        var batch = _fixture.BatchToken;
        var sheetName = _fixture.CreateTestSheet(batch);
        RequireSuccess(_commands.Format(
            batch, sheetName, ["D1"], new() { VerticalAlignment = "top" }));

        // Act - 'middle' is a common alias for 'center'
        var result = _commands.Format(
            batch, sheetName, ["A1:C3"], new() { VerticalAlignment = "middle" });

        // Assert
        RequireSuccess(result);
        var actual = _fixture.ExecuteRawVerification((ctx, ct) =>
        {
            Excel.Sheets? sheets = null;
            Excel.Worksheet? sheet = null;
            Excel.Range? range = null;
            Excel.Range? neighbor = null;
            try
            {
                sheets = ctx.Book.Worksheets;
                sheet = (Excel.Worksheet)sheets[sheetName];
                range = sheet.Range["A1:C3"];
                neighbor = sheet.Range["D1"];
                return (
                    Target: Convert.ToInt32(range.VerticalAlignment, CultureInfo.InvariantCulture),
                    Neighbor: Convert.ToInt32(neighbor.VerticalAlignment, CultureInfo.InvariantCulture));
            }
            finally
            {
                ComUtilities.Release(ref range);
                ComUtilities.Release(ref neighbor);
                ComUtilities.Release(ref sheet);
                ComUtilities.Release(ref sheets);
            }
        });
        Assert.Equal((int)Excel.XlVAlign.xlVAlignCenter, actual.Target);
        Assert.Equal((int)Excel.XlVAlign.xlVAlignTop, actual.Neighbor);
    }

    private void AssertAppliedStyle(string sheetName, string cell, string expected)
    {
        var result = RequireSuccess(_commands.GetStyle(_fixture.BatchToken, sheetName, cell));
        Assert.Equal(expected, result.StyleName);
    }

    private string CaptureStyleGuard(IExcelBatch batch, string sheetName)
    {
        var values = RequireSuccess(_commands.GetValues(batch, sheetName, "A1:C1"));
        var formulas = RequireSuccess(_commands.GetFormulas(batch, sheetName, "A1:C1"));
        return System.Text.Json.JsonSerializer.Serialize(new
        {
            values.Values,
            formulas.Formulas
        });
    }
}
