using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.Core.Commands.Table;
using Xunit;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sbroenne.ExcelMcp.Service.Tests;

/// <summary>
/// Tests for range formulas operations
/// </summary>
public sealed partial class PersistentServiceRangeSpecializedTests
{
    // === FORMULA OPERATIONS TESTS ===

    // === TABLE FORMULA COMPATIBILITY TESTS ===

    [Fact]
    public void SetFormulas_InExcelTable_UsesSessionFormulaSemantics()
    {
        // Regression test: Range.Formula (legacy) injects @ implicit intersection operator
        // inside Excel Tables, causing #FIELD! errors with custom functions that return
        // entity cards. Range.Formula2 (modern) respects dynamic array semantics.

        // Arrange - create a sheet with data and an Excel Table
        var batch = _fixture.BatchToken;
        var sheetName = _fixture.CreateTestSheet(batch);
        bool supportsFormula2 = _fixture.ExecuteRawVerification(
            (ctx, ct) => ctx.Capabilities.SupportsFormula2);

        // Set up data first
        _commands.SetValues(batch, sheetName, "A1:C4",
        [
            ["Name", "Value", "Doubled"],
            ["Alpha", 10, null],
            ["Beta", 20, null],
            ["Gamma", 30, null]
        ]);

        // Create an Excel Table over the data
        var tableCommands = _fixture.CreateCommands<ITableCommands>();
        var tableResult = tableCommands.Create(batch, sheetName, "Formula2TestTable", "A1:C4");
        Assert.True(tableResult.Success);

        // Act - set formulas INSIDE the table (column C, within table range)
        var setResult = _commands.SetFormulas(batch, sheetName, "C2:C4",
        [
            ["=B2*2"],
            ["=B3*2"],
            ["=B4*2"]
        ]);

        // Assert - formulas should be set successfully
        Assert.True(setResult.Success);

        var readResult = _commands.GetFormulas(batch, sheetName, "C2:C4");
        Assert.True(readResult.Success);
        Assert.Equal(3, readResult.Formulas.Count);
        var legacyFormulas = supportsFormula2
        ? null
        : ReadLegacyTableFormulas(sheetName, "C2:C4");

        for (int i = 0; i < readResult.Formulas.Count; i++)
        {
            var formula = readResult.Formulas[i][0];
            _output.WriteLine($"C{i + 2} formula: {formula}");

            if (supportsFormula2)
            {
                Assert.DoesNotContain("@", formula);
                Assert.Equal($"=B{i + 2}*2", formula);
            }
            else
            {
                Assert.Equal(legacyFormulas![i + 1, 1], formula);
            }
        }

        // Verify calculated values are correct
        Assert.Equal(20.0, Convert.ToDouble(readResult.Values[0][0], System.Globalization.CultureInfo.InvariantCulture));
        Assert.Equal(40.0, Convert.ToDouble(readResult.Values[1][0], System.Globalization.CultureInfo.InvariantCulture));
        Assert.Equal(60.0, Convert.ToDouble(readResult.Values[2][0], System.Globalization.CultureInfo.InvariantCulture));
    }

    [Fact]
    public void GetFormulas_InExcelTable_UsesSessionFormulaSemantics()
    {
        // Arrange - create a sheet with data, table, and formulas
        var batch = _fixture.BatchToken;
        var sheetName = _fixture.CreateTestSheet(batch);
        bool supportsFormula2 = _fixture.ExecuteRawVerification(
            (ctx, ct) => ctx.Capabilities.SupportsFormula2);

        _commands.SetValues(batch, sheetName, "A1:C4",
        [
            ["X", "Y", "Sum"],
            [1, 2, null],
            [3, 4, null],
            [5, 6, null]
        ]);

        // Set formulas before creating the table, using the session's selected API.
        _commands.SetFormulas(batch, sheetName, "C2:C4",
        [
            ["=A2+B2"],
            ["=A3+B3"],
            ["=A4+B4"]
        ]);

        // Create Excel Table around the data including formula column
        var tableCommands = _fixture.CreateCommands<ITableCommands>();
        var tableResult = tableCommands.Create(batch, sheetName, "GetFormula2TestTable", "A1:C4");
        Assert.True(tableResult.Success);

        // Act - read formulas back from inside the table
        var readResult = _commands.GetFormulas(batch, sheetName, "C2:C4");

        Assert.True(readResult.Success);
        Assert.Equal(3, readResult.Formulas.Count);
        var legacyFormulas = supportsFormula2
            ? null
            : ReadLegacyTableFormulas(sheetName, "C2:C4");

        for (int i = 0; i < readResult.Formulas.Count; i++)
        {
            var formula = readResult.Formulas[i][0];
            _output.WriteLine($"C{i + 2} formula: {formula}");

            if (supportsFormula2)
            {
                Assert.DoesNotContain("@", formula);
                Assert.Equal($"=A{i + 2}+B{i + 2}", formula);
            }
            else
            {
                Assert.Equal(legacyFormulas![i + 1, 1], formula);
            }
        }

        // Verify calculated values
        Assert.Equal(3.0, Convert.ToDouble(readResult.Values[0][0], System.Globalization.CultureInfo.InvariantCulture));
        Assert.Equal(7.0, Convert.ToDouble(readResult.Values[1][0], System.Globalization.CultureInfo.InvariantCulture));
        Assert.Equal(11.0, Convert.ToDouble(readResult.Values[2][0], System.Globalization.CultureInfo.InvariantCulture));
    }

    private object[,] ReadLegacyTableFormulas(string sheetName, string address) =>
        _fixture.ExecuteRawVerification((ctx, ct) =>
        {
            Excel.Worksheet? sheet = null;
            Excel.Range? range = null;
            try
            {
                sheet = ComUtilities.FindSheet(ctx.Book, sheetName);
                Assert.NotNull(sheet);
                range = sheet.Range[address];
                return Assert.IsType<object[,]>(range.Formula);
            }
            finally
            {
                ComUtilities.Release(ref range);
                ComUtilities.Release(ref sheet);
            }
        });
}


