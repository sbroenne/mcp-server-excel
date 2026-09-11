using System.Globalization;
using System.Runtime.InteropServices;
using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Commands.Table;
using Xunit;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sbroenne.ExcelMcp.Core.Tests.Commands.Range;

public partial class RangeCommandsTests
{
    [Fact]
    [Trait("Layer", "Core")]
    public void FormulaCompatibility_TableArrayArgument_UsesSessionSemanticsAndLegacyControl()
    {
        using var batch = ExcelSession.BeginBatch(_fixture.CreateTestFile());
        var sheetName = _fixture.CreateTestSheet(batch);
        bool supportsFormula2 = batch.Execute((ctx, ct) => ctx.Capabilities.SupportsFormula2);
        const string formula = "=SUM(SQRT($A$2:$A$3))";
        _commands.SetValues(batch, sheetName, "A1:C3",
        [
            ["Input", "Session", "Legacy control"],
            [4, null, null],
            [9, null, null]
        ]);
        Assert.True(new TableCommands().Create(batch, sheetName, "ArrayArgumentTable", "A1:C3").Success);
        Assert.True(_commands.SetFormulas(batch, sheetName, "B2:B3", [[formula], [formula]]).Success);

        // Exercise the actual legacy API on the installed Excel, not an old-Excel simulation.
        batch.Execute((ctx, ct) =>
        {
            Excel.Worksheet? sheet = null;
            Excel.Range? range = null;
            try
            {
                sheet = ComUtilities.FindSheet(ctx.Book, sheetName);
                Assert.NotNull(sheet);
                range = sheet.Range["C2:C3"];
                range.Formula = formula;
            }
            finally
            {
                ComUtilities.Release(ref range);
                ComUtilities.Release(ref sheet);
            }
        });

        var legacyValues = _commands.GetValues(batch, sheetName, "C2:C3");
        Assert.Equal(2.0, Convert.ToDouble(legacyValues.Values[0][0], CultureInfo.InvariantCulture));
        Assert.Equal(3.0, Convert.ToDouble(legacyValues.Values[1][0], CultureInfo.InvariantCulture));
        var result = _commands.GetFormulas(batch, sheetName, "B2:B3");
        Assert.Equal(2, result.Formulas.Count);
        Assert.Empty(result.CellErrors);
        if (supportsFormula2)
        {
            Assert.All(result.Formulas, row =>
            {
                Assert.DoesNotContain("@", row[0]);
                Assert.Equal(formula, row[0]);
            });
            Assert.All(result.Values, row => Assert.Equal(5.0, Convert.ToDouble(row[0], CultureInfo.InvariantCulture)));
        }
        else
        {
            var legacyFormulas = ReadLegacyTableFormulas(batch, sheetName, "C2:C3");
            for (int row = 0; row < 2; row++)
            {
                Assert.Equal(legacyFormulas[row + 1, 1], result.Formulas[row][0]);
                Assert.Equal(legacyValues.Values[row][0], result.Values[row][0]);
            }
        }
    }

    [Fact]
    [Trait("Layer", "Core")]
    public void FormulaCompatibility_LegacySafeFormulas_RoundTripAndEnrichErrors()
    {
        string path = _fixture.CreateTestFile();
        string sheetName;
        using (var batch = ExcelSession.BeginBatch(path))
        {
            sheetName = _fixture.CreateTestSheet(batch);
            _commands.SetValues(batch, sheetName, "A1:A2", [[10], [20]]);
            Assert.Equal(string.Empty, _commands.GetFormulas(batch, sheetName, "B1").Formulas[0][0]);
            var constants = _commands.GetFormulas(batch, sheetName, "A1:A2");
            Assert.All(constants.Formulas, row => Assert.Equal(string.Empty, row[0]));

            Assert.True(_commands.SetFormulas(batch, sheetName, "A3", [["=A1+A2"]]).Success);
            var routed = _commands.SetValues(batch, sheetName, "B1:B2", [["=A1+A2"], ["=1/0"]]);
            Assert.True(routed.Success);
            Assert.True(string.IsNullOrEmpty(routed.ErrorMessage));
            Assert.Contains("set-formulas", routed.Message);

            var single = _commands.GetValues(batch, sheetName, "B2");
            Assert.Equal("#DIV/0!", single.Values[0][0]);
            Assert.Equal("=1/0", Assert.Single(single.CellErrors).Formula);
            var multiple = _commands.GetValues(batch, sheetName, "B1:B2");
            Assert.Equal(30.0, Convert.ToDouble(multiple.Values[0][0], CultureInfo.InvariantCulture));
            Assert.Equal("#DIV/0!", multiple.Values[1][0]);
            var error = Assert.Single(multiple.CellErrors);
            Assert.Equal("B2", error.CellAddress);
            Assert.Equal("=1/0", error.Formula);
            var formulas = _commands.GetFormulas(batch, sheetName, "B1:B2");
            Assert.Equal("=A1+A2", formulas.Formulas[0][0]);
            Assert.Equal("=1/0", formulas.Formulas[1][0]);
            Assert.Equal("#DIV/0!", formulas.Values[1][0]);
            batch.Save();
        }

        using var reopened = ExcelSession.BeginBatch(path);
        var persisted = _commands.GetFormulas(reopened, sheetName, "A3");
        Assert.Equal("=A1+A2", persisted.Formulas[0][0]);
        Assert.Equal(30.0, Convert.ToDouble(persisted.Values[0][0], CultureInfo.InvariantCulture));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    [Trait("Layer", "Core")]
    public void FormulaCompatibility_WriteFailure_DoesNotDowngradeSession(bool protectSheet)
    {
        using var batch = ExcelSession.BeginBatch(_fixture.CreateTestFile());
        var sheetName = _fixture.CreateTestSheet(batch);
        _commands.SetFormulas(batch, sheetName, "A1", [["=42"]]);
        bool supported = batch.Execute((ctx, ct) => ctx.Capabilities.SupportsFormula2);

        void SetProtection(bool protect)
        {
            batch.Execute((ctx, ct) =>
            {
                Excel.Worksheet? sheet = null;
                try
                {
                    sheet = ComUtilities.FindSheet(ctx.Book, sheetName);
                    Assert.NotNull(sheet);
                    if (protect) sheet.Protect();
                    else sheet.Unprotect();
                }
                finally
                {
                    ComUtilities.Release(ref sheet);
                }
            });
        }

        if (protectSheet) SetProtection(true);
        var error = Assert.Throws<COMException>(() =>
            _commands.SetFormulas(batch, sheetName, "A1", [[protectSheet ? "=43" : "=1+"]]));
        Assert.Equal(unchecked((int)0x800A03EC), error.HResult);
        Assert.Equal("=42", _commands.GetFormulas(batch, sheetName, "A1").Formulas[0][0]);
        Assert.Equal(supported, batch.Execute((ctx, ct) => ctx.Capabilities.SupportsFormula2));
        if (protectSheet) SetProtection(false);

        if (supported)
        {
            _commands.SetFormulas(batch, sheetName, "C1", [["=SEQUENCE(2)"]]);
            var spilled = _commands.GetValues(batch, sheetName, "C1:C2");
            Assert.Equal(1.0, Convert.ToDouble(spilled.Values[0][0], CultureInfo.InvariantCulture));
            Assert.Equal(2.0, Convert.ToDouble(spilled.Values[1][0], CultureInfo.InvariantCulture));
        }
    }
}
