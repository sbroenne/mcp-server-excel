using Sbroenne.ExcelMcp.Core.Models;
using Sbroenne.ExcelMcp.Core.Utilities;
using Xunit;

namespace Sbroenne.ExcelMcp.Portable.Tests;

[Trait("RequiresExcel", "false")]
public sealed class SharedRangeSemanticsTests
{
    [Theory]
    [InlineData(0)]
    [InlineData(1)]
    public void FormulaMatricesPreserveBoundsErrorsAndNonformulaCells(int lowerBound)
    {
        var formulas = (object[,])Array.CreateInstance(typeof(object), [2, 2], [lowerBound, lowerBound]);
        var values = (object[,])Array.CreateInstance(typeof(object), [2, 2], [lowerBound, lowerBound]);
        formulas[lowerBound, lowerBound] = "=1/0";
        formulas[lowerBound, lowerBound + 1] = 2007d;
        formulas[lowerBound + 1, lowerBound] = "=\"\"";
        values[lowerBound, lowerBound] = -2146826281;
        values[lowerBound, lowerBound + 1] = 2007d;
        values[lowerBound + 1, lowerBound] = "";
        var result = RangeFormulaResults.Create("owned.xlsx", "Sheet1", "$AA$3:$AB$4", 3, 27, formulas, values);
        Assert.True(result.Success);
        Assert.True(string.IsNullOrEmpty(result.ErrorMessage));
        Assert.Equal(2, result.RowCount);
        Assert.Equal(2, result.ColumnCount);
        Assert.Equal(["=1/0", ""], result.Formulas[0]);
        Assert.Equal(["=\"\"", ""], result.Formulas[1]);
        Assert.Equal("#DIV/0!", result.Values[0][0]);
        Assert.Equal(2007d, result.Values[0][1]);
        Assert.Equal("", result.Values[1][0]);
        Assert.Null(result.Values[1][1]);
        var error = Assert.Single(result.CellErrors);
        Assert.Equal("AA3", error.CellAddress);
        Assert.Equal(3, error.Row);
        Assert.Equal(27, error.Column);
        Assert.Equal("=1/0", error.Formula);
        Assert.Equal(-2146826281, error.ErrorCode);
        Assert.Equal(-2146826281, error.CurrentValue);
        Assert.Equal("#DIV/0! - Division by zero", error.ErrorMessage);
        Assert.Equal("Ensure the formula does not divide by zero.", error.Suggestion);
    }

    [Fact]
    public void SingleFormulaResultIsAlwaysOneByOne()
    {
        var result = RangeFormulaResults.Create("owned.xlsx", "Sheet1", "$B$3", 3, 2, 42d, 42d);
        Assert.Equal("", Assert.Single(Assert.Single(result.Formulas)));
        Assert.Equal(42d, Assert.Single(Assert.Single(result.Values)));
        Assert.Empty(result.CellErrors);
    }

    [Fact]
    public void DimensionValidationPreservesColumnBeforeRowDiagnosticAndParameter()
    {
        var error = Assert.Throws<ArgumentException>(() =>
            RangeCommandValidation.ValidateDimensions<string>([["=1"]], 2, 2, "formulas", "Formula"));
        Assert.Equal("formulas", error.ParamName);
        Assert.StartsWith("Formula array row 1 column count (1) doesn't match range column count (2)", error.Message);
        error = Assert.Throws<ArgumentException>(() =>
            RangeCommandValidation.ValidateDimensions<string>([["=1", "=2"]], 2, 2, "formulas", "Formula"));
        Assert.StartsWith("Formula array row count (1) doesn't match range row count (2).", error.Message);
    }

    [Theory]
    [InlineData(null, null, false)]
    [InlineData("", null, true)]
    [InlineData(null, "=\"\"", true)]
    [InlineData(" ", null, true)]
    [InlineData(0, null, true)]
    public void OccupancyIncludesEmptyValuedFormulasAndNonNullConstants(object? value, object? formula, bool expected) =>
        Assert.Equal(expected, RangeCommandValidation.IsOccupied(value, formula));
}
