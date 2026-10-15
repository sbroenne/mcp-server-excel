using Sbroenne.ExcelMcp.ComInterop;
using Xunit;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sbroenne.ExcelMcp.Service.Tests;

public sealed partial class PersistentServiceRangeValidationTests
{
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void GetValidation_NeverValidatedCells_ReportAbsence(bool neighboringRule)
    {
        var batch = _fixture.BatchToken;
        var sheet = _fixture.CreateTestSheet(batch);
        if (neighboringRule)
            RequireSuccess(_commands.ValidateRange(batch, sheet, "D1", "whole", "between", "1", "10",
                null, null, null, null, null, null, null, null, null));

        AssertNativeTypeRejectsAbsentRule(sheet, "A1:A3");
        var result = RequireSuccess(_commands.GetValidation(batch, sheet, "A1:A3"));

        Assert.False(result.HasValidation);
        Assert.Equal(sheet, result.SheetName);
        Assert.Equal("A1:A3", result.RangeAddress);
        Assert.Null(result.ValidationType);
        Assert.Null(result.Formula1);
        Assert.Null(result.Formula2);
        Assert.True(string.IsNullOrEmpty(result.ErrorMessage));
        Assert.All(RequireSuccess(_commands.GetValues(batch, sheet, "A1:A3")).Values,
            row => Assert.Null(Assert.Single(row)));
        if (neighboringRule)
        {
            var neighbor = RequireSuccess(_commands.GetValidation(batch, sheet, "D1"));
            Assert.True(neighbor.HasValidation);
            Assert.Equal("whole", neighbor.ValidationType);
            Assert.Equal("1", neighbor.Formula1);
            Assert.Equal("10", neighbor.Formula2);
        }
    }

    [Fact]
    public void GetValidation_AfterRemoval_ReportsAbsenceAndPreservesNeighbor()
    {
        var batch = _fixture.BatchToken;
        var sheet = _fixture.CreateTestSheet(batch);
        RequireSuccess(_commands.SetValues(batch, sheet, "A1:B1", [[5, 6]]));
        RequireSuccess(_commands.ValidateRange(batch, sheet, "A1:B1", "whole", "between", "1", "10",
            null, null, null, null, null, null, null, null, null));
        var before = RequireSuccess(_commands.GetValidation(batch, sheet, "A1"));
        Assert.True(before.HasValidation);
        Assert.Equal("whole", before.ValidationType);
        Assert.Equal("1", before.Formula1);
        Assert.Equal("10", before.Formula2);

        RequireSuccess(_commands.RemoveValidation(batch, sheet, "A1"));
        AssertNativeTypeRejectsAbsentRule(sheet, "A1");
        var result = RequireSuccess(_commands.GetValidation(batch, sheet, "A1"));

        Assert.False(result.HasValidation);
        Assert.Equal(sheet, result.SheetName);
        Assert.Equal("A1", result.RangeAddress);
        Assert.Null(result.ValidationType);
        Assert.Null(result.Formula1);
        Assert.Null(result.Formula2);
        Assert.True(string.IsNullOrEmpty(result.ErrorMessage));
        var neighbor = RequireSuccess(_commands.GetValidation(batch, sheet, "B1"));
        Assert.True(neighbor.HasValidation);
        Assert.Equal("whole", neighbor.ValidationType);
        Assert.Equal("1", neighbor.Formula1);
        Assert.Equal("10", neighbor.Formula2);
        PowerQueryStateAssertions.AssertRows([[5, 6]],
            RequireSuccess(_commands.GetValues(batch, sheet, "A1:B1")).Values);
    }

    [Fact]
    public void GetValidation_InvalidNativeRange_PreservesNativeFailure()
    {
        var batch = _fixture.BatchToken;
        var sheet = _fixture.CreateTestSheet(batch);
        var failure = Assert.Throws<InvalidOperationException>(() =>
            _commands.GetValidation(batch, sheet, "XFE1048577"));
        Assert.Contains("ComInterop/COMException", failure.Message, StringComparison.Ordinal);
        Assert.Contains("0x800A03EC", failure.Message, StringComparison.Ordinal);
    }

    private void AssertNativeTypeRejectsAbsentRule(string sheetName, string address) =>
        _fixture.ExecuteRawVerification((context, _) =>
        {
            Excel.Sheets? sheets = null;
            Excel.Worksheet? sheet = null;
            Excel.Range? range = null;
            Excel.Validation? validation = null;
            try
            {
                sheets = context.Book.Worksheets;
                sheet = (Excel.Worksheet)sheets[sheetName];
                range = sheet.Range[address];
                validation = range.Validation;
                var error = Assert.Throws<System.Runtime.InteropServices.COMException>(() => validation.Type);
                Assert.Equal(unchecked((int)0x800A03EC), error.HResult);
            }
            finally
            {
                ComUtilities.Release(ref validation);
                ComUtilities.Release(ref range);
                ComUtilities.Release(ref sheet);
                ComUtilities.Release(ref sheets);
            }
        });
}
