using System.Globalization;
using System.Runtime.InteropServices;
using Sbroenne.ExcelMcp.ComInterop;
using Xunit;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sbroenne.ExcelMcp.Service.Tests;

public sealed partial class PersistentServiceRangeAdvancedTests
{
    [Fact]
    public void RemoveValidation_RemovesOnlySelectedCellsAndPreservesValues()
    {
        var batch = _fixture.BatchToken;
        string sheet = _fixture.CreateTestSheet(batch);
        RequireSuccess(_commands.SetValues(batch, sheet, "A1:B2", [[1, 2], [3, 4]]));
        WithNativeRange(sheet, "A1:B2", range =>
        {
            Excel.Validation? validation = null;
            try
            {
                validation = range.Validation;
                validation.Add(Excel.XlDVType.xlValidateWholeNumber,
                    Excel.XlDVAlertStyle.xlValidAlertStop, Excel.XlFormatConditionOperator.xlBetween, "1", "10");
            }
            finally { ComUtilities.Release(ref validation); }
        });
        Assert.True(RequireSuccess(_commands.GetValidation(batch, sheet, "A1")).HasValidation);
        RequireSuccess(_commands.RemoveValidation(batch, sheet, "A1:A2"));
        foreach (string cell in new[] { "A1", "A2" })
        {
            Assert.False(RequireSuccess(_commands.GetValidation(batch, sheet, cell)).HasValidation);
            WithNativeRange(sheet, cell, range =>
            {
                Excel.Validation? validation = null;
                try
                {
                    validation = range.Validation;
                    var selectedValidation = validation ??
                        throw new InvalidOperationException("Excel did not return a validation object.");
                    Assert.Throws<COMException>(() => selectedValidation.Type);
                }
                finally { ComUtilities.Release(ref validation); }
            });
        }
        foreach (string cell in new[] { "B1", "B2" })
        {
            var retained = RequireSuccess(_commands.GetValidation(batch, sheet, cell));
            Assert.True(retained.HasValidation);
            Assert.Equal("1", retained.Formula1);
            Assert.Equal("10", retained.Formula2);
        }
        Assert.Equal([1d, 2d, 3d, 4d],
            RequireSuccess(_commands.GetValues(batch, sheet, "A1:B2")).Values
                .SelectMany(row => row).Select(value => Convert.ToDouble(value, CultureInfo.InvariantCulture)));
    }

    [Fact]
    public void AutoFitColumns_FitsContentWithoutChangingAdjacentColumnOrRows()
    {
        var batch = _fixture.BatchToken;
        string sheet = _fixture.CreateTestSheet(batch);
        const string text = "A deliberately long heading that cannot fit inside a narrow column";
        RequireSuccess(_commands.SetValues(batch, sheet, "A1:B1", [[text, "neighbor"]]));
        RequireSuccess(_commands.SetValues(batch, sheet, "D1", [[text]]));
        WithNativeRange(sheet, "A1:B1", range => range.ColumnWidth = 3);
        var before = PersistentServiceRangeVerification.ReadGeometry(_fixture, sheet, "A1");
        var neighbor = PersistentServiceRangeVerification.ReadGeometry(_fixture, sheet, "B1");
        RequireSuccess(_commands.AutoFitColumns(batch, sheet, "A1:A2"));
        WithNativeRange(sheet, "D1:D2", range =>
        {
            Excel.Range? columns = null;
            try
            {
                columns = range.Columns;
                columns.AutoFit();
            }
            finally { ComUtilities.Release(ref columns); }
        });
        var expected = PersistentServiceRangeVerification.ReadGeometry(_fixture, sheet, "D1");
        var after = PersistentServiceRangeVerification.ReadGeometry(_fixture, sheet, "A1");
        var retained = PersistentServiceRangeVerification.ReadGeometry(_fixture, sheet, "B1");
        Assert.True(after.Width > before.Width, $"Width did not increase: {before.Width} -> {after.Width}");
        Assert.Equal(expected.Width, after.Width);
        Assert.Equal(before.Height, after.Height);
        Assert.Equal(neighbor.Width, retained.Width);
        Assert.Equal(neighbor.Height, retained.Height);
        Assert.Equal(new[] { text, "neighbor" },
            RequireSuccess(_commands.GetValues(batch, sheet, "A1:B1")).Values[0].Select(value => value?.ToString()));
    }

    [Fact]
    public void AutoFitRows_FitsWrappedContentWithoutChangingAdjacentRowOrColumns()
    {
        var batch = _fixture.BatchToken;
        string sheet = _fixture.CreateTestSheet(batch);
        const string text = "First line\nSecond line\nThird line";
        RequireSuccess(_commands.SetValues(batch, sheet, "A1:A2", [[text], ["neighbor"]]));
        RequireSuccess(_commands.SetValues(batch, sheet, "A4", [[text]]));
        WithNativeRange(sheet, "A4", range => range.WrapText = true);
        WithNativeRange(sheet, "A1", range =>
        {
            range.WrapText = true;
            range.RowHeight = 8;
        });
        WithNativeRange(sheet, "A2", range => range.RowHeight = 28);
        var before = PersistentServiceRangeVerification.ReadGeometry(_fixture, sheet, "A1");
        var neighbor = PersistentServiceRangeVerification.ReadGeometry(_fixture, sheet, "A2");
        RequireSuccess(_commands.AutoFitRows(batch, sheet, "A1:B1"));
        WithNativeRange(sheet, "A4:B4", range =>
        {
            Excel.Range? rows = null;
            try
            {
                rows = range.Rows;
                rows.AutoFit();
            }
            finally { ComUtilities.Release(ref rows); }
        });
        var expected = PersistentServiceRangeVerification.ReadGeometry(_fixture, sheet, "A4");
        var after = PersistentServiceRangeVerification.ReadGeometry(_fixture, sheet, "A1");
        var retained = PersistentServiceRangeVerification.ReadGeometry(_fixture, sheet, "A2");
        Assert.True(after.Height > before.Height, $"Height did not increase: {before.Height} -> {after.Height}");
        Assert.Equal(expected.Height, after.Height);
        Assert.Equal(before.Width, after.Width);
        Assert.Equal(neighbor.Height, retained.Height);
        Assert.Equal(neighbor.Width, retained.Width);
        Assert.Equal(new[] { text, "neighbor" },
            RequireSuccess(_commands.GetValues(batch, sheet, "A1:A2")).Values.Select(row => row[0]?.ToString()));
    }

    [Fact]
    public void UnmergeCells_RetainsTopLeftPayloadAndUnrelatedMerge()
    {
        var batch = _fixture.BatchToken;
        string sheet = _fixture.CreateTestSheet(batch);
        RequireSuccess(_commands.SetValues(batch, sheet, "A1", [["payload"]]));
        RequireSuccess(_commands.SetValues(batch, sheet, "D1", [["neighbor"]]));
        WithNativeRange(sheet, "A1:B2", range => range.Merge());
        WithNativeRange(sheet, "D1:E1", range => range.Merge());
        WithNativeRange(sheet, "A1:B2", range => Assert.True(Convert.ToBoolean(((object)range.MergeCells), CultureInfo.InvariantCulture)));
        RequireSuccess(_commands.UnmergeCells(batch, sheet, "A1:B2"));
        WithNativeRange(sheet, "A1:B2", range => Assert.False(Convert.ToBoolean(((object)range.MergeCells), CultureInfo.InvariantCulture)));
        WithNativeRange(sheet, "D1:E1", range => Assert.True(Convert.ToBoolean(((object)range.MergeCells), CultureInfo.InvariantCulture)));
        var values = RequireSuccess(_commands.GetValues(batch, sheet, "A1:B2")).Values;
        Assert.Equal("payload", values[0][0]?.ToString());
        Assert.Null(values[0][1]);
        Assert.All(values[1], Assert.Null);
        Assert.Equal("neighbor", RequireSuccess(_commands.GetValues(batch, sheet, "D1")).Values[0][0]?.ToString());
    }

    private void WithNativeRange(string sheetName, string address, Action<Excel.Range> inspect) =>
        _fixture.ExecuteRawVerification((context, _) =>
        {
            Excel.Worksheet? sheet = null;
            Excel.Range? range = null;
            try
            {
                sheet = ComUtilities.FindSheet(context.Book, sheetName) ??
                    throw new InvalidOperationException($"Excel did not return worksheet '{sheetName}'.");
                range = sheet.Range[address];
                inspect(range ?? throw new InvalidOperationException("Excel did not return the requested range."));
            }
            finally
            {
                ComUtilities.Release(ref range);
                ComUtilities.Release(ref sheet);
            }
        });
}
