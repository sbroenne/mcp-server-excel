using System.Globalization;
using System.Text.Json;
using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.Core.Models;
using Xunit;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sbroenne.ExcelMcp.Service.Tests;

public sealed partial class PersistentServicePivotTableTests
{
    private T ReadNativePivot<T>(string sheetName, string pivotName, Func<Excel.PivotTable, T> read) =>
        _fixture.ExecuteRawVerification((context, _) =>
        {
            Excel.Worksheet? sheet = null;
            Excel.PivotTables? pivots = null;
            Excel.PivotTable? pivot = null;
            try
            {
                sheet = ComUtilities.FindSheet(context.Book, sheetName);
                Assert.NotNull(sheet);
                pivots = (Excel.PivotTables)sheet.PivotTables();
                pivot = pivots.Item(pivotName);
                return read(pivot);
            }
            finally
            {
                ComUtilities.Release(ref pivot);
                ComUtilities.Release(ref pivots);
                ComUtilities.Release(ref sheet);
            }
        });

    private void AssertCreatedPivot(PivotTableCreateResult result, string sheetName, string anchor)
    {
        RequireSuccess(result);
        Assert.Equal(sheetName, result.SheetName);
        Assert.Equal(5, result.SourceRowCount);
        Assert.Equal(["Date", "Product", "Region", "Sales"],
            result.AvailableFields.Order(StringComparer.Ordinal));
        ReadNativePivot(sheetName, result.PivotTableName, pivot =>
        {
            Excel.Range? range = null;
            Excel.Range? firstCell = null;
            Excel.Range? cells = null;
            Excel.PivotCache? cache = null;
            try
            {
                Assert.Equal(result.PivotTableName, pivot.Name);
                range = pivot.TableRange2;
                Assert.Equal(result.Range, range.Address);
                cells = range.Cells;
                firstCell = (Excel.Range)cells[1, 1];
                Assert.Equal(anchor, firstCell.Address[false, false]);
                cache = pivot.PivotCache();
                Assert.Equal(5, cache.RecordCount);
                Assert.False(cache.OLAP);
                Assert.Equal(Excel.XlPivotTableSourceType.xlDatabase, cache.SourceType);
                return 0;
            }
            finally
            {
                ComUtilities.Release(ref cache);
                ComUtilities.Release(ref firstCell);
                ComUtilities.Release(ref cells);
                ComUtilities.Release(ref range);
            }
        });
        RequireSuccess(_pivotCommands.AddRowField(_fixture.BatchToken, result.PivotTableName, "Region"));
        RequireSuccess(_pivotCommands.AddValueField(_fixture.BatchToken, result.PivotTableName, "Sales"));
        RequireSuccess(_pivotCommands.Refresh(_fixture.BatchToken, result.PivotTableName));
        AssertPivotSales(325, 325, result.PivotTableName);
        AssertOriginalSales();
    }

    private void AssertNativeField(
        string pivotName, string fieldName, PivotFieldArea area,
        string? sheetName = null, AggregationFunction? function = null, string? caption = null)
    {
        ReadNativePivot(sheetName ?? _salesSheetName, pivotName, pivot =>
        {
            Excel.PivotField? field = null;
            Excel.PivotFields? dataFields = null;
            try
            {
                if (area == PivotFieldArea.Value)
                {
                    dataFields = (Excel.PivotFields)pivot.DataFields;
                    Assert.Equal(1, dataFields.Count);
                    field = dataFields.Item(1);
                    Assert.Equal(fieldName, field.SourceName);
                }
                else
                {
                    field = (Excel.PivotField)pivot.PivotFields(caption ?? fieldName);
                    Assert.Equal(fieldName, field.SourceName);
                }
                Assert.Equal((int)area, (int)field.Orientation);
                if (area != PivotFieldArea.Hidden)
                {
                    Assert.Equal(1, field.Position);
                }
                if (caption is not null)
                {
                    Assert.Equal(caption, field.Caption);
                }
                if (function is not null)
                {
                    var expected = function switch
                    {
                        AggregationFunction.Sum => Excel.XlConsolidationFunction.xlSum,
                        AggregationFunction.Average => Excel.XlConsolidationFunction.xlAverage,
                        _ => throw new ArgumentOutOfRangeException(nameof(function))
                    };
                    Assert.Equal(expected, field.Function);
                }
                return 0;
            }
            finally
            {
                ComUtilities.Release(ref field);
                ComUtilities.Release(ref dataFields);
            }
        });
    }

    private void AssertOriginalSales()
    {
        var values = RequireSuccess(_commands.GetValues(_fixture.BatchToken, _salesSheetName, "A1:D6")).Values;
        Assert.Equal(6, values.Count);
        Assert.Equal(new object?[] { "Region", "Product", "Sales", "Date" }, values[0]);
        var regions = new[] { "North", "North", "South", "North", "South" };
        var products = new[] { "Widget", "Widget", "Gadget", "Gadget", "Widget" };
        var sales = new[] { 100d, 150d, 200d, 75d, 125d };
        var dates = new[]
        {
            new DateTime(2025, 1, 15), new DateTime(2025, 1, 20),
            new DateTime(2025, 2, 10), new DateTime(2025, 2, 15), new DateTime(2025, 3, 5)
        };
        for (var index = 0; index < sales.Length; index++)
        {
            var row = values[index + 1];
            Assert.Equal(4, row.Count);
            Assert.Equal(regions[index], row[0]);
            Assert.Equal(products[index], row[1]);
            Assert.Equal(sales[index], Convert.ToDouble(row[2], CultureInfo.InvariantCulture));
            Assert.Equal(dates[index], DateTime.FromOADate(Convert.ToDouble(row[3], CultureInfo.InvariantCulture)));
        }
    }

    private string SnapshotPivot(string pivotName = "TestPivot") =>
        JsonSerializer.Serialize(RequireSuccess(_pivotCommands.Read(_fixture.BatchToken, pivotName)));

    private void AssertPivotNumberFormat(string expectedFormat, double expectedValue, string expectedText)
    {
        RequireSuccess(_pivotCommands.Refresh(_fixture.BatchToken, "TestPivot"));
        ReadNativePivot(_salesSheetName, "TestPivot", pivot =>
        {
            Excel.Range? body = null;
            Excel.Range? cells = null;
            Excel.Range? cell = null;
            try
            {
                body = pivot.DataBodyRange;
                Assert.NotNull(body);
                cells = body.Cells;
                cell = (Excel.Range)cells[1, 1];
                string format = Convert.ToString(cell.NumberFormat, CultureInfo.InvariantCulture) ?? string.Empty;
                Assert.Equal(expectedFormat, format.Replace("\\$", "$", StringComparison.Ordinal));
                Assert.Equal(expectedValue, Convert.ToDouble(cell.Value2, CultureInfo.InvariantCulture));
                Assert.Equal(expectedText, Convert.ToString(cell.Text, CultureInfo.InvariantCulture));
                return 0;
            }
            finally
            {
                ComUtilities.Release(ref cell);
                ComUtilities.Release(ref cells);
                ComUtilities.Release(ref body);
            }
        });
        AssertOriginalSales();
    }

    private void AssertNativeLayout(string sheetName, string pivotName, int layout)
    {
        ReadNativePivot(sheetName, pivotName, pivot =>
        {
            Excel.PivotFields? rows = null;
            try
            {
                rows = (Excel.PivotFields)pivot.RowFields;
                Assert.True(rows.Count > 0);
                for (var index = 1; index <= rows.Count; index++)
                {
                    Excel.PivotField? field = null;
                    try
                    {
                        field = rows.Item(index);
                        Assert.Equal(layout == 0, field.LayoutCompactRow);
                        Assert.Equal(layout == 1 ? Excel.XlLayoutFormType.xlTabular : Excel.XlLayoutFormType.xlOutline,
                            field.LayoutForm);
                    }
                    finally
                    {
                        ComUtilities.Release(ref field);
                    }
                }
                return 0;
            }
            finally
            {
                ComUtilities.Release(ref rows);
            }
        });
        AssertOriginalSales();
    }
}
