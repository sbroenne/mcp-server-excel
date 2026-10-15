using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.Core.Models;
using Xunit;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sbroenne.ExcelMcp.Service.Tests;

public partial class PersistentServicePivotTableOlapFieldTests
{
    [Fact]
    [Trait("Speed", "Medium")]
    public void SetFieldFormat_NativeDataFieldAndCells_ConfirmAppliedFormat()
    {
        var batch = _fixture.BatchToken;
        var (sheetName, pivotName) = CreateIsolatedPivot();
        RequireSuccess(_pivotCommands.AddValueField(
            batch, pivotName, "[RegionalSalesTable].[Sales]", AggregationFunction.Sum, "Native Format"));
        RequireSuccess(_pivotCommands.Refresh(batch, pivotName));
        var formatted = RequireSuccess(_pivotCommands.SetFieldFormat(
            batch, pivotName, "[Measures].[Native Format]", "0%"));
        Assert.Equal("0%", formatted.NumberFormat);
        AssertPivotFieldFormat(sheetName, pivotName, "Native Format", "0%");
        _fixture.ExecuteRawVerification((context, _) =>
        {
            Excel.Worksheet? sheet = null;
            Excel.PivotTables? pivots = null;
            Excel.PivotTable? pivot = null;
            Excel.PivotFields? dataFields = null;
            Excel.PivotField? dataField = null;
            Excel.Range? range = null;
            try
            {
                sheet = ComUtilities.FindSheet(context.Book, sheetName);
                Assert.NotNull(sheet);
                pivots = (Excel.PivotTables)sheet.PivotTables();
                pivot = pivots.Item(pivotName);
                dataFields = (Excel.PivotFields)pivot.DataFields;
                dataField = dataFields.Item(1);
                Assert.Equal("0%", dataField.NumberFormat);
                range = dataField.DataRange;
                Assert.Equal("0%", range.NumberFormat);
                Assert.Equal(49000d, Convert.ToDouble(range.Value2, System.Globalization.CultureInfo.InvariantCulture));
            }
            finally
            {
                ComUtilities.Release(ref range);
                ComUtilities.Release(ref dataField);
                ComUtilities.Release(ref dataFields);
                ComUtilities.Release(ref pivot);
                ComUtilities.Release(ref pivots);
                ComUtilities.Release(ref sheet);
            }
        });
        AssertMeasureValue("Native Format", 49000);
    }
}
