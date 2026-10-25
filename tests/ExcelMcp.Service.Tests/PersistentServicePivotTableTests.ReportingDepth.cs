using System.Globalization;
using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.Core.Commands.PivotTable;
using Sbroenne.ExcelMcp.Core.Commands.Range;
using Sbroenne.ExcelMcp.Core.Models;
using Xunit;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sbroenne.ExcelMcp.Service.Tests;

public sealed partial class PersistentServicePivotTableTests
{
    private string CreateReportingDepthPivot(string row = "Region")
    {
        var name = $"ReportDepth_{Guid.NewGuid():N}";
        var sheet = _fixture.CreateTestSheet(_fixture.BatchToken);
        RequireSuccess(_pivotCommands.CreateFromRange(
            _fixture.BatchToken, _salesSheetName, "A1:D6", sheet, "A1", name));
        RequireSuccess(_pivotCommands.AddRowField(_fixture.BatchToken, name, row));
        RequireSuccess(_pivotCommands.AddValueField(_fixture.BatchToken, name, "Sales"));
        return name;
    }

    [Fact]
    public void ReportingDepth_SortByValue_UsesSelectedDisplayedValuesInstance()
    {
        var name = CreateReportingDepthPivot("Product");
        var batch = _fixture.BatchToken;
        RequireSuccess(_pivotCommands.AddValueField(batch, name, "Sales", AggregationFunction.Average, "Average Sales"));
        var value = Assert.Single(RequireSuccess(_pivotCommands.ListFields(batch, name)).ValueFields,
            field => field.FieldName == "Average Sales");
        RequireSuccess(_pivotCommands.SortField(batch, name, "Product", SortDirection.Descending, value.FieldName));
        AssertLabels(["Gadget", "Widget"]);
        RequireSuccess(_pivotCommands.SortField(batch, name, "Product", SortDirection.Ascending, value.FieldName));
        AssertLabels(["Widget", "Gadget"]);
        var before = RequireSuccess(_pivotCommands.GetData(batch, name)).Values;
        var invalid = Assert.Throws<ArgumentException>(() =>
            _pivotCommands.SortField(batch, name, "Product", SortDirection.Descending, "Missing value"));
        Assert.Contains("Displayed value field", invalid.Message);
        Assert.Equal(before.SelectMany(row => row), RequireSuccess(_pivotCommands.GetData(batch, name)).Values.SelectMany(row => row));
        AssertOriginalSales();

        void AssertLabels(string[] expected)
        {
            var data = RequireSuccess(_pivotCommands.GetData(batch, name));
            Assert.Equal(expected, data.Values.Select(row => row[0]?.ToString())
                .Where(label => label is "Widget" or "Gadget"));
            Assert.Equal(650d, Convert.ToDouble(data.Values[^1][1], CultureInfo.InvariantCulture));
        }
    }

    [Fact]
    public async Task ReportingDepth_ShowItemsWithNoData_ChangesActualRowsAndPersists()
    {
        var batch = _fixture.BatchToken;
        RequireSuccess(_commands.SetValues(batch, _salesSheetName, "A1:D6",
        [
            ["Region", "Product", "Sales", "Date"],
            ["North", "Widget", 100, "2025-01-15"],
            ["North", "Widget", 150, "2025-01-20"],
            ["South", "Gadget", 200, "2025-02-10"],
            ["North", "Widget", 75, "2025-02-15"],
            ["South", "Gadget", 125, "2025-03-05"]
        ], overwritePolicy: OverwritePolicy.Allow));
        var name = CreateReportingDepthPivot("Product");
        RequireSuccess(_pivotCommands.AddFilterField(batch, name, "Region"));
        RequireSuccess(_pivotCommands.SetFieldFilter(batch, name, "Region", ["North"]));
        RequireSuccess(_pivotCommands.SetFieldOptions(batch, name, "Product", false));
        Assert.DoesNotContain(RequireSuccess(_pivotCommands.GetData(batch, name)).Values,
            row => row[0]?.ToString() == "Gadget");
        var set = RequireSuccess(_pivotCommands.SetFieldOptions(batch, name, "Product", true));
        Assert.True(set.ShowAllItems);
        Assert.True(RequireSuccess(_pivotCommands.GetFieldOptions(batch, name, "Product")).ShowAllItems);
        Assert.Contains(RequireSuccess(_pivotCommands.GetData(batch, name)).Values,
            row => row[0]?.ToString() == "Gadget");
        var display = RequireSuccess(_pivotCommands.SetLayoutOptions(batch, name,
            new PivotLayoutOptions { DisplayNullString = true, NullString = "No sales" }));
        Assert.True(display.DisplayNullString);
        Assert.Equal("No sales", display.NullString);
        Assert.Contains("No sales", ReadReportingDepthTexts(name));
        await _fixture.SaveAndReopenAsync();
        Assert.True(RequireSuccess(_pivotCommands.GetFieldOptions(batch, name, "Product")).ShowAllItems);
        Assert.Contains(RequireSuccess(_pivotCommands.GetData(batch, name)).Values,
            row => row[0]?.ToString() == "Gadget");
        Assert.Contains("No sales", ReadReportingDepthTexts(name));
        var empty = RequireSuccess(_pivotCommands.SetLayoutOptions(batch, name,
            new PivotLayoutOptions { NullString = string.Empty }));
        Assert.True(empty.DisplayNullString);
        Assert.Equal(string.Empty, empty.NullString);
        Assert.DoesNotContain("No sales", ReadReportingDepthTexts(name));
    }

    [Fact]
    public async Task ReportingDepth_DisplayAndPrintingOptions_RoundTripAndOmittedSettingsRemain()
    {
        var name = CreateReportingDepthPivot();
        var batch = _fixture.BatchToken;
        RequireSuccess(_pivotCommands.AddRowField(batch, name, "Product"));
        var options = new PivotLayoutOptions
        {
            RowLayout = 2,
            MergeLabels = true,
            PrintTitles = true,
            RepeatItemsOnEachPrintedPage = true,
            BlankLineAfterEachItem = true,
            DisplayNullString = true,
            NullString = "No sales",
            DisplayErrorString = true,
            ErrorString = "Unavailable"
        };
        AssertOptions(RequireSuccess(_pivotCommands.SetLayoutOptions(batch, name, options)));
        AssertOptions(RequireSuccess(_pivotCommands.SetLayoutOptions(batch, name,
            new PivotLayoutOptions { PreserveFormatting = true })));
        await _fixture.SaveAndReopenAsync();
        AssertOptions(RequireSuccess(_pivotCommands.GetLayoutOptions(batch, name)));
        var invalid = Assert.Throws<NotSupportedException>(() =>
            _pivotCommands.SetLayoutOptions(batch, name, new PivotLayoutOptions { ShowEmptyRows = true }));
        Assert.Contains("OLAP", invalid.Message);
        AssertOptions(RequireSuccess(_pivotCommands.GetLayoutOptions(batch, name)));
        var cleared = RequireSuccess(_pivotCommands.SetLayoutOptions(batch, name,
            new PivotLayoutOptions
            {
                DisplayNullString = false,
                DisplayErrorString = false,
                MergeLabels = false,
                PrintTitles = false,
                BlankLineAfterEachItem = false,
                RepeatItemsOnEachPrintedPage = false
            }));
        Assert.False(cleared.DisplayNullString);
        Assert.False(cleared.DisplayErrorString);
        Assert.False(cleared.MergeLabels);
        Assert.False(cleared.PrintTitles);
        Assert.False(cleared.RepeatItemsOnEachPrintedPage);
        Assert.All(cleared.RowFields, field =>
        {
            Assert.False(field.BlankLineAfterEachItem);
        });

        static void AssertOptions(PivotLayoutResult result)
        {
            Assert.True(result.MergeLabels);
            Assert.True(result.PrintTitles);
            Assert.True(result.RepeatItemsOnEachPrintedPage);
            Assert.True(result.DisplayNullString);
            Assert.Equal("No sales", result.NullString);
            Assert.True(result.DisplayErrorString);
            Assert.Equal("Unavailable", result.ErrorString);
            Assert.Null(result.ShowEmptyRows);
            Assert.Null(result.ShowEmptyColumns);
            Assert.Equal(2, result.RowFields.Count);
            Assert.All(result.RowFields, field =>
            {
                Assert.True(field.BlankLineAfterEachItem);
            });
        }
    }

    [Fact]
    public async Task ReportingDepth_SubtotalFunctions_ReplacePersistAndRejectInvalidWithoutMutation()
    {
        var name = CreateReportingDepthPivot();
        var batch = _fixture.BatchToken;
        RequireSuccess(_pivotCommands.AddRowField(batch, name, "Product"));
        RequireSuccess(_pivotCommands.SetLayout(batch, name, 1));
        var expected = new List<PivotSubtotalFunction> { PivotSubtotalFunction.Sum, PivotSubtotalFunction.Average };
        Assert.Equal(expected, RequireSuccess(_pivotCommands.SetSubtotalFunctions(batch, name, "Region", expected)).Functions);
        Assert.Equal(expected, RequireSuccess(_pivotCommands.GetSubtotalFunctions(batch, name, "Region")).Functions);
        var data = RequireSuccess(_pivotCommands.GetData(batch, name)).Values;
        Assert.Contains(data, row => row[0]?.ToString()?.Contains("North Sum", StringComparison.Ordinal) == true);
        Assert.Contains(data, row => row[0]?.ToString()?.Contains("North Average", StringComparison.Ordinal) == true);
        Assert.Equal(325d, Convert.ToDouble(Assert.Single(data,
            row => row[0]?.ToString() == "North Sum")[^1], CultureInfo.InvariantCulture));
        Assert.Equal(325d / 3, Convert.ToDouble(Assert.Single(data,
            row => row[0]?.ToString() == "North Average")[^1], CultureInfo.InvariantCulture), 8);
        await _fixture.SaveAndReopenAsync();
        Assert.Equal(expected, RequireSuccess(_pivotCommands.GetSubtotalFunctions(batch, name, "Region")).Functions);
        var invalid = Assert.Throws<ArgumentException>(() => _pivotCommands.SetSubtotalFunctions(
            batch, name, "Region", [PivotSubtotalFunction.Automatic, PivotSubtotalFunction.Sum]));
        Assert.Contains("Automatic cannot be combined", invalid.Message);
        Assert.Equal(expected, RequireSuccess(_pivotCommands.GetSubtotalFunctions(batch, name, "Region")).Functions);
        Assert.Empty(RequireSuccess(_pivotCommands.SetSubtotalFunctions(batch, name, "Region", [])).Functions);
        Assert.Equal([PivotSubtotalFunction.Automatic],
            RequireSuccess(_pivotCommands.SetSubtotalFunctions(batch, name, "Region", [PivotSubtotalFunction.Automatic])).Functions);
        foreach (var function in Enum.GetValues<PivotSubtotalFunction>())
        {
            Assert.Equal([function],
                RequireSuccess(_pivotCommands.SetSubtotalFunctions(batch, name, "Region", [function])).Functions);
            Assert.Equal([function], RequireSuccess(_pivotCommands.GetSubtotalFunctions(batch, name, "Region")).Functions);
        }
    }

    [Fact]
    public async Task ReportingDepth_CalculatedItem_ChangesTotalsPersistsAndDeletesOnlyCalculatedItem()
    {
        var name = CreateReportingDepthPivot();
        var batch = _fixture.BatchToken;
        Assert.Empty(RequireSuccess(_pivotCommands.ListCalculatedItems(batch, name, "Region")).Items);
        var created = RequireSuccess(_pivotCommands.CreateCalculatedItem(batch, name, "Region", "Combined", "=North+South"));
        var item = Assert.Single(created.Items);
        Assert.Equal("Combined", item.Name);
        Assert.Contains("North", item.Formula);
        Assert.Contains("South", item.Formula);
        AssertCalculatedTotal(1300);
        await _fixture.SaveAndReopenAsync();
        Assert.Equal("Combined", Assert.Single(RequireSuccess(_pivotCommands.ListCalculatedItems(batch, name, "Region")).Items).Name);
        AssertCalculatedTotal(1300);
        var invalid = Assert.Throws<ArgumentException>(() => _pivotCommands.DeleteCalculatedItem(batch, name, "Region", "North"));
        Assert.Contains("Calculated item 'North' was not found", invalid.Message);
        AssertCalculatedTotal(1300);
        var duplicate = Assert.Throws<ArgumentException>(() => _pivotCommands.CreateCalculatedItem(batch, name, "Region", "Combined", "=North"));
        Assert.Contains("already exists", duplicate.Message);
        Assert.Single(RequireSuccess(_pivotCommands.ListCalculatedItems(batch, name, "Region")).Items);
        Assert.Empty(RequireSuccess(_pivotCommands.DeleteCalculatedItem(batch, name, "Region", "Combined")).Items);
        AssertCalculatedTotal(650);
        AssertOriginalSales();

        void AssertCalculatedTotal(double total)
        {
            var data = RequireSuccess(_pivotCommands.GetData(batch, name));
            Assert.Equal(total, Convert.ToDouble(data.Values[^1][^1], CultureInfo.InvariantCulture));
            if (total == 1300)
                Assert.Contains(data.Values, row => row[0]?.ToString() == "Combined"
                    && Convert.ToDouble(row[^1], CultureInfo.InvariantCulture) == 650);
        }
    }

    [Fact]
    public void ReportingDepth_CalculatedItem_RejectsGroupedPivotWithoutChangingTotals()
    {
        var name = CreateReportingDepthPivot();
        var batch = _fixture.BatchToken;
        var grouped = RequireSuccess(_pivotCommands.GroupItems(batch, name, "Region", ["North", "South"], "All Regions"));
        var before = RequireSuccess(_pivotCommands.GetData(batch, name)).Values.SelectMany(row => row).ToArray();
        var error = Assert.Throws<NotSupportedException>(() =>
            _pivotCommands.CreateCalculatedItem(batch, name, "Region", "Combined", "=North+South"));
        Assert.Contains("grouped fields", error.Message);
        Assert.Equal(before, RequireSuccess(_pivotCommands.GetData(batch, name)).Values.SelectMany(row => row));
        RequireSuccess(_pivotCommands.UngroupField(batch, name, grouped.GroupedFieldName));
        Assert.Empty(RequireSuccess(_pivotCommands.ListCalculatedItems(batch, name, "Region")).Items);
        AssertOriginalSales();
    }

    [Fact]
    public void ReportingDepth_ErrorText_ChangesDisplayedCellsNotSourceData()
    {
        var name = CreateReportingDepthPivot("Product");
        var batch = _fixture.BatchToken;
        RequireSuccess(_pivotCommands.AddFilterField(batch, name, "Region"));
        RequireSuccess(_pivotCommands.SetFieldFilter(batch, name, "Region", ["South"]));
        RequireSuccess(_pivotCommands.CreateCalculatedField(batch, name, "Broken", "=Sales/0"));
        RequireSuccess(_pivotCommands.AddValueField(batch, name, "Broken"));
        var result = RequireSuccess(_pivotCommands.SetLayoutOptions(batch, name, new PivotLayoutOptions
        {
            DisplayErrorString = true,
            ErrorString = "Unavailable"
        }));
        Assert.True(result.DisplayErrorString);
        Assert.Contains("Unavailable", ReadReportingDepthTexts(name));
        var disabled = RequireSuccess(_pivotCommands.SetLayoutOptions(batch, name,
            new PivotLayoutOptions { DisplayErrorString = false, ErrorString = string.Empty }));
        Assert.False(disabled.DisplayErrorString);
        Assert.Equal(string.Empty, disabled.ErrorString);
        Assert.DoesNotContain("Unavailable", ReadReportingDepthTexts(name));
        Assert.Contains("#DIV/0!", ReadReportingDepthTexts(name));
        AssertOriginalSales();
    }

    private List<string> ReadReportingDepthTexts(string name)
    {
        var sheetName = Assert.Single(RequireSuccess(_pivotCommands.List(_fixture.BatchToken)).PivotTables,
            value => value.Name == name).SheetName;
        return _fixture.ExecuteRawVerification((context, ct) =>
        {
            Excel.Worksheet? sheet = null;
            Excel.PivotTables? tables = null;
            Excel.PivotTable? pivot = null;
            Excel.Range? range = null;
            Excel.Range? cells = null;
            Excel.Range? cell = null;
            try
            {
                sheet = ComUtilities.FindSheet(context.Book, sheetName);
                Assert.NotNull(sheet);
                tables = sheet.PivotTables();
                pivot = tables.Item(name);
                range = pivot.DataBodyRange;
                cells = range.Cells;
                var result = new List<string>();
                for (int index = 1; index <= cells.Count; index++)
                {
                    ct.ThrowIfCancellationRequested();
                    try
                    {
                        cell = (Excel.Range)cells[index];
                        result.Add(Convert.ToString(((object)cell.Text), CultureInfo.InvariantCulture) ?? string.Empty);
                    }
                    finally
                    {
                        ComUtilities.Release(ref cell);
                    }
                }
                return result;
            }
            finally
            {
                ComUtilities.Release(ref cell);
                ComUtilities.Release(ref cells);
                ComUtilities.Release(ref range);
                ComUtilities.Release(ref pivot);
                ComUtilities.Release(ref tables);
                ComUtilities.Release(ref sheet);
            }
        });
    }
}
