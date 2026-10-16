using System.Globalization;
using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.Core.Commands.PivotTable;
using Sbroenne.ExcelMcp.Core.Models;
using Xunit;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sbroenne.ExcelMcp.Service.Tests;

public sealed partial class PersistentServiceConnectionTests
{
    [ConfiguredExternalOlapFact]
    [Trait("RunType", "OnDemand")]
    public async Task ExternalOlapPivotTable_ReportingOptionsAndValueSortRoundTripWithExplicitRegularLimits()
    {
        var (sheet, name, _) = CreateConfiguredExternalPivot("ExternalReporting");
        var batch = _fixture.BatchToken;
        var pivots = _fixture.CreateCommands<IPersistentPivotTableCommands>();
        RequireSuccess(pivots.AddRowField(batch, name, ExternalCalendar));
        RequireSuccess(pivots.AddValueField(batch, name, ExternalAmountMeasure));
        var value = Assert.Single(RequireSuccess(pivots.ListFields(batch, name)).ValueFields);
        RequireSuccess(pivots.SortField(batch, name, ExternalCalendar, SortDirection.Descending, value.FieldName));
        AssertLabels(["2025", "2024"]);
        RequireSuccess(pivots.SortField(batch, name, ExternalCalendar, SortDirection.Ascending, value.FieldName));
        AssertLabels(["2024", "2025"]);
        var set = RequireSuccess(pivots.SetLayoutOptions(batch, name, new PivotLayoutOptions
        {
            ShowEmptyRows = true,
            ShowEmptyColumns = true,
            DisplayNullString = true,
            NullString = "No amount",
            DisplayErrorString = true,
            ErrorString = "Unavailable"
        }));
        AssertOptions(set);
        AssertNativeOptions();
        var before = RequireSuccess(pivots.GetData(batch, name)).Values.SelectMany(row => row).ToArray();
        AssertRegularLimit(() => pivots.GetFieldOptions(batch, name, ExternalCalendar));
        AssertRegularLimit(() => pivots.SetFieldOptions(batch, name, ExternalCalendar, true));
        AssertRegularLimit(() => pivots.GetSubtotalFunctions(batch, name, ExternalCalendar));
        AssertRegularLimit(() => pivots.SetSubtotalFunctions(batch, name, ExternalCalendar, [PivotSubtotalFunction.Sum]));
        AssertRegularLimit(() => pivots.ListCalculatedItems(batch, name, ExternalCalendar));
        AssertRegularLimit(() => pivots.CreateCalculatedItem(batch, name, ExternalCalendar, "Combined", "=2024+2025"));
        AssertRegularLimit(() => pivots.DeleteCalculatedItem(batch, name, ExternalCalendar, "2024"));
        Assert.Equal(before, RequireSuccess(pivots.GetData(batch, name)).Values.SelectMany(row => row));
        AssertOptions(RequireSuccess(pivots.GetLayoutOptions(batch, name)));
        await _fixture.SaveAndReopenAsync();
        AssertOptions(RequireSuccess(pivots.GetLayoutOptions(batch, name)));
        AssertNativeOptions();
        AssertLabels(["2024", "2025"]);
        RequireSuccess(pivots.Refresh(batch, name));
        var cleared = RequireSuccess(pivots.SetLayoutOptions(batch, name, new PivotLayoutOptions
        {
            ShowEmptyRows = false,
            ShowEmptyColumns = false,
            DisplayNullString = false,
            DisplayErrorString = false
        }));
        Assert.False(cleared.ShowEmptyRows);
        Assert.False(cleared.ShowEmptyColumns);
        Assert.False(cleared.DisplayNullString);
        Assert.False(cleared.DisplayErrorString);
        AssertLabels(["2024", "2025"]);

        static void AssertRegularLimit(Action action)
        {
            var error = Assert.Throws<NotSupportedException>(action);
            Assert.Contains("regular PivotTables only", error.Message);
            Assert.Contains("OLAP/Data Model", error.Message);
        }

        static void AssertOptions(PivotLayoutResult result)
        {
            Assert.True(result.ShowEmptyRows);
            Assert.True(result.ShowEmptyColumns);
            Assert.True(result.DisplayNullString);
            Assert.Equal("No amount", result.NullString);
            Assert.True(result.DisplayErrorString);
            Assert.Equal("Unavailable", result.ErrorString);
        }

        void AssertLabels(string[] expected)
        {
            var data = RequireSuccess(pivots.GetData(batch, name));
            Assert.Equal(expected, data.Values.Select(row => row[0]?.ToString())
                .Where(label => label is "2024" or "2025"));
            Assert.Equal(300d, Convert.ToDouble(data.Values[^1][1], CultureInfo.InvariantCulture));
        }

        void AssertNativeOptions() => _fixture.ExecuteRawVerification((context, _) =>
        {
            Excel.Worksheet? worksheet = null;
            Excel.PivotTables? tables = null;
            Excel.PivotTable? pivot = null;
            try
            {
                worksheet = ComUtilities.FindSheet(context.Book, sheet);
                Assert.NotNull(worksheet);
                tables = worksheet.PivotTables();
                pivot = tables.Item(name);
                Assert.True(pivot.DisplayEmptyRow);
                Assert.True(pivot.DisplayEmptyColumn);
                Assert.True(pivot.DisplayNullString);
                Assert.Equal("No amount", pivot.NullString);
                Assert.True(pivot.DisplayErrorString);
                Assert.Equal("Unavailable", pivot.ErrorString);
                return true;
            }
            finally
            {
                ComUtilities.Release(ref pivot);
                ComUtilities.Release(ref tables);
                ComUtilities.Release(ref worksheet);
            }
        });
    }
}
