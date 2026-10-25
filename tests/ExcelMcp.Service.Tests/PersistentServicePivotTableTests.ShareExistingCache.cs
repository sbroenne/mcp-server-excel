using System.Text.Json;
using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.Core.Commands.PivotTable;
using Sbroenne.ExcelMcp.Core.Commands.Range;
using Sbroenne.ExcelMcp.Core.Models;
using Xunit;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sbroenne.ExcelMcp.Service.Tests;

public sealed partial class PersistentServicePivotTableTests
{
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void ShareExistingCache_MovesReportWithoutCreatingCache(bool table)
    {
        var batch = _fixture.BatchToken;
        var firstSheet = _fixture.CreateTestSheet(batch);
        var secondSheet = _fixture.CreateTestSheet(batch);
        var thirdSheet = _fixture.CreateTestSheet(batch);
        if (table)
            RequireSuccess(_tableCommands.Create(batch, _salesSheetName, "SharedSales", "A1:D6", true));
        RequireSuccess(CreateCacheReport(table, firstSheet, "First"));
        RequireSuccess(CreateCacheReport(table, secondSheet, "Second"));
        RequireSuccess(CreateCacheReport(table, thirdSheet, "Independent"));
        foreach (var name in new[] { "First", "Second", "Independent" })
        {
            RequireSuccess(_pivotCommands.AddRowField(batch, name, name == "Second" ? "Product" : "Region"));
            RequireSuccess(_pivotCommands.AddValueField(batch, name, "Sales"));
        }
        var targetIndex = ReadNativePivot(firstSheet, "First", pivot => pivot.CacheIndex);
        var oldIndex = ReadNativePivot(secondSheet, "Second", pivot => pivot.CacheIndex);
        var count = CountNativePivotCaches();
        Assert.NotEqual(targetIndex, oldIndex);
        var before = SharingReportSnapshot("Second");

        var response = _fixture.Send("pivottable.share-cache", new
        {
            sheetName = secondSheet,
            pivotTableName = "Second",
            shareCacheWith = "First",
            shareCacheWithSheet = firstSheet
        });

        using var json = JsonDocument.Parse(response.Result!);
        var result = json.RootElement;
        Assert.True(result.GetProperty("changed").GetBoolean());
        Assert.Equal(oldIndex, result.GetProperty("previousCacheIndex").GetInt32());
        Assert.Equal(targetIndex, result.GetProperty("cacheIndex").GetInt32());
        Assert.Equal(targetIndex, ReadNativePivot(secondSheet, "Second", pivot => pivot.CacheIndex));
        Assert.InRange(CountNativePivotCaches(), count - 1, count);
        Assert.Equal(before, SharingReportSnapshot("Second"));
        Assert.NotEqual(targetIndex, ReadNativePivot(thirdSheet, "Independent", pivot => pivot.CacheIndex));
        var users = result.GetProperty("sharedPivotTables").EnumerateArray().ToArray();
        Assert.Equal(2, users.Length);
        Assert.Contains(users, item => item.GetProperty("sheetName").GetString() == firstSheet &&
            item.GetProperty("pivotTableName").GetString() == "First");
        Assert.Contains(users, item => item.GetProperty("sheetName").GetString() == secondSheet &&
            item.GetProperty("pivotTableName").GetString() == "Second");
        RequireSuccess(_pivotCommands.CreateSlicer(batch, "First", "Region", "SharedRegion", firstSheet, "H2"));
        RequireSuccess(_pivotCommands.ConnectPivotTable(batch, "SharedRegion", "Second"));
        RequireSuccess(_pivotCommands.SetSlicerSelection(batch, "SharedRegion", ["North"]));
        AssertFilteredPivot("First", "North", 325);
        AssertCacheReportTotals("Second", new Dictionary<string, double> { ["Gadget"] = 75, ["Widget"] = 250 }, 325);
        AssertPivotSales(325, 325, "Independent");
        RequireSuccess(_pivotCommands.SetSlicerSelection(batch, "SharedRegion", []));
        AssertPivotSales(325, 325, "First");
        AssertCacheReportTotals("Second", new Dictionary<string, double> { ["Gadget"] = 275, ["Widget"] = 375 }, 650);
    }

    private string SharingReportSnapshot(string name)
    {
        var read = RequireSuccess(_pivotCommands.Read(_fixture.BatchToken, name));
        read.PivotTable.CacheIndex = 0;
        read.PivotTable.LastRefresh = null;
        read.Fields = read.Fields.OrderBy(field => field.Name, StringComparer.Ordinal).ToList();
        return JsonSerializer.Serialize(read);
    }

    private (string Target, string Moved) CreateSharingPair()
    {
        var batch = _fixture.BatchToken;
        var target = _fixture.CreateTestSheet(batch);
        var moved = _fixture.CreateTestSheet(batch);
        RequireSuccess(CreateCacheReport(false, target, "First"));
        RequireSuccess(CreateCacheReport(false, moved, "Second"));
        foreach (var name in new[] { "First", "Second" })
        {
            RequireSuccess(_pivotCommands.AddRowField(batch, name, "Region"));
            RequireSuccess(_pivotCommands.AddValueField(batch, name, "Sales"));
        }
        return (target, moved);
    }

    private PivotCacheSharingResult ShareExisting(string target, string moved) =>
        RequireSuccess(_pivotCommands.ShareCache(_fixture.BatchToken, moved, "Second", "First", target));

    [Theory]
    [InlineData("value-filter")]
    [InlineData("value-sort")]
    [InlineData("manual-filter")]
    [InlineData("manual-sort")]
    [InlineData("manual-and-value-filter")]
    public void ShareExistingCache_AdoptsDifferentAggregatesWithoutChangingSettings(string setting)
    {
        var (target, moved) = CreateSharingPair();
        var batch = _fixture.BatchToken;
        if (setting == "manual-and-value-filter")
        {
            RequireSuccess(_pivotCommands.SetLayoutOptions(batch, "Second",
                new PivotLayoutOptions { AllowMultipleFilters = true }));
            RequireSuccess(_pivotCommands.SetFieldFilter(batch, "Second", "Region", ["South"]));
            RequireSuccess(_pivotCommands.AddFieldFilter(batch, "Second", "Region",
                new PivotFilterOptions { Type = PivotFilterType.ValueIsGreaterThan, Number1 = 400, DataFieldName = "Sum of Sales" }));
        }
        else if (setting == "value-filter")
            RequireSuccess(_pivotCommands.AddFieldFilter(batch, "Second", "Region",
                new PivotFilterOptions { Type = PivotFilterType.ValueIsGreaterThan, Number1 = 300, DataFieldName = "Sum of Sales" }));
        else if (setting == "value-sort")
            RequireSuccess(_pivotCommands.SortField(batch, "Second", "Region", SortDirection.Descending, "Sum of Sales"));
        else if (setting == "manual-filter")
            RequireSuccess(_pivotCommands.SetFieldFilter(batch, "Second", "Region", ["North"]));
        else
            ReadNativePivot(moved, "Second", pivot =>
            {
                Excel.PivotField? field = null;
                Excel.PivotItem? item = null;
                try
                {
                    field = (Excel.PivotField)pivot.PivotFields("Region");
                    item = (Excel.PivotItem)field.PivotItems("South");
                    item.Position = 1;
                    return 0;
                }
                finally
                {
                    ComUtilities.Release(ref item);
                    ComUtilities.Release(ref field);
                }
            });
        var before = RequireSuccess(_pivotCommands.Read(batch, "Second"));
        RequireSuccess(_commands.SetValues(batch, _salesSheetName, "A2:A6",
            [["North"], ["South"], ["South"], ["South"], ["North"]], overwritePolicy: OverwritePolicy.Allow));
        RequireSuccess(_pivotCommands.Refresh(batch, "First"));
        AssertPivotSales(225, 425, "First");
        if (setting == "manual-filter")
            AssertFilteredPivot("Second", "North", 325);
        else if (setting == "manual-and-value-filter")
            Assert.DoesNotContain(RequireSuccess(_pivotCommands.GetData(batch, "Second")).Values,
                row => row[0]?.ToString() is "North" or "South");
        else
            AssertPivotSales(325, 325, "Second");
        var originalIndex = ReadNativePivot(moved, "Second", pivot => pivot.CacheIndex);
        Assert.NotEqual(ReadNativePivot(target, "First", pivot => pivot.CacheIndex), originalIndex);

        var result = ShareExisting(target, moved);

        Assert.True(result.Changed);
        Assert.Equal(originalIndex, result.PreviousCacheIndex);
        Assert.Equal(ReadNativePivot(target, "First", pivot => pivot.CacheIndex),
            ReadNativePivot(moved, "Second", pivot => pivot.CacheIndex));
        AssertPivotSales(225, 425, "First");
        var after = RequireSuccess(_pivotCommands.Read(batch, "Second"));
        Assert.Equal(JsonSerializer.Serialize(before.ValueFields), JsonSerializer.Serialize(after.ValueFields));
        var region = Assert.Single(after.FieldSelections, field => field.FieldName == "Region");
        Assert.True(region.FiltersAvailable);
        Assert.Equal(JsonSerializer.Serialize(Assert.Single(before.FieldSelections, field => field.FieldName == "Region").Filters),
            JsonSerializer.Serialize(region.Filters));
        if (setting is "value-filter" or "manual-and-value-filter")
        {
            AssertFilteredPivot("Second", "South", 425);
            if (setting == "manual-and-value-filter")
                Assert.False(ReadNativePivot(moved, "Second", pivot =>
                {
                    Excel.PivotField? field = null;
                    Excel.PivotItem? item = null;
                    try
                    {
                        field = (Excel.PivotField)pivot.PivotFields("Region");
                        item = (Excel.PivotItem)field.PivotItems("North");
                        return item.Visible;
                    }
                    finally
                    {
                        ComUtilities.Release(ref item);
                        ComUtilities.Release(ref field);
                    }
                }));
        }
        else if (setting == "manual-filter")
        {
            AssertFilteredPivot("Second", "North", 225);
            Assert.Equal(["North"], region.SelectedMemberUniqueNames);
        }
        else
        {
            AssertPivotSales(225, 425, "Second");
            Assert.Equal(["South", "North"],
                RequireSuccess(_pivotCommands.GetData(batch, "Second")).Values
                    .Select(row => row[0]?.ToString()).Where(label => label is "North" or "South"));
        }
    }

    [Theory]
    [InlineData("manual-filter")]
    [InlineData("page-filter")]
    [InlineData("label-filter")]
    [InlineData("value-filter")]
    [InlineData("label-sort")]
    [InlineData("value-sort")]
    [InlineData("manual-sort")]
    [InlineData("calculation")]
    [InlineData("formatting")]
    [InlineData("direct-formatting")]
    [InlineData("layout")]
    public void ShareExistingCache_PreservesReportSettings(string setting)
    {
        var (target, moved) = CreateSharingPair();
        var batch = _fixture.BatchToken;
        switch (setting)
        {
            case "manual-filter":
                RequireSuccess(_pivotCommands.SetFieldFilter(batch, "Second", "Region", ["North"]));
                break;
            case "page-filter":
                RequireSuccess(_pivotCommands.AddFilterField(batch, "Second", "Product"));
                RequireSuccess(_pivotCommands.SetFieldFilter(batch, "Second", "Product", ["Widget"]));
                break;
            case "label-filter":
                RequireSuccess(_pivotCommands.AddFieldFilter(batch, "Second", "Region",
                    new PivotFilterOptions { Type = PivotFilterType.CaptionBeginsWith, Text1 = "N" }));
                break;
            case "value-filter":
                RequireSuccess(_pivotCommands.AddFieldFilter(batch, "Second", "Region",
                    new PivotFilterOptions { Type = PivotFilterType.ValueIsGreaterThan, Number1 = 300, DataFieldName = "Sum of Sales" }));
                break;
            case "label-sort":
            case "value-sort":
                RequireSuccess(_pivotCommands.SortField(batch, "Second", "Region", SortDirection.Descending,
                    setting == "value-sort" ? "Sum of Sales" : null));
                break;
            case "manual-sort":
                ReadNativePivot(moved, "Second", pivot =>
                {
                    Excel.PivotField? field = null;
                    Excel.PivotItem? item = null;
                    try
                    {
                        field = (Excel.PivotField)pivot.PivotFields("Region");
                        item = (Excel.PivotItem)field.PivotItems("South");
                        item.Position = 1;
                        return 0;
                    }
                    finally
                    {
                        ComUtilities.Release(ref item);
                        ComUtilities.Release(ref field);
                    }
                });
                break;
            case "calculation":
                RequireSuccess(_pivotCommands.SetFieldCalculation(batch, "Second", "Sum of Sales",
                    PivotFieldCalculation.PercentOfTotal));
                break;
            case "formatting":
                RequireSuccess(_pivotCommands.SetLayoutOptions(batch, "Second",
                    new PivotLayoutOptions { RowLayout = 1, StyleName = "PivotStyleMedium9", PreserveFormatting = true, ShowRowStripes = true }));
                RequireSuccess(_pivotCommands.SetFieldFormat(batch, "Second", "Sum of Sales", "#,##0.00"));
                RequireSuccess(_commands.SetNumberFormat(batch, moved, "B2", "0.000"));
                break;
            case "direct-formatting":
                ReadNativePivot(moved, "Second", pivot =>
                {
                    Excel.Range? extent = null;
                    Excel.Font? font = null;
                    Excel.Interior? fill = null;
                    try
                    {
                        extent = pivot.TableRange2;
                        font = extent.Font;
                        fill = extent.Interior;
                        font.Bold = true;
                        font.Color = 255;
                        fill.Color = 65535;
                        return 0;
                    }
                    finally
                    {
                        ComUtilities.Release(ref fill);
                        ComUtilities.Release(ref font);
                        ComUtilities.Release(ref extent);
                    }
                });
                break;
            case "layout":
                RequireSuccess(_pivotCommands.AddColumnField(batch, "Second", "Product"));
                RequireSuccess(_pivotCommands.AddValueField(batch, "Second", "Sales",
                    AggregationFunction.Average, "Average Sales"));
                RequireSuccess(_pivotCommands.SetSubtotalFunctions(batch, "Second", "Region",
                    [PivotSubtotalFunction.Sum, PivotSubtotalFunction.Average]));
                RequireSuccess(_pivotCommands.SetLayoutOptions(batch, "Second",
                    new PivotLayoutOptions { RowLayout = 1 }));
                ReadNativePivot(moved, "Second", pivot =>
                {
                    Excel.PivotField? values = null;
                    try
                    {
                        pivot.RowGrand = false;
                        pivot.ColumnGrand = false;
                        values = pivot.DataPivotField;
                        values.Orientation = Excel.XlPivotFieldOrientation.xlRowField;
                        return 0;
                    }
                    finally { ComUtilities.Release(ref values); }
                });
                var layout = RequireSuccess(_pivotCommands.GetLayoutOptions(batch, "Second"));
                Assert.Equal("Region", Assert.Single(layout.RowFields).FieldName);
                Assert.Equal(1, layout.RowFields[0].RowLayout);
                break;
        }
        var before = SharingReportSnapshot("Second");
        var nativeBefore = NativeSharingSettings(moved, "Second");
        var valuesBefore = JsonSerializer.Serialize(RequireSuccess(_pivotCommands.GetData(batch, "Second")).Values);

        ShareExisting(target, moved);

        Assert.Equal(before, SharingReportSnapshot("Second"));
        Assert.Equal(nativeBefore, NativeSharingSettings(moved, "Second"));
        Assert.Equal(valuesBefore, JsonSerializer.Serialize(RequireSuccess(_pivotCommands.GetData(batch, "Second")).Values));
    }

    private string NativeSharingSettings(string sheet, string name) =>
        ReadNativePivot(sheet, name, pivot =>
        {
            Excel.PivotField? field = null;
            Excel.PivotItems? items = null;
            Excel.Range? extent = null;
            Excel.Range? cells = null;
            Excel.PivotFields? values = null;
            Excel.PivotField? valuesPlacement = null;
            try
            {
                field = (Excel.PivotField)pivot.PivotFields("Region");
                items = (Excel.PivotItems)field.PivotItems();
                List<string> order = [];
                for (int i = 1; i <= items.Count; i++)
                {
                    Excel.PivotItem? item = null;
                    try
                    {
                        item = items.Item(i);
                        order.Add(item.Name);
                    }
                    finally { ComUtilities.Release(ref item); }
                }
                extent = pivot.TableRange2;
                cells = extent.Cells;
                List<object?> formats = [];
                for (int i = 1; i <= cells.Count; i++)
                {
                    Excel.Range? cell = null;
                    Excel.Font? font = null;
                    Excel.Interior? fill = null;
                    try
                    {
                        cell = (Excel.Range)cells[i];
                        font = cell.Font;
                        fill = cell.Interior;
                        formats.Add(new
                        {
                            Format = (object)cell.NumberFormat,
                            Bold = (object)font.Bold,
                            FontColor = (object)font.Color,
                            FillColor = (object)fill.Color
                        });
                    }
                    finally
                    {
                        ComUtilities.Release(ref fill);
                        ComUtilities.Release(ref font);
                        ComUtilities.Release(ref cell);
                    }
                }
                values = (Excel.PivotFields)pivot.DataFields;
                if (values.Count > 1)
                    valuesPlacement = pivot.DataPivotField;
                return JsonSerializer.Serialize(new
                {
                    field.AutoSortOrder,
                    field.AutoSortField,
                    Order = order,
                    Style = (object)pivot.TableStyle2,
                    pivot.PreserveFormatting,
                    ValuesPlacement = valuesPlacement is null ? null : new
                    {
                        valuesPlacement.Orientation,
                        Position = Convert.ToInt32((object)valuesPlacement.Position, System.Globalization.CultureInfo.InvariantCulture)
                    },
                    pivot.ShowTableStyleRowStripes,
                    pivot.RowGrand,
                    pivot.ColumnGrand,
                    field.LayoutForm,
                    field.LayoutCompactRow,
                    Formats = formats
                });
            }
            finally
            {
                ComUtilities.Release(ref valuesPlacement);
                ComUtilities.Release(ref values);
                ComUtilities.Release(ref cells);
                ComUtilities.Release(ref extent);
                ComUtilities.Release(ref items);
                ComUtilities.Release(ref field);
            }
        });

    [Theory]
    [InlineData("numeric", false)]
    [InlineData("numeric", true)]
    [InlineData("date", false)]
    [InlineData("date", true)]
    [InlineData("manual", false)]
    [InlineData("manual", true)]
    [InlineData("calculated-item", false)]
    [InlineData("calculated-item", true)]
    [InlineData("calculated-field-missing", false)]
    [InlineData("calculated-field-conflict", false)]
    [InlineData("slicer", false)]
    [InlineData("protected", false)]
    public void ShareExistingCache_RejectsUnsupportedConfigurationBeforeWrite(string setting, bool onTarget)
    {
        var (target, moved) = CreateSharingPair();
        var batch = _fixture.BatchToken;
        string name = onTarget ? "First" : "Second";
        switch (setting)
        {
            case "numeric":
                RequireSuccess(_pivotCommands.RemoveField(batch, name, "Region"));
                RequireSuccess(_pivotCommands.AddRowField(batch, name, "Sales"));
                RequireSuccess(_pivotCommands.GroupByNumeric(batch, name, "Sales", 0, 300, 100));
                break;
            case "date":
                RequireSuccess(_pivotCommands.RemoveField(batch, name, "Region"));
                RequireSuccess(_pivotCommands.AddRowField(batch, name, "Date"));
                RequireSuccess(_pivotCommands.GroupByDate(batch, name, "Date", DateGroupingInterval.Months));
                break;
            case "manual":
                RequireSuccess(_pivotCommands.AddRowField(batch, name, "Product"));
                RequireSuccess(_pivotCommands.GroupItems(batch, name, "Product", ["Widget", "Gadget"], "Combined"));
                break;
            case "calculated-item":
                RequireSuccess(_pivotCommands.CreateCalculatedItem(batch, name, "Region", "Both", "=North+South"));
                break;
            case "calculated-field-missing":
            case "calculated-field-conflict":
                RequireSuccess(_pivotCommands.CreateCalculatedField(batch, "Second", "DoubleSales", "=Sales*2"));
                if (setting == "calculated-field-conflict")
                    RequireSuccess(_pivotCommands.CreateCalculatedField(batch, "First", "DoubleSales", "=Sales*3"));
                break;
            case "slicer":
                RequireSuccess(_pivotCommands.CreateSlicer(batch, "Second", "Region", "MovedSlicer", moved, "H2"));
                break;
            case "protected":
                ReadNativePivot(moved, "Second", pivot =>
                {
                    Excel.Worksheet? sheet = null;
                    try { sheet = (Excel.Worksheet)pivot.Parent; sheet.Protect(); return 0; }
                    finally { ComUtilities.Release(ref sheet); }
                });
                break;
        }
        var original = SnapshotPivot("Second");
        var targetBefore = SnapshotPivot("First");
        var count = CountNativePivotCaches();
        try
        {
            var error = Assert.ThrowsAny<Exception>(() => ShareExisting(target, moved));
            Assert.DoesNotContain("Actual partial state", error.Message, StringComparison.Ordinal);
            Assert.Equal(original, SnapshotPivot("Second"));
            Assert.Equal(targetBefore, SnapshotPivot("First"));
            Assert.Equal(count, CountNativePivotCaches());
        }
        finally
        {
            if (setting == "protected")
                ReadNativePivot(moved, "Second", pivot =>
                {
                    Excel.Worksheet? sheet = null;
                    try { sheet = (Excel.Worksheet)pivot.Parent; sheet.Unprotect(); return 0; }
                    finally { ComUtilities.Release(ref sheet); }
                });
        }
    }

    [Fact]
    public async Task ShareExistingCache_CalculationsAndNoOpPersistAfterReopen()
    {
        var (target, moved) = CreateSharingPair();
        var batch = _fixture.BatchToken;
        RequireSuccess(_pivotCommands.CreateCalculatedField(batch, "First", "DoubleSales", "=Sales*2"));
        RequireSuccess(_pivotCommands.CreateCalculatedField(batch, "Second", "DoubleSales", "=Sales*2"));
        RequireSuccess(_pivotCommands.AddValueField(batch, "Second", "DoubleSales"));
        RequireSuccess(_pivotCommands.CreateCalculatedField(batch, "First", "TripleSales", "=Sales*3"));
        Assert.True(ShareExisting(target, moved).Changed);
        Assert.Contains(RequireSuccess(_pivotCommands.ListCalculatedFields(batch, "Second")).CalculatedFields,
            field => field.Name == "TripleSales" && field.Formula == "=Sales*3");
        Assert.False(ShareExisting(target, moved).Changed);
        var read = RequireSuccess(_pivotCommands.Read(batch, "Second"));
        Assert.Equal(ReadNativePivot(target, "First", pivot => pivot.CacheIndex), read.PivotTable.CacheIndex);
        var values = RequireSuccess(_pivotCommands.GetData(batch, "Second")).Values;
        Assert.Equal(1300d, Convert.ToDouble(values[^1][2], System.Globalization.CultureInfo.InvariantCulture));

        await _fixture.SaveAndReopenAsync();

        Assert.Equal(ReadNativePivot(target, "First", pivot => pivot.CacheIndex),
            ReadNativePivot(moved, "Second", pivot => pivot.CacheIndex));
        Assert.False(ShareExisting(target, moved).Changed);
        Assert.Equal(JsonSerializer.Serialize(values),
            JsonSerializer.Serialize(RequireSuccess(_pivotCommands.GetData(_fixture.BatchToken, "Second")).Values));
    }

    [Fact]
    public void ShareExistingCache_MatchingCalculationsInDifferentOrderPreservePlacedFields()
    {
        var (target, moved) = CreateSharingPair();
        var batch = _fixture.BatchToken;
        RequireSuccess(_pivotCommands.CreateCalculatedField(batch, "First", "TripleSales", "=Sales*3"));
        RequireSuccess(_pivotCommands.CreateCalculatedField(batch, "First", "DoubleSales", "=Sales*2"));
        RequireSuccess(_pivotCommands.CreateCalculatedField(batch, "Second", "DoubleSales", "=Sales*2"));
        RequireSuccess(_pivotCommands.CreateCalculatedField(batch, "Second", "TripleSales", "=Sales*3"));
        RequireSuccess(_pivotCommands.AddValueField(batch, "Second", "DoubleSales"));
        RequireSuccess(_pivotCommands.AddValueField(batch, "Second", "TripleSales"));
        var before = SharingReportSnapshot("Second");
        var values = JsonSerializer.Serialize(RequireSuccess(_pivotCommands.GetData(batch, "Second")).Values);

        Assert.True(ShareExisting(target, moved).Changed);

        Assert.Equal(before, SharingReportSnapshot("Second"));
        Assert.Equal(values, JsonSerializer.Serialize(RequireSuccess(_pivotCommands.GetData(batch, "Second")).Values));
    }

    [Theory]
    [InlineData("different-sheet")]
    [InlineData("different-range")]
    [InlineData("table-versus-range")]
    [InlineData("missing-report")]
    [InlineData("missing-sheet")]
    public void ShareExistingCache_RejectsWrongSourceOrMissingReport(string scenario)
    {
        var (target, moved) = CreateSharingPair();
        var batch = _fixture.BatchToken;
        var other = _fixture.CreateTestSheet(batch);
        if (scenario == "different-sheet")
        {
            RequireSuccess(_commands.SetValues(batch, other, "A1:D6",
                RequireSuccess(_commands.GetValues(batch, _salesSheetName, "A1:D6")).Values));
            RequireSuccess(_pivotCommands.CreateFromRange(batch, other, "A1:D6", other, "H1", "Wrong"));
        }
        if (scenario == "different-range")
            RequireSuccess(_pivotCommands.CreateFromRange(batch, _salesSheetName, "A1:D5", other, "H1", "Wrong"));
        if (scenario == "table-versus-range")
        {
            RequireSuccess(_tableCommands.Create(batch, _salesSheetName, "SharedSales", "A1:D6", true));
            RequireSuccess(_pivotCommands.CreateFromTable(batch, "SharedSales", other, "H1", "Wrong"));
        }
        var before = SnapshotPivot("Second");
        var count = CountNativePivotCaches();
        var error = Assert.ThrowsAny<Exception>(() => _pivotCommands.ShareCache(batch, moved, "Second",
            scenario == "missing-report" ? "Missing" : scenario == "missing-sheet" ? "First" : "Wrong",
            scenario == "missing-sheet" ? "Missing" : scenario == "missing-report" ? target : other));
        Assert.Contains(scenario.StartsWith("missing", StringComparison.Ordinal) ? "not found" : "source definitions differ",
            error.Message, StringComparison.Ordinal);
        Assert.Equal(before, SnapshotPivot("Second"));
        Assert.Equal(count, CountNativePivotCaches());
    }

    [Fact]
    public void ShareExistingCache_RepeatedNamesAndOldCachePeersRemainIndependent()
    {
        var (target, moved) = CreateSharingPair();
        var batch = _fixture.BatchToken;
        var peer = _fixture.CreateTestSheet(batch);
        RequireSuccess(CreateCacheReport(false, peer, "Peer", "Second", moved));
        var decoy = _fixture.CreateTestSheet(batch);
        RequireSuccess(_pivotCommands.CreateFromRange(batch, _salesSheetName, "A1:D5", decoy, "A1", "Decoy"));
        ReadNativePivot(decoy, "Decoy", pivot => { pivot.Name = "First"; return 0; });
        var before = ReadNativePivot(peer, "Peer", pivot => pivot.CacheIndex);
        var count = CountNativePivotCaches();

        var result = ShareExisting(target, moved);

        Assert.True(result.Changed);
        Assert.Equal(count, CountNativePivotCaches());
        var old = Assert.Single(result.PreviousCacheUsers);
        Assert.Equal(peer, old.SheetName);
        Assert.Equal("Peer", old.PivotTableName);
        Assert.Equal(before, ReadNativePivot(peer, "Peer", pivot => pivot.CacheIndex));
        Assert.NotEqual(result.CacheIndex, ReadNativePivot(decoy, "First", pivot => pivot.CacheIndex));
        Assert.Equal(ReadNativePivot(target, "First", pivot => pivot.CacheIndex), result.CacheIndex);
    }

    [Fact]
    public void ShareExistingCache_AlreadySharingWithControls_IsUnchanged()
    {
        var (target, moved) = CreateSharingPair();
        ShareExisting(target, moved);
        RequireSuccess(_pivotCommands.CreateSlicer(_fixture.BatchToken, "Second", "Region", "Existing", moved, "H2"));
        var before = SnapshotPivot("Second");
        var count = CountNativePivotCaches();

        var result = ShareExisting(target, moved);

        Assert.False(result.Changed);
        Assert.Equal(result.CacheIndex, result.PreviousCacheIndex);
        Assert.Equal(result.CacheIndex, result.TargetCacheIndex);
        Assert.Equal(before, SnapshotPivot("Second"));
        Assert.Equal(count, CountNativePivotCaches());
        Assert.Equal(2, result.PreviousCacheUsers.Count);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void ShareExistingCache_TimelineRequiresExplicitDisconnection(bool disconnect)
    {
        var (target, moved) = CreateSharingPair();
        var batch = _fixture.BatchToken;
        RequireSuccess(_pivotCommands.CreateTimeline(batch, "Second", "Date", "MovedTimeline", moved, "H2"));
        if (disconnect)
        {
            var peer = _fixture.CreateTestSheet(batch);
            RequireSuccess(CreateCacheReport(false, peer, "TimelinePeer", "Second", moved));
            RequireSuccess(_pivotCommands.ConnectPivotTable(batch, "MovedTimeline", "TimelinePeer"));
            RequireSuccess(_pivotCommands.DisconnectPivotTable(batch, "MovedTimeline", "Second"));
            Assert.True(ShareExisting(target, moved).Changed);
            var controls = RequireSuccess(_pivotCommands.GetConnection(batch, moved, "Second")).ConnectedSlicerCaches;
            Assert.Empty(controls);
        }
        else
        {
            var before = SnapshotPivot("Second");
            var error = Assert.Throws<InvalidOperationException>(() => ShareExisting(target, moved));
            Assert.Contains("Disconnect", error.Message, StringComparison.Ordinal);
            Assert.Equal(before, SnapshotPivot("Second"));
        }
    }

    [Theory]
    [InlineData("numeric")]
    [InlineData("date")]
    [InlineData("days")]
    public void ShareExistingCache_MatchingGroupingIsStillRejected(string grouping)
    {
        var (target, moved) = CreateSharingPair();
        var batch = _fixture.BatchToken;
        foreach (var name in new[] { "First", "Second" })
        {
            RequireSuccess(_pivotCommands.RemoveField(batch, name, "Region"));
            RequireSuccess(_pivotCommands.AddRowField(batch, name, grouping == "numeric" ? "Sales" : "Date"));
            if (grouping == "numeric")
                RequireSuccess(_pivotCommands.GroupByNumeric(batch, name, "Sales", 0, 300, 100));
            else
                RequireSuccess(_pivotCommands.GroupByDate(batch, name, "Date",
                    grouping == "days" ? DateGroupingInterval.Days : DateGroupingInterval.Months));
        }
        var before = SnapshotPivot("Second");
        var error = Assert.Throws<NotSupportedException>(() => ShareExisting(target, moved));
        Assert.Contains("Grouped field", error.Message, StringComparison.Ordinal);
        Assert.Equal(before, SnapshotPivot("Second"));
    }

    [Fact]
    public void ShareExistingCache_RejectsStaleCachedMembersBeforeWrites()
    {
        var (target, moved) = CreateSharingPair();
        var batch = _fixture.BatchToken;
        RequireSuccess(_commands.SetValues(batch, _salesSheetName, "C2", [[999]], overwritePolicy: OverwritePolicy.Allow));
        RequireSuccess(_pivotCommands.Refresh(batch, "First"));
        RequireSuccess(_commands.SetValues(batch, _salesSheetName, "C2", [[100]], overwritePolicy: OverwritePolicy.Allow));
        var before = SnapshotPivot("Second");
        var targetBefore = SnapshotPivot("First");

        var error = Assert.Throws<NotSupportedException>(() => ShareExisting(target, moved));

        Assert.Contains("incompatible cached items", error.Message, StringComparison.Ordinal);
        Assert.Equal(before, SnapshotPivot("Second"));
        Assert.Equal(targetBefore, SnapshotPivot("First"));
    }

    [Fact]
    public void ShareExistingCache_TargetRenumberingDoesNotMisreportNoOp()
    {
        var (target, moved) = CreateSharingPair();
        var oldIndex = ReadNativePivot(target, "First", pivot => pivot.CacheIndex);
        var chosenIndex = ReadNativePivot(moved, "Second", pivot => pivot.CacheIndex);
        Assert.NotEqual(oldIndex, chosenIndex);

        var result = RequireSuccess(_pivotCommands.ShareCache(_fixture.BatchToken, target, "First", "Second", moved));

        Assert.True(result.Changed);
        Assert.Equal(oldIndex, result.PreviousCacheIndex);
        Assert.Equal(ReadNativePivot(target, "First", pivot => pivot.CacheIndex), result.CacheIndex);
        Assert.Equal(ReadNativePivot(moved, "Second", pivot => pivot.CacheIndex), result.TargetCacheIndex);
        Assert.Equal(result.TargetCacheIndex, result.CacheIndex);
        Assert.Empty(result.PreviousCacheUsers);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void ShareExistingCache_QualifiedSourceFormsResolveSameIdentity(bool table)
    {
        var batch = _fixture.BatchToken;
        var source = $"O'Brien {Guid.NewGuid():N}"[..31];
        _fixture.CreateNamedTestSheet(batch, source);
        RequireSuccess(_commands.SetValues(batch, source, "A1:D6",
            RequireSuccess(_commands.GetValues(batch, _salesSheetName, "A1:D6")).Values));
        var target = _fixture.CreateTestSheet(batch);
        var moved = _fixture.CreateTestSheet(batch);
        if (table)
            RequireSuccess(_tableCommands.Create(batch, source, "QualifiedSales", "A1:D6", true));
        foreach (var (sheet, name) in new[] { (target, "First"), (moved, "Second") })
            RequireSuccess(table
                ? _pivotCommands.CreateFromTable(batch, "QualifiedSales", sheet, "A1", name)
                : _pivotCommands.CreateFromRange(batch, source, "A1:D6", sheet, "A1", name));
        ReadNativePivot(target, "First", pivot =>
        {
            Excel.PivotCache? cache = null;
            try
            {
                cache = pivot.PivotCache();
                var quoted = source.Replace("'", "''", StringComparison.Ordinal);
                cache.SourceData = table ? $"'{quoted}'!QualifiedSales" :
                    $"'[{System.IO.Path.GetFileName(_fixture.WorkbookPath)}]{quoted}'!R1C1:R6C4";
                return 0;
            }
            finally { ComUtilities.Release(ref cache); }
        });

        Assert.True(ShareExisting(target, moved).Changed);
        Assert.Equal(ReadNativePivot(target, "First", pivot => pivot.CacheIndex),
            ReadNativePivot(moved, "Second", pivot => pivot.CacheIndex));
    }
}
