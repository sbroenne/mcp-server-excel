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
    public void ReadAndListFields_PreserveRealColumnsMatchingTheValuesMarker(bool multipleValues)
    {
        var batch = _fixture.BatchToken;
        var source = _fixture.CreateTestSheet(batch);
        RequireSuccess(_commands.SetNumberFormat(batch, source, "D1", "@"));
        RequireSuccess(_commands.SetValues(batch, source, "A1:E5",
            [
                ["Region", "Sales", "Quantity", "-2146826246", "Values"],
                ["North", 17, 1, "One", "First"],
                ["South", 29, 2, "Two", "Second"],
                ["North", 31, 3, "Three", "Third"],
                ["South", 43, 4, "Four", "Fourth"]
            ], overwritePolicy: OverwritePolicy.Allow));
        RequireSuccess(_pivotCommands.CreateFromRange(batch, source, "A1:E5", source, "G1", "MarkerPivot"));
        RequireSuccess(_pivotCommands.AddRowField(batch, "MarkerPivot", "Region"));
        if (multipleValues)
        {
            RequireSuccess(_pivotCommands.AddValueField(batch, "MarkerPivot", "Sales"));
            RequireSuccess(_pivotCommands.AddValueField(batch, "MarkerPivot", "Quantity"));
        }

        var listed = RequireSuccess(_pivotCommands.ListFields(batch, "MarkerPivot"));
        var snapshot = RequireSuccess(_pivotCommands.Read(batch, "MarkerPivot", source));

        // The version-4 cache used by create-from-range reserves Values and exposes the source field as Values2.
        string[] expected = ["Region", "Sales", "Quantity", "-2146826246", "Values2"];
        Assert.Equal(expected.Order(StringComparer.Ordinal), listed.Fields.Select(field => field.Name).Order(StringComparer.Ordinal));
        Assert.Equal(expected.Order(StringComparer.Ordinal), snapshot.Fields.Select(field => field.Name).Order(StringComparer.Ordinal));
        Assert.True(snapshot.SnapshotComplete, System.Text.Json.JsonSerializer.Serialize(snapshot));
        Assert.Equal("Region", Assert.Single(snapshot.FieldSelections).FieldName);
        Assert.Equal(multipleValues ? 2 : 0, snapshot.ValueFields.Count);
        var unchanged = RequireSuccess(_commands.GetValues(batch, source, "A1:E5"));
        Assert.Equal(["Region", "Sales", "Quantity", "-2146826246", "Values"], unchanged.Values[0]);
        Assert.Equal(["South", 43, 4, "Four", "Fourth"], unchanged.Values[4]);
    }

    [Fact]
    public void Read_ReturnsBoundedSemanticSelectionsForPlacedRegularFields()
    {
        var batch = _fixture.BatchToken;
        var pivotName = $"Snapshot_{Guid.NewGuid():N}"[..24];
        RequireSuccess(_pivotCommands.CreateFromRange(
            batch, _salesSheetName, "A1:D6", _salesSheetName, "F2", pivotName));
        RequireSuccess(_pivotCommands.AddRowField(batch, pivotName, "Region"));
        RequireSuccess(_pivotCommands.AddRowField(batch, pivotName, "Product"));
        RequireSuccess(_pivotCommands.AddValueField(batch, pivotName, "Sales"));
        RequireSuccess(_pivotCommands.SetFieldFilter(batch, pivotName, "Region", ["North"]));
        RequireSuccess(_pivotCommands.SetFieldFilter(batch, pivotName, "Product", ["Widget"]));
        RequireSuccess(_pivotCommands.AddFieldFilter(batch, pivotName, "Product",
            new PivotFilterOptions { Type = PivotFilterType.CaptionEquals, Text1 = "Widget" }));

        var snapshot = RequireSuccess(_pivotCommands.Read(
            batch, pivotName, _salesSheetName, maxSelectedMembers: 1));

        Assert.False(snapshot.IsOlap);
        Assert.True(snapshot.IsOlapAvailable);
        Assert.True(snapshot.FieldSelectionsAvailable, snapshot.FieldSelectionsError);
        Assert.True(snapshot.ValueFieldsAvailable);
        Assert.Equal(1, snapshot.MaxSelectedMembers);
        Assert.False(snapshot.SnapshotComplete);

        var product = Assert.Single(snapshot.FieldSelections,
            selection => selection.FieldName == "Product");
        var region = Assert.Single(snapshot.FieldSelections,
            selection => selection.FieldName == "Region");
        Assert.True(product.SelectionStateAvailable);
        Assert.False(product.ShowAll);
        Assert.Equal(["Widget"], product.SelectedMemberUniqueNames);
        Assert.Equal(1, product.SelectedMemberCount);
        Assert.False(product.MembersTruncated);
        Assert.True(product.FiltersAvailable);
        var filter = Assert.Single(product.Filters);
        Assert.Equal("CaptionEquals", filter.Type);
        Assert.Equal("Widget", filter.Value1);

        Assert.True(region.SelectionStateAvailable);
        Assert.False(region.ShowAll);
        Assert.Empty(region.SelectedMemberUniqueNames);
        Assert.Equal(1, region.SelectedMemberCount);
        Assert.True(region.MembersTruncated);
        Assert.True(region.FiltersAvailable);
        AssertSumOfSalesValueField(snapshot);
    }

    [Fact]
    public void Read_UnfilteredRowAndSingleReportFilterReturnCompleteSnapshot()
    {
        var batch = _fixture.BatchToken;
        var pivotName = $"SnapAll_{Guid.NewGuid():N}"[..24];
        RequireSuccess(_pivotCommands.CreateFromRange(
            batch, _salesSheetName, "A1:D6", _salesSheetName, "F2", pivotName));
        RequireSuccess(_pivotCommands.AddRowField(batch, pivotName, "Region"));
        RequireSuccess(_pivotCommands.AddFilterField(batch, pivotName, "Product"));
        RequireSuccess(_pivotCommands.AddValueField(batch, pivotName, "Sales"));
        SetRegularReportFilter(pivotName, "Product", multiple: false, ["Widget"]);

        var snapshot = RequireSuccess(_pivotCommands.Read(batch, pivotName, _salesSheetName));

        Assert.True(snapshot.FieldSelectionsAvailable);
        Assert.True(snapshot.ValueFieldsAvailable);
        Assert.Equal(100, snapshot.MaxSelectedMembers);
        Assert.True(snapshot.SnapshotComplete);
        Assert.Equal(["Product", "Region"], snapshot.FieldSelections.Select(selection => selection.FieldName));

        var region = snapshot.FieldSelections[1];
        Assert.Equal(PivotFieldArea.Row, region.Area);
        Assert.Equal(1, region.Position);
        Assert.True(region.SelectionStateAvailable);
        Assert.True(region.ShowAll);
        Assert.Null(region.MultipleSelectionEnabled);
        Assert.Empty(region.SelectedMemberUniqueNames);
        Assert.Null(region.SelectedMemberCount);
        Assert.False(region.MembersTruncated);
        Assert.True(region.FiltersAvailable);
        Assert.Empty(region.Filters);

        var product = snapshot.FieldSelections[0];
        Assert.Equal(PivotFieldArea.Filter, product.Area);
        Assert.Equal(1, product.Position);
        Assert.True(product.SelectionStateAvailable);
        Assert.False(product.ShowAll);
        Assert.False(product.MultipleSelectionEnabled);
        Assert.Equal(["Widget"], product.SelectedMemberUniqueNames);
        Assert.Equal(1, product.SelectedMemberCount);
        Assert.False(product.MembersTruncated);
        Assert.True(product.FiltersAvailable);
        Assert.Empty(product.Filters);
        AssertSumOfSalesValueField(snapshot);
    }

    [Fact]
    public void Read_MultiSelectReportFilterReturnsExactVisibleMembersAndAllState()
    {
        var batch = _fixture.BatchToken;
        var pivotName = $"SnapMulti_{Guid.NewGuid():N}"[..24];
        RequireSuccess(_pivotCommands.CreateFromRange(
            batch, _salesSheetName, "A1:D6", _salesSheetName, "F2", pivotName));
        RequireSuccess(_pivotCommands.AddFilterField(batch, pivotName, "Product"));
        RequireSuccess(_pivotCommands.AddValueField(batch, pivotName, "Sales"));
        SetRegularReportFilter(pivotName, "Product", multiple: true, ["Gadget"]);

        var filtered = RequireSuccess(_pivotCommands.Read(batch, pivotName, _salesSheetName));

        Assert.True(filtered.SnapshotComplete);
        var product = Assert.Single(filtered.FieldSelections);
        Assert.Equal("Product", product.FieldName);
        Assert.True(product.SelectionStateAvailable);
        Assert.False(product.ShowAll);
        Assert.True(product.MultipleSelectionEnabled);
        Assert.Equal(["Gadget"], product.SelectedMemberUniqueNames);
        Assert.Equal(1, product.SelectedMemberCount);

        SetRegularReportFilter(pivotName, "Product", multiple: true, []);
        var all = RequireSuccess(_pivotCommands.Read(batch, pivotName, _salesSheetName));

        Assert.True(all.SnapshotComplete);
        product = Assert.Single(all.FieldSelections);
        Assert.True(product.ShowAll);
        Assert.True(product.MultipleSelectionEnabled);
        Assert.Empty(product.SelectedMemberUniqueNames);
        Assert.Null(product.SelectedMemberCount);
    }

    [Fact]
    public void Read_MultipleValueFieldsExcludeValuesPlaceholderAndReturnCompleteSnapshot()
    {
        var batch = _fixture.BatchToken;
        var pivotName = $"SnapVals_{Guid.NewGuid():N}"[..24];
        RequireSuccess(_pivotCommands.CreateFromRange(
            batch, _salesSheetName, "A1:D6", _salesSheetName, "F2", pivotName));
        RequireSuccess(_pivotCommands.AddColumnField(batch, pivotName, "Region"));
        RequireSuccess(_pivotCommands.AddValueField(batch, pivotName, "Sales"));
        RequireSuccess(_pivotCommands.AddValueField(
            batch, pivotName, "Sales", AggregationFunction.Count, "Count of Sales"));

        var snapshot = RequireSuccess(_pivotCommands.Read(batch, pivotName, _salesSheetName));

        Assert.True(snapshot.FieldSelectionsAvailable, System.Text.Json.JsonSerializer.Serialize(snapshot));
        Assert.True(snapshot.SnapshotComplete, System.Text.Json.JsonSerializer.Serialize(snapshot));
        var region = Assert.Single(snapshot.FieldSelections);
        Assert.Equal("Region", region.FieldName);
        Assert.Equal(PivotFieldArea.Column, region.Area);
        Assert.Equal(1, region.Position);
        Assert.True(region.SelectionStateAvailable);
        Assert.True(region.ShowAll);
        Assert.True(region.FiltersAvailable);
        Assert.Empty(region.Filters);
        Assert.True(snapshot.ValueFieldsAvailable);
        Assert.Equal(["Sum of Sales", "Count of Sales"], snapshot.ValueFields.Select(value => value.FieldName));
        Assert.Equal([AggregationFunction.Sum, AggregationFunction.Count],
            snapshot.ValueFields.Select(value => value.Function));
        Assert.Equal(["Date", "Product", "Region", "Sales"],
            snapshot.Fields.Select(field => field.Name).Order(StringComparer.Ordinal));
    }

    [Fact]
    public void Read_SheetNameSelectsSameNamedPivotOnRequestedSheet()
    {
        var batch = _fixture.BatchToken;
        var pivotName = $"SnapDup_{Guid.NewGuid():N}"[..24];
        var otherSheetName = $"Dup_{Guid.NewGuid():N}"[..31];
        RequireSuccess(_pivotCommands.CreateFromRange(
            batch, _salesSheetName, "A1:D6", _salesSheetName, "F2", pivotName));
        RequireSuccess(_pivotCommands.AddRowField(batch, pivotName, "Region"));
        _fixture.CreateNamedTestSheet(batch, otherSheetName);
        CreateSameNamedPivotWithProductColumn(pivotName, otherSheetName);

        var other = RequireSuccess(_pivotCommands.Read(batch, pivotName, otherSheetName));
        var sales = RequireSuccess(_pivotCommands.Read(batch, pivotName, _salesSheetName));

        Assert.Equal(otherSheetName, other.PivotTable.SheetName);
        var product = Assert.Single(other.FieldSelections);
        Assert.Equal("Product", product.FieldName);
        Assert.Equal(PivotFieldArea.Column, product.Area);
        Assert.Equal(_salesSheetName, sales.PivotTable.SheetName);
        var region = Assert.Single(sales.FieldSelections);
        Assert.Equal("Region", region.FieldName);
        Assert.Equal(PivotFieldArea.Row, region.Area);
    }

    private static void AssertSumOfSalesValueField(PivotTableInfoResult snapshot)
    {
        var value = Assert.Single(snapshot.ValueFields);
        Assert.Equal("Sum of Sales", value.FieldName);
        Assert.Equal("Sales", value.SourceName);
        Assert.Equal(1, value.Position);
        Assert.Equal(AggregationFunction.Sum, value.Function);
        Assert.Equal(PivotFieldCalculation.Normal, value.Calculation);
        Assert.False(value.IsOlap);
        Assert.Null(value.BaseFieldName);
        Assert.Null(value.BaseItemKind);
        Assert.Null(value.BaseItemName);
    }

    private void CreateSameNamedPivotWithProductColumn(string pivotName, string targetSheetName)
    {
        _fixture.ExecuteRawVerification((context, _) =>
        {
            Excel.Worksheet? sheet = null;
            Excel.PivotTables? pivots = null;
            Excel.PivotTable? source = null;
            Excel.PivotCache? cache = null;
            Excel.Worksheet? target = null;
            Excel.Range? destination = null;
            Excel.PivotTable? copy = null;
            Excel.PivotField? product = null;
            try
            {
                sheet = ComUtilities.FindSheet(context.Book, _salesSheetName);
                pivots = (Excel.PivotTables)sheet!.PivotTables();
                source = pivots.Item(pivotName);
                cache = source.PivotCache();
                target = ComUtilities.FindSheet(context.Book, targetSheetName);
                destination = target!.Range["A3"];
                copy = cache.CreatePivotTable(destination, pivotName);
                product = (Excel.PivotField)copy.PivotFields("Product");
                product.Orientation = Excel.XlPivotFieldOrientation.xlColumnField;
            }
            finally
            {
                ComUtilities.Release(ref product);
                ComUtilities.Release(ref copy);
                ComUtilities.Release(ref destination);
                ComUtilities.Release(ref target);
                ComUtilities.Release(ref cache);
                ComUtilities.Release(ref source);
                ComUtilities.Release(ref pivots);
                ComUtilities.Release(ref sheet);
            }
        });
    }

    private void SetRegularReportFilter(
        string pivotName, string fieldName, bool multiple, IReadOnlyCollection<string> visibleItems)
    {
        _fixture.ExecuteRawVerification((context, _) =>
        {
            Excel.Sheets? sheets = null;
            Excel.Worksheet? sheet = null;
            Excel.PivotTables? pivots = null;
            Excel.PivotTable? pivot = null;
            Excel.PivotField? field = null;
            Excel.PivotItems? items = null;
            try
            {
                sheets = context.Book.Worksheets;
                sheet = (Excel.Worksheet)sheets[_salesSheetName];
                pivots = (Excel.PivotTables)sheet.PivotTables();
                pivot = pivots.Item(pivotName);
                field = (Excel.PivotField)pivot.PivotFields(fieldName);
                field.ClearAllFilters();
                field.EnableMultiplePageItems = multiple;
                if (!multiple)
                {
                    field.CurrentPage = visibleItems.Single();
                    return;
                }

                items = (Excel.PivotItems)field.PivotItems();
                for (int index = 1; index <= items.Count; index++)
                {
                    Excel.PivotItem? item = null;
                    try
                    {
                        item = items.Item(index);
                        if (visibleItems.Count > 0 && !visibleItems.Contains(item.Name))
                            item.Visible = false;
                    }
                    finally
                    {
                        ComUtilities.Release(ref item);
                    }
                }
            }
            finally
            {
                ComUtilities.Release(ref items);
                ComUtilities.Release(ref field);
                ComUtilities.Release(ref pivot);
                ComUtilities.Release(ref pivots);
                ComUtilities.Release(ref sheet);
                ComUtilities.Release(ref sheets);
            }
        });
    }
}
