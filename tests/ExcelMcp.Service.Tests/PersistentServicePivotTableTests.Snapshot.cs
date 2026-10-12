using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.Core.Commands.PivotTable;
using Sbroenne.ExcelMcp.Core.Models;
using Xunit;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sbroenne.ExcelMcp.Service.Tests;

public sealed partial class PersistentServicePivotTableTests
{
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
        Assert.True(snapshot.FieldSelectionsAvailable);
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
        Assert.False(snapshot.ManualUpdate);
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
