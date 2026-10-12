using Sbroenne.ExcelMcp.Core.Commands.PivotTable;
using Xunit;

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
    }
}
