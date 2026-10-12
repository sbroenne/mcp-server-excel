using System.Runtime.InteropServices;
using Sbroenne.ExcelMcp.Core.Commands.PivotTable;
using Sbroenne.ExcelMcp.Core.Models;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Unit;

[Trait("RequiresExcel", "false")]
[Trait("Category", "Unit")]
[Trait("Feature", "PivotTables")]
[Trait("Layer", "Core")]
[Trait("Speed", "Fast")]
public sealed class PivotTableSemanticSnapshotTests
{
    [Fact]
    public void SingleSelection_ReturnsExactSemanticName()
    {
        var selection = PivotTableSnapshotHelpers.CreateFieldSelection(
            "[Region].[State]", PivotFieldArea.Filter, 1, ["[Region].[State].&[WA]"],
            showAll: false, multipleSelectionEnabled: false, maxMembers: 10);

        Assert.True(selection.SelectionStateAvailable);
        Assert.False(selection.ShowAll);
        Assert.False(selection.MultipleSelectionEnabled);
        Assert.Equal(["[Region].[State].&[WA]"], selection.SelectedMemberUniqueNames);
        Assert.Equal(1, selection.SelectedMemberCount);
        Assert.False(selection.MembersTruncated);
    }

    [Fact]
    public void MultipleSelection_IsSortedAndExplicitlyTruncatedAtTheBound()
    {
        var selection = PivotTableSnapshotHelpers.CreateFieldSelection(
            "[Region].[State]", PivotFieldArea.Filter, 1,
            ["[Region].[State].&[WA]", "[Region].[State].&[CA]", "[Region].[State].&[OR]"],
            showAll: false, multipleSelectionEnabled: true, maxMembers: 2);

        Assert.True(selection.SelectionStateAvailable);
        Assert.False(selection.ShowAll);
        Assert.True(selection.MultipleSelectionEnabled);
        Assert.Equal(["[Region].[State].&[CA]", "[Region].[State].&[OR]"],
            selection.SelectedMemberUniqueNames);
        Assert.Equal(3, selection.SelectedMemberCount);
        Assert.True(selection.MembersTruncated);
    }

    [Fact]
    public void AllSelection_HasNoSelectedMembersAndIsNotTruncated()
    {
        var selection = PivotTableSnapshotHelpers.CreateFieldSelection(
            "[Region].[State]", PivotFieldArea.Filter, 1, [],
            showAll: true, multipleSelectionEnabled: true, maxMembers: 1);

        Assert.True(selection.SelectionStateAvailable);
        Assert.True(selection.ShowAll);
        Assert.Empty(selection.SelectedMemberUniqueNames);
        Assert.Null(selection.SelectedMemberCount);
        Assert.False(selection.MembersTruncated);
    }

    [Fact]
    public void UnavailableGetter_IsNotReportedAsAllOrEmptySelection()
    {
#pragma warning disable CA2201 // Simulate an Excel COM getter failure without starting Excel.
        var read = PivotTableSnapshotHelpers.TryReadSnapshotProperty<string[]>(
            () => throw new COMException("Selection getter unavailable."));
#pragma warning restore CA2201

        Assert.False(read.Available);
        Assert.Null(read.Value);
        Assert.Equal("Selection getter unavailable.", read.ErrorMessage);

        var selection = PivotTableSnapshotHelpers.CreateFieldSelection(
            "[Region].[State]", PivotFieldArea.Filter, 1, read.Value ?? [],
            showAll: null, multipleSelectionEnabled: null, maxMembers: 10,
            selectionStateAvailable: read.Available, selectionError: read.ErrorMessage);

        Assert.False(selection.SelectionStateAvailable);
        Assert.Null(selection.ShowAll);
        Assert.Empty(selection.SelectedMemberUniqueNames);
        Assert.Null(selection.SelectedMemberCount);
        Assert.Equal("Selection getter unavailable.", selection.SelectionError);
    }

    [Fact]
    public void SemanticNames_AreStableAcrossCollectionOrderChanges()
    {
        var beforeSave = PivotTableSnapshotHelpers.CreateFieldSelection(
            "[Region].[State]", PivotFieldArea.Filter, 1,
            ["[Region].[State].&[WA]", "[Region].[State].&[CA]"],
            showAll: false, multipleSelectionEnabled: true, maxMembers: 10);
        var afterSave = PivotTableSnapshotHelpers.CreateFieldSelection(
            "[Region].[State]", PivotFieldArea.Filter, 1,
            ["[Region].[State].&[CA]", "[Region].[State].&[WA]"],
            showAll: false, multipleSelectionEnabled: true, maxMembers: 10);

        Assert.Equal(beforeSave.FieldName, afterSave.FieldName);
        Assert.Equal(beforeSave.SelectedMemberUniqueNames, afterSave.SelectedMemberUniqueNames);
    }

    [Theory]
    [InlineData(0)]
    [InlineData(1001)]
    public void MemberLimit_RejectsValuesOutsideSupportedRange(int maxMembers)
    {
        Assert.Throws<ArgumentOutOfRangeException>(() =>
            PivotTableSnapshotHelpers.ValidateMaxMembers(maxMembers));
    }
}
