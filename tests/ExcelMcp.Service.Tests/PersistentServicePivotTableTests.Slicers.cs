using System.Globalization;
using Sbroenne.ExcelMcp.Core.Models;
using Xunit;

namespace Sbroenne.ExcelMcp.Service.Tests;

public sealed partial class PersistentServicePivotTableTests
{
    [Fact]
    [Trait("Speed", "Medium")]
    public void Slicer_NativeReposition_MatchesCoordinatesAndReportedAnchor()
    {
        var batch = _fixture.BatchToken;
        RequireSuccess(_pivotCommands.CreateFromRange(
            batch, _salesSheetName, "A1:D6", _salesSheetName, "F2", "PositionControl"));
        RequireSuccess(_pivotCommands.AddRowField(batch, "PositionControl", "Region"));
        var result = RequireSuccess(_pivotCommands.CreateSlicer(
            batch, "PositionControl", "Region", "PositionControlSlicer", _salesSheetName, "I2"));
        _fixture.ExecuteRawVerification((context, _) =>
        {
            Microsoft.Office.Interop.Excel.Worksheet? sheet = null;
            Microsoft.Office.Interop.Excel.Range? target = null;
            Microsoft.Office.Interop.Excel.SlicerCaches? caches = null;
            Microsoft.Office.Interop.Excel.SlicerCache? cache = null;
            Microsoft.Office.Interop.Excel.Slicers? slicers = null;
            Microsoft.Office.Interop.Excel.Slicer? slicer = null;
            Microsoft.Office.Interop.Excel.Shape? shape = null;
            Microsoft.Office.Interop.Excel.Range? anchor = null;
            try
            {
                sheet = Sbroenne.ExcelMcp.ComInterop.ComUtilities.FindSheet(context.Book, _salesSheetName);
                Assert.NotNull(sheet);
                target = sheet.Range["I2"];
                caches = context.Book.SlicerCaches;
                cache = caches.Item[1];
                slicers = cache.Slicers;
                slicer = slicers.Item["PositionControlSlicer"];
                slicer.Left = Convert.ToDouble(target.Left);
                slicer.Top = Convert.ToDouble(target.Top);
                shape = slicer.Shape;
                anchor = shape.TopLeftCell;
                Assert.Equal(Convert.ToDouble(target.Left), shape.Left, precision: 3);
                Assert.Equal(Convert.ToDouble(target.Top), shape.Top, precision: 3);
                Assert.Equal(result.Position, anchor.Address[false, false]);
            }
            finally
            {
                Sbroenne.ExcelMcp.ComInterop.ComUtilities.Release(ref anchor);
                Sbroenne.ExcelMcp.ComInterop.ComUtilities.Release(ref shape);
                Sbroenne.ExcelMcp.ComInterop.ComUtilities.Release(ref slicer);
                Sbroenne.ExcelMcp.ComInterop.ComUtilities.Release(ref slicers);
                Sbroenne.ExcelMcp.ComInterop.ComUtilities.Release(ref cache);
                Sbroenne.ExcelMcp.ComInterop.ComUtilities.Release(ref caches);
                Sbroenne.ExcelMcp.ComInterop.ComUtilities.Release(ref target);
                Sbroenne.ExcelMcp.ComInterop.ComUtilities.Release(ref sheet);
            }
        });
        AssertOriginalSales();
    }

    /// <summary>
    /// Tests creating a slicer for a PivotTable field.
    /// </summary>
    [Fact]
    [Trait("Speed", "Medium")]
    public void CreateSlicer_ValidField_CreatesSlicerSuccessfully()
    {
        // Arrange
        var batch = _fixture.BatchToken;

        // Create PivotTable
        var createResult = _pivotCommands.CreateFromRange(
            batch, _salesSheetName, "A1:D6", _salesSheetName, "F2", "SlicerTest");
        RequireSuccess(createResult);

        // Add Region to Row area
        var addFieldResult = _pivotCommands.AddRowField(batch, "SlicerTest", "Region");
        RequireSuccess(addFieldResult);

        // Act - Create slicer for Region field
        var slicerResult = _pivotCommands.CreateSlicer(
            batch,
            pivotTableName: "SlicerTest",
            fieldName: "Region",
            slicerName: "RegionSlicer",
            destinationSheet: _salesSheetName,
            position: "I2");

        // Assert
        RequireSuccess(slicerResult);
        Assert.Equal("RegionSlicer", slicerResult.Name);
        Assert.Equal("Region", slicerResult.FieldName);
        Assert.Equal(_salesSheetName, slicerResult.SheetName);
        Assert.NotNull(slicerResult.AvailableItems);
        Assert.Contains("North", slicerResult.AvailableItems);
        Assert.Contains("South", slicerResult.AvailableItems);
        Assert.NotNull(slicerResult.WorkflowHint);
        RequireSuccess(slicerResult);
        Assert.Equal("PivotTable", slicerResult.SourceType);
        AssertPivotSlicer("RegionSlicer", "SlicerTest", "Region", "I2", ["North", "South"]);
        AssertOriginalSales();
    }

    /// <summary>
    /// Tests listing slicers in a workbook with no filter.
    /// </summary>
    [Fact]
    [Trait("Speed", "Medium")]
    public void ListSlicers_WithSlicers_ReturnsAllSlicers()
    {
        // Arrange
        var batch = _fixture.BatchToken;

        // Create PivotTable
        var createResult = _pivotCommands.CreateFromRange(
            batch, _salesSheetName, "A1:D6", _salesSheetName, "F2", "ListSlicersTest");
        RequireSuccess(createResult);

        // Add fields
        RequireSuccess(_pivotCommands.AddRowField(batch, "ListSlicersTest", "Region"));
        RequireSuccess(_pivotCommands.AddRowField(batch, "ListSlicersTest", "Product"));

        // Create two slicers
        var slicer1Result = _pivotCommands.CreateSlicer(
            batch, "ListSlicersTest", "Region", "RegionSlicer1", _salesSheetName, "I2");
        RequireSuccess(slicer1Result);

        var slicer2Result = _pivotCommands.CreateSlicer(
            batch, "ListSlicersTest", "Product", "ProductSlicer1", _salesSheetName, "I10");
        RequireSuccess(slicer2Result);

        // Act
        var listResult = _pivotCommands.ListSlicers(batch);

        // Assert
        RequireSuccess(listResult);
        Assert.NotNull(listResult.Slicers);
        Assert.Equal(2, listResult.Slicers.Count);
        Assert.Contains(listResult.Slicers, s => s.Name == "RegionSlicer1");
        Assert.Contains(listResult.Slicers, s => s.Name == "ProductSlicer1");
        RequireSuccess(listResult);
        AssertPivotSlicer("RegionSlicer1", "ListSlicersTest", "Region", "I2", ["North", "South"]);
        AssertPivotSlicer("ProductSlicer1", "ListSlicersTest", "Product", "I10", ["Gadget", "Widget"]);
    }

    /// <summary>
    /// Tests listing slicers filtered by PivotTable name.
    /// </summary>
    [Fact]
    [Trait("Speed", "Medium")]
    public void ListSlicers_FilterByPivotTable_ReturnsConnectedSlicersOnly()
    {
        // Arrange
        var batch = _fixture.BatchToken;

        // Create PivotTable
        var createResult = _pivotCommands.CreateFromRange(
            batch, _salesSheetName, "A1:D6", _salesSheetName, "F2", "FilterSlicersTest");
        RequireSuccess(createResult);

        // Add Region field and create slicer
        RequireSuccess(_pivotCommands.AddRowField(batch, "FilterSlicersTest", "Region"));
        var slicerResult = _pivotCommands.CreateSlicer(
            batch, "FilterSlicersTest", "Region", "FilterRegionSlicer", _salesSheetName, "I2");
        RequireSuccess(slicerResult);
        var otherSheet = _fixture.CreateTestSheet(batch);
        RequireSuccess(_pivotCommands.CreateFromRange(
            batch, _salesSheetName, "A1:D6", otherSheet, "A1", "OtherPivot"));
        RequireSuccess(_pivotCommands.AddRowField(batch, "OtherPivot", "Region"));
        RequireSuccess(_pivotCommands.CreateSlicer(
            batch, "OtherPivot", "Region", "OtherSlicer", otherSheet, "I2"));

        // Act
        var listResult = _pivotCommands.ListSlicers(batch, pivotTableName: "FilterSlicersTest");

        // Assert
        RequireSuccess(listResult);
        Assert.NotNull(listResult.Slicers);
        Assert.Single(listResult.Slicers);
        Assert.Equal("FilterRegionSlicer", listResult.Slicers[0].Name);
        RequireSuccess(listResult);
        Assert.Equal(2, RequireSuccess(_pivotCommands.ListSlicers(batch)).Slicers.Count);
        Assert.Equal("OtherSlicer", Assert.Single(
            RequireSuccess(_pivotCommands.ListSlicers(batch, "OtherPivot")).Slicers).Name);
    }

    /// <summary>
    /// Tests setting slicer selection to specific items.
    /// </summary>
    [Fact]
    [Trait("Speed", "Medium")]
    public void SetSlicerSelection_SpecificItems_SelectsOnlyThoseItems()
    {
        // Arrange
        var batch = _fixture.BatchToken;

        // Create PivotTable with Region field
        var createResult = _pivotCommands.CreateFromRange(
            batch, _salesSheetName, "A1:D6", _salesSheetName, "F2", "SelectionTest");
        RequireSuccess(createResult);

        RequireSuccess(_pivotCommands.AddRowField(batch, "SelectionTest", "Region"));
        RequireSuccess(_pivotCommands.AddValueField(batch, "SelectionTest", "Sales"));
        AssertPivotSales(325, 325, "SelectionTest");

        // Create slicer
        var slicerResult = _pivotCommands.CreateSlicer(
            batch, "SelectionTest", "Region", "SelectionSlicer", _salesSheetName, "I2");
        RequireSuccess(slicerResult);

        // Act - Select only "North"
        var selectionResult = _pivotCommands.SetSlicerSelection(
            batch, "SelectionSlicer", new List<string> { "North" }, clearFirst: true);

        // Assert
        RequireSuccess(selectionResult);
        Assert.NotNull(selectionResult.SelectedItems);
        Assert.Single(selectionResult.SelectedItems);
        Assert.Contains("North", selectionResult.SelectedItems);
        Assert.DoesNotContain("South", selectionResult.SelectedItems);
        Assert.NotNull(selectionResult.WorkflowHint);
        RequireSuccess(selectionResult);
        Assert.Equal("PivotTable", selectionResult.SourceType);
        AssertPivotSlicer("SelectionSlicer", "SelectionTest", "Region", "I2", ["North"]);
        AssertFilteredPivot("SelectionTest", "North", 325);
        AssertOriginalSales();
    }

    /// <summary>
    /// Tests clearing slicer selection (selecting all items).
    /// </summary>
    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    [Trait("Speed", "Medium")]
    public void SetSlicerSelection_EmptyList_ClearsFilterSelectsAll(bool clearFirst)
    {
        // Arrange
        var batch = _fixture.BatchToken;

        // Create PivotTable
        var createResult = _pivotCommands.CreateFromRange(
            batch, _salesSheetName, "A1:D6", _salesSheetName, "F2", "ClearFilterTest");
        RequireSuccess(createResult);

        RequireSuccess(_pivotCommands.AddRowField(batch, "ClearFilterTest", "Region"));
        RequireSuccess(_pivotCommands.AddValueField(batch, "ClearFilterTest", "Sales"));

        // Create slicer
        var slicerResult = _pivotCommands.CreateSlicer(
            batch, "ClearFilterTest", "Region", "ClearFilterSlicer", _salesSheetName, "I2");
        RequireSuccess(slicerResult);

        // First, filter to just "North"
        var filtered = _pivotCommands.SetSlicerSelection(batch, "ClearFilterSlicer", ["North"]);
        AssertRegularSlicerState(filtered, 325, "North");
        AssertFilteredPivot("ClearFilterTest", "North", 325);

        // Act - Clear filter by passing empty list
        var clearResult = _pivotCommands.SetSlicerSelection(
            batch, "ClearFilterSlicer", [], clearFirst);

        // Assert
        AssertRegularSlicerState(clearResult, 650, "North", "South");
        Assert.Contains("cleared", clearResult.WorkflowHint, StringComparison.OrdinalIgnoreCase);
        RequireSuccess(clearResult);
        Assert.Equal("PivotTable", clearResult.SourceType);
        AssertPivotSlicer("ClearFilterSlicer", "ClearFilterTest", "Region", "I2", ["North", "South"]);
        AssertPivotSales(325, 325, "ClearFilterTest");
        AssertOriginalSales();
    }

    [Fact]
    [Trait("Speed", "Medium")]
    public void SetSlicerSelection_ClearFirstFalse_AddsWithoutReplacingAndPreservesUnfilteredState()
    {
        var batch = _fixture.BatchToken;
        var created = _pivotCommands.CreateFromRange(
            batch, _salesSheetName, "A1:D6", _salesSheetName, "F2", "AddSelectionTest");
        RequireSuccess(created);
        RequireSuccess(_pivotCommands.AddRowField(batch, "AddSelectionTest", "Region"));
        RequireSuccess(_pivotCommands.AddValueField(batch, "AddSelectionTest", "Sales"));
        var slicer = _pivotCommands.CreateSlicer(
            batch, "AddSelectionTest", "Region", "AddSelectionSlicer", _salesSheetName, "I2");
        AssertRegularSlicerState(slicer, 650, "North", "South");
        var unfiltered = _pivotCommands.SetSlicerSelection(
            batch, "AddSelectionSlicer", ["North"], clearFirst: false);
        AssertRegularSlicerState(unfiltered, 650, "North", "South");
        var filtered = _pivotCommands.SetSlicerSelection(batch, "AddSelectionSlicer", ["North"]);
        AssertRegularSlicerState(filtered, 325, "North");
        var added = _pivotCommands.SetSlicerSelection(
            batch, "AddSelectionSlicer", ["South"], clearFirst: false);
        AssertRegularSlicerState(added, 650, "North", "South");
        var replaced = _pivotCommands.SetSlicerSelection(
            batch, "AddSelectionSlicer", ["South"], clearFirst: true);
        AssertRegularSlicerState(replaced, 325, "South");
    }

    private void AssertRegularSlicerState(SlicerResult result, double total, params string[] regions)
    {
        RequireSuccess(result);
        Assert.True(string.IsNullOrEmpty(result.ErrorMessage));
        Assert.Equal(["North", "South"], result.AvailableItems.Order());
        Assert.Equal(regions.Order(), result.SelectedItems.Order());
        var listed = _pivotCommands.ListSlicers(_fixture.BatchToken);
        RequireSuccess(listed);
        Assert.Equal(regions.Order(), Assert.Single(listed.Slicers).SelectedItems.Order());
        var values = _commands.GetValues(_fixture.BatchToken, _salesSheetName, "F2:G6");
        RequireSuccess(values);
        var rows = values.Values.Where(row => row[0] is not null
            && double.TryParse(row[1]?.ToString(), NumberStyles.Float, CultureInfo.InvariantCulture, out _)).ToList();
        Assert.Equal(regions.Length + 1, rows.Count);
        Assert.Equal(regions.Order(), rows.Take(regions.Length).Select(row => row[0]!.ToString()).Order());
        Assert.All(rows.Take(regions.Length), row =>
            Assert.Equal(325, double.Parse(row[1]!.ToString()!, CultureInfo.InvariantCulture)));
        Assert.Equal(total, double.Parse(rows[^1][1]!.ToString()!, CultureInfo.InvariantCulture));
    }

    [Theory]
    [InlineData("North", "South")]
    [InlineData("South", "North")]
    [Trait("Speed", "Medium")]
    public void SetSlicerSelection_ReplaceSoleSelectedItem_SelectsOnlyReplacement(
        string initial, string replacement)
    {
        var batch = _fixture.BatchToken;
        var created = _pivotCommands.CreateFromRange(
            batch, _salesSheetName, "A1:D6", _salesSheetName, "F2", "ReplaceSoleItemPivot");
        RequireSuccess(created);
        RequireSuccess(_pivotCommands.AddRowField(batch, "ReplaceSoleItemPivot", "Region"));
        RequireSuccess(_pivotCommands.AddValueField(batch, "ReplaceSoleItemPivot", "Sales"));
        var slicer = _pivotCommands.CreateSlicer(batch, "ReplaceSoleItemPivot", "Region",
            "ReplaceSoleItemSlicer", _salesSheetName, "I2");
        AssertRegularSlicerState(slicer, 650, "North", "South");

        AssertRegularSlicerState(_pivotCommands.SetSlicerSelection(
            batch, "ReplaceSoleItemSlicer", [initial]), 325, initial);
        AssertRegularSlicerState(_pivotCommands.SetSlicerSelection(
            batch, "ReplaceSoleItemSlicer", [replacement], clearFirst: true), 325, replacement);
        AssertRegularSlicerState(_pivotCommands.SetSlicerSelection(
            batch, "ReplaceSoleItemSlicer", [initial], clearFirst: true), 325, initial);
        AssertRegularSlicerState(_pivotCommands.SetSlicerSelection(
            batch, "ReplaceSoleItemSlicer", [replacement], clearFirst: false), 650, "North", "South");
        AssertRegularSlicerState(_pivotCommands.SetSlicerSelection(
            batch, "ReplaceSoleItemSlicer", [replacement], clearFirst: true), 325, replacement);
        AssertRegularSlicerState(_pivotCommands.SetSlicerSelection(
            batch, "ReplaceSoleItemSlicer", [initial.ToLowerInvariant(), initial, "missing"],
            clearFirst: true), 325, initial);
        AssertRegularSlicerState(_pivotCommands.SetSlicerSelection(
            batch, "ReplaceSoleItemSlicer", []), 650, "North", "South");
    }

    /// <summary>
    /// Tests deleting a slicer from the workbook.
    /// </summary>
    [Fact]
    [Trait("Speed", "Medium")]
    public void DeleteSlicer_ExistingSlicer_RemovesSuccessfully()
    {
        // Arrange
        var batch = _fixture.BatchToken;

        // Create PivotTable
        var createResult = _pivotCommands.CreateFromRange(
            batch, _salesSheetName, "A1:D6", _salesSheetName, "F2", "DeleteSlicerTest");
        RequireSuccess(createResult);

        RequireSuccess(_pivotCommands.AddRowField(batch, "DeleteSlicerTest", "Region"));
        RequireSuccess(_pivotCommands.AddValueField(batch, "DeleteSlicerTest", "Sales"));

        // Create slicer
        var slicerResult = _pivotCommands.CreateSlicer(
            batch, "DeleteSlicerTest", "Region", "SlicerToDelete", _salesSheetName, "I2");
        RequireSuccess(slicerResult);

        // Verify slicer exists
        RequireSuccess(_pivotCommands.SetSlicerSelection(batch, "SlicerToDelete", ["South"]));
        AssertFilteredPivot("DeleteSlicerTest", "South", 325);
        var listBeforeResult = _pivotCommands.ListSlicers(batch);
        Assert.Contains(listBeforeResult.Slicers, s => s.Name == "SlicerToDelete");
        RequireSuccess(listBeforeResult);
        var pivotBefore = SnapshotPivot("DeleteSlicerTest");

        // Act
        var deleteResult = _pivotCommands.DeleteSlicer(batch, "SlicerToDelete");

        // Assert
        RequireSuccess(deleteResult);

        // Verify slicer is gone
        var listAfterResult = _pivotCommands.ListSlicers(batch);
        Assert.DoesNotContain(listAfterResult.Slicers, s => s.Name == "SlicerToDelete");
        RequireSuccess(deleteResult);
        RequireSuccess(listAfterResult);
        Assert.Empty(listAfterResult.Slicers);
        Assert.Equal(pivotBefore, SnapshotPivot("DeleteSlicerTest"));
        AssertFilteredPivot("DeleteSlicerTest", "South", 325);
        AssertOriginalSales();
    }

    /// <summary>
    /// Tests deleting a non-existent slicer returns error.
    /// </summary>
    [Fact]
    [Trait("Speed", "Medium")]
    public void DeleteSlicer_NonExistentSlicer_ReturnsError()
    {
        // Arrange
        var batch = _fixture.BatchToken;

        // Create PivotTable (no slicer)
        var createResult = _pivotCommands.CreateFromRange(
            batch, _salesSheetName, "A1:D6", _salesSheetName, "F2", "NoSlicerTest");
        RequireSuccess(createResult);
        RequireSuccess(_pivotCommands.AddRowField(batch, "NoSlicerTest", "Region"));
        RequireSuccess(_pivotCommands.AddValueField(batch, "NoSlicerTest", "Sales"));
        RequireSuccess(_pivotCommands.CreateSlicer(
            batch, "NoSlicerTest", "Region", "RetainedSlicer", _salesSheetName, "I2"));
        RequireSuccess(_pivotCommands.SetSlicerSelection(batch, "RetainedSlicer", ["South"]));
        AssertFilteredPivot("NoSlicerTest", "South", 325);

        // Act
        var deleteResult = _pivotCommands.DeleteSlicer(batch, "NonExistentSlicer");

        // Assert
        Assert.False(deleteResult.Success);
        Assert.Contains("not found", deleteResult.ErrorMessage, StringComparison.OrdinalIgnoreCase);
        AssertPivotSlicer("RetainedSlicer", "NoSlicerTest", "Region", "I2", ["South"]);
        AssertFilteredPivot("NoSlicerTest", "South", 325);
        AssertOriginalSales();
    }

    /// <summary>
    /// Tests setting slicer selection for non-existent slicer returns error.
    /// </summary>
    [Fact]
    [Trait("Speed", "Medium")]
    public void SetSlicerSelection_NonExistentSlicer_ReturnsError()
    {
        // Arrange
        var batch = _fixture.BatchToken;

        // Create PivotTable (no slicer)
        var createResult = _pivotCommands.CreateFromRange(
            batch, _salesSheetName, "A1:D6", _salesSheetName, "F2", "NoSlicerTest2");
        RequireSuccess(createResult);
        RequireSuccess(_pivotCommands.AddRowField(batch, "NoSlicerTest2", "Region"));
        RequireSuccess(_pivotCommands.AddValueField(batch, "NoSlicerTest2", "Sales"));
        RequireSuccess(_pivotCommands.CreateSlicer(
            batch, "NoSlicerTest2", "Region", "RetainedSlicer", _salesSheetName, "I2"));
        RequireSuccess(_pivotCommands.SetSlicerSelection(batch, "RetainedSlicer", ["South"]));
        AssertFilteredPivot("NoSlicerTest2", "South", 325);

        // Act
        var selectionResult = _pivotCommands.SetSlicerSelection(
            batch, "NonExistentSlicer", new List<string> { "North" });

        // Assert
        Assert.False(selectionResult.Success);
        Assert.Contains("not found", selectionResult.ErrorMessage, StringComparison.OrdinalIgnoreCase);
        AssertPivotSlicer("RetainedSlicer", "NoSlicerTest2", "Region", "I2", ["South"]);
        AssertFilteredPivot("NoSlicerTest2", "South", 325);
        AssertOriginalSales();
    }

    /// <summary>
    /// Tests listing slicers when workbook has no slicers.
    /// </summary>
    [Fact]
    [Trait("Speed", "Medium")]
    public void ListSlicers_NoSlicers_ReturnsEmptyList()
    {
        // Arrange
        var batch = _fixture.BatchToken;

        // Create PivotTable without any slicers
        var createResult = _pivotCommands.CreateFromRange(
            batch, _salesSheetName, "A1:D6", _salesSheetName, "F2", "NoSlicerPivot");
        RequireSuccess(createResult);

        // Act
        var listResult = _pivotCommands.ListSlicers(batch);

        // Assert
        RequireSuccess(listResult);
        Assert.NotNull(listResult.Slicers);
        Assert.Empty(listResult.Slicers);
    }

    /// <summary>
    /// Tests that slicer shows connected PivotTables.
    /// </summary>
    [Fact]
    [Trait("Speed", "Medium")]
    public void CreateSlicer_ShowsConnectedPivotTable()
    {
        // Arrange
        var batch = _fixture.BatchToken;

        // Create PivotTable
        var createResult = _pivotCommands.CreateFromRange(
            batch, _salesSheetName, "A1:D6", _salesSheetName, "F2", "ConnectedPivot");
        RequireSuccess(createResult);

        RequireSuccess(_pivotCommands.AddRowField(batch, "ConnectedPivot", "Region"));

        // Act - Create slicer
        var slicerResult = _pivotCommands.CreateSlicer(
            batch, "ConnectedPivot", "Region", "ConnectedSlicer", _salesSheetName, "I2");

        // Assert
        RequireSuccess(slicerResult);
        Assert.NotNull(slicerResult.ConnectedPivotTables);
        Assert.Contains("ConnectedPivot", slicerResult.ConnectedPivotTables);
        AssertPivotSlicer("ConnectedSlicer", "ConnectedPivot", "Region", "I2", ["North", "South"]);
    }

    /// <summary>
    /// Tests that slicer Position is returned as a valid cell reference.
    /// This test catches bugs where Position is empty due to incorrect COM API usage.
    /// Bug context: TopLeftCell is on Slicer.Shape, not Slicer directly.
    /// </summary>
    [Fact]
    [Trait("Speed", "Medium")]
    public void CreateSlicer_ReturnsValidPosition()
    {
        // Arrange
        var batch = _fixture.BatchToken;

        // Create PivotTable
        var createResult = _pivotCommands.CreateFromRange(
            batch, _salesSheetName, "A1:D6", _salesSheetName, "F2", "PositionTestPivot");
        RequireSuccess(createResult);
        RequireSuccess(_pivotCommands.AddRowField(batch, "PositionTestPivot", "Region"));

        // Act - Create slicer at I2
        var slicerResult = _pivotCommands.CreateSlicer(
            batch, "PositionTestPivot", "Region", "PositionTestSlicer", _salesSheetName, "I2");

        // Assert - Position must be a valid cell reference, not empty
        RequireSuccess(slicerResult);
        Assert.False(string.IsNullOrEmpty(slicerResult.Position),
            "Slicer Position should not be empty - verify Shape.TopLeftCell API is used correctly");
        RequireSuccess(slicerResult);
        Assert.Equal("PivotTable", slicerResult.SourceType);
        AssertPivotSlicer("PositionTestSlicer", "PositionTestPivot", "Region", "I2", ["North", "South"]);
    }

    /// <summary>
    /// Tests that ListSlicers returns valid Position for each slicer.
    /// </summary>
    [Fact]
    [Trait("Speed", "Medium")]
    public void ListSlicers_ReturnsValidPositionForEachSlicer()
    {
        // Arrange
        var batch = _fixture.BatchToken;

        // Create PivotTable with two slicers
        var createResult = _pivotCommands.CreateFromRange(
            batch, _salesSheetName, "A1:D6", _salesSheetName, "F2", "ListPosTestPivot");
        RequireSuccess(createResult);
        RequireSuccess(_pivotCommands.AddRowField(batch, "ListPosTestPivot", "Region"));
        RequireSuccess(_pivotCommands.AddRowField(batch, "ListPosTestPivot", "Product"));

        RequireSuccess(_pivotCommands.CreateSlicer(batch, "ListPosTestPivot", "Region", "ListPosSlicer1", _salesSheetName, "I2"));
        RequireSuccess(_pivotCommands.CreateSlicer(batch, "ListPosTestPivot", "Product", "ListPosSlicer2", _salesSheetName, "K2"));

        // Act
        var listResult = _pivotCommands.ListSlicers(batch);

        // Assert - All slicers should have valid positions
        RequireSuccess(listResult);
        foreach (var slicer in listResult.Slicers)
        {
            Assert.False(string.IsNullOrEmpty(slicer.Position),
                $"Slicer '{slicer.Name}' has empty Position - verify Shape.TopLeftCell API");
        }
        Assert.Equal(2, listResult.Slicers.Count);
        AssertPivotSlicer("ListPosSlicer1", "ListPosTestPivot", "Region", "I2", ["North", "South"]);
        AssertPivotSlicer("ListPosSlicer2", "ListPosTestPivot", "Product", "K2", ["Gadget", "Widget"]);
    }

    /// <summary>
    /// Tests that FieldName is returned correctly (not "Unknown").
    /// This test catches bugs where SourceName property access fails silently.
    /// </summary>
    [Fact]
    [Trait("Speed", "Medium")]
    public void CreateSlicer_ReturnsCorrectFieldName()
    {
        // Arrange
        var batch = _fixture.BatchToken;

        // Create PivotTable
        var createResult = _pivotCommands.CreateFromRange(
            batch, _salesSheetName, "A1:D6", _salesSheetName, "F2", "FieldNameTestPivot");
        RequireSuccess(createResult);
        RequireSuccess(_pivotCommands.AddRowField(batch, "FieldNameTestPivot", "Region"));

        // Act - Create slicer for "Region" field
        var slicerResult = _pivotCommands.CreateSlicer(
            batch, "FieldNameTestPivot", "Region", "FieldNameTestSlicer", _salesSheetName, "I2");

        // Assert - FieldName must match the field name, not be "Unknown"
        RequireSuccess(slicerResult);
        Assert.NotEqual("Unknown", slicerResult.FieldName);
        Assert.Equal("Region", slicerResult.FieldName);
        AssertPivotSlicer("FieldNameTestSlicer", "FieldNameTestPivot", "Region", "I2", ["North", "South"]);
    }

    /// <summary>
    /// Tests that ConnectedPivotTables is returned correctly (not empty).
    /// This test catches bugs where PivotTables collection access fails silently.
    /// </summary>
    [Fact]
    [Trait("Speed", "Medium")]
    public void ListSlicers_ReturnsCorrectConnectedPivotTables()
    {
        // Arrange
        var batch = _fixture.BatchToken;

        // Create PivotTable
        var createResult = _pivotCommands.CreateFromRange(
            batch, _salesSheetName, "A1:D6", _salesSheetName, "F2", "ConnPivotTestPivot");
        RequireSuccess(createResult);
        RequireSuccess(_pivotCommands.AddRowField(batch, "ConnPivotTestPivot", "Region"));
        RequireSuccess(_pivotCommands.CreateSlicer(batch, "ConnPivotTestPivot", "Region", "ConnPivotTestSlicer", _salesSheetName, "I2"));

        // Act
        var listResult = _pivotCommands.ListSlicers(batch);

        // Assert - ConnectedPivotTables must contain our PivotTable
        RequireSuccess(listResult);
        var slicer = listResult.Slicers.FirstOrDefault(s => s.Name == "ConnPivotTestSlicer");
        Assert.NotNull(slicer);
        Assert.NotNull(slicer.ConnectedPivotTables);
        Assert.NotEmpty(slicer.ConnectedPivotTables);
        Assert.Contains("ConnPivotTestPivot", slicer.ConnectedPivotTables);
        AssertPivotSlicer("ConnPivotTestSlicer", "ConnPivotTestPivot", "Region", "I2", ["North", "South"]);
    }

    private void AssertFilteredPivot(string pivotName, string region, double expected)
    {
        var values = RequireSuccess(_pivotCommands.GetData(_fixture.BatchToken, pivotName)).Values;
        Assert.Equal(3, values.Count);
        Assert.All(values, row => Assert.Equal(2, row.Count));
        Assert.Equal(region, values[1][0]);
        Assert.Equal(expected, Convert.ToDouble(values[1][1], System.Globalization.CultureInfo.InvariantCulture));
        Assert.Equal(expected, Convert.ToDouble(values[2][1], System.Globalization.CultureInfo.InvariantCulture));
    }

    [Fact]
    [Trait("Speed", "Medium")]
    public void SetSlicerSelection_Additive_RetainsPreviouslySelectedItems()
    {
        var batch = _fixture.BatchToken;
        RequireSuccess(_pivotCommands.CreateFromRange(
            batch, _salesSheetName, "A1:D6", _salesSheetName, "F2", "AdditivePivot"));
        RequireSuccess(_pivotCommands.AddRowField(batch, "AdditivePivot", "Region"));
        RequireSuccess(_pivotCommands.AddValueField(batch, "AdditivePivot", "Sales"));
        RequireSuccess(_pivotCommands.CreateSlicer(
            batch, "AdditivePivot", "Region", "AdditiveSlicer", _salesSheetName, "I2"));
        RequireSuccess(_pivotCommands.SetSlicerSelection(batch, "AdditiveSlicer", ["North"]));
        AssertPivotSlicer("AdditiveSlicer", "AdditivePivot", "Region", "I2", ["North"]);
        AssertFilteredPivot("AdditivePivot", "North", 325);

        var result = RequireSuccess(_pivotCommands.SetSlicerSelection(
            batch, "AdditiveSlicer", ["South"], clearFirst: false));

        Assert.Equal(["North", "South"], result.SelectedItems.Order(StringComparer.Ordinal));
        Assert.Equal("PivotTable", result.SourceType);
        AssertPivotSlicer("AdditiveSlicer", "AdditivePivot", "Region", "I2", ["North", "South"]);
        AssertPivotSales(325, 325, "AdditivePivot");
        AssertOriginalSales();
    }

    [Fact]
    [Trait("Speed", "Medium")]
    public void ListSlicers_MixedSources_ReportsEachSourceType()
    {
        var batch = _fixture.BatchToken;
        RequireSuccess(_tableCommands.Create(
            batch, _salesSheetName, "SlicerSourceTable", "A1:D6", true, "TableStyleMedium2"));
        RequireSuccess(_tableCommands.CreateTableSlicer(
            batch, "SlicerSourceTable", "Product", "TableSourceSlicer", _salesSheetName, "M2"));
        RequireSuccess(_pivotCommands.CreateFromRange(
            batch, _salesSheetName, "A1:D6", _salesSheetName, "F2", "MixedSourcePivot"));
        RequireSuccess(_pivotCommands.AddRowField(batch, "MixedSourcePivot", "Region"));
        RequireSuccess(_pivotCommands.AddValueField(batch, "MixedSourcePivot", "Sales"));
        var created = RequireSuccess(_pivotCommands.CreateSlicer(
            batch, "MixedSourcePivot", "Region", "PivotSourceSlicer", _salesSheetName, "I2"));

        var listed = RequireSuccess(_pivotCommands.ListSlicers(batch));

        Assert.Equal("PivotTable", created.SourceType);
        Assert.Equal(2, listed.Slicers.Count);
        var table = Assert.Single(listed.Slicers, item => item.Name == "TableSourceSlicer");
        Assert.Equal("Table", table.SourceType);
        Assert.Empty(table.ConnectedPivotTables);
        Assert.Equal(["Gadget", "Widget"], table.AvailableItems.Order(StringComparer.Ordinal));
        AssertPivotSlicer("PivotSourceSlicer", "MixedSourcePivot", "Region", "I2", ["North", "South"]);
        AssertPivotSales(325, 325, "MixedSourcePivot");
        AssertOriginalSales();
    }

    private void AssertPivotSlicer(
        string name, string pivotName, string fieldName, string position, string[] selected)
    {
        var listed = RequireSuccess(_pivotCommands.ListSlicers(_fixture.BatchToken, pivotName));
        var slicer = Assert.Single(listed.Slicers, item => item.Name == name);
        Assert.Equal(_salesSheetName, slicer.SheetName);
        Assert.Equal(fieldName, slicer.FieldName);
        Assert.Equal([pivotName], slicer.ConnectedPivotTables);
        Assert.Equal(selected.Order(StringComparer.Ordinal), slicer.SelectedItems.Order(StringComparer.Ordinal));
        _fixture.ExecuteRawVerification((context, _) =>
        {
            Microsoft.Office.Interop.Excel.SlicerCaches? caches = null;
            var matches = 0;
            try
            {
                caches = context.Book.SlicerCaches;
                for (var cacheIndex = 1; cacheIndex <= caches.Count; cacheIndex++)
                {
                    Microsoft.Office.Interop.Excel.SlicerCache? cache = null;
                    Microsoft.Office.Interop.Excel.Slicers? slicers = null;
                    try
                    {
                        cache = caches.Item[cacheIndex];
                        slicers = cache.Slicers;
                        for (var slicerIndex = 1; slicerIndex <= slicers.Count; slicerIndex++)
                        {
                            Microsoft.Office.Interop.Excel.Slicer? native = null;
                            Microsoft.Office.Interop.Excel.Shape? shape = null;
                            Microsoft.Office.Interop.Excel.Range? anchor = null;
                            Microsoft.Office.Interop.Excel.Worksheet? sheet = null;
                            Microsoft.Office.Interop.Excel.Range? expectedAnchor = null;
                            Microsoft.Office.Interop.Excel.SlicerItems? items = null;
                            Microsoft.Office.Interop.Excel.SlicerPivotTables? connected = null;
                            Microsoft.Office.Interop.Excel.PivotTable? connectedPivot = null;
                            try
                            {
                                native = slicers.Item[slicerIndex];
                                if (native.Name != name)
                                {
                                    continue;
                                }
                                matches++;
                                Assert.Equal(fieldName, cache.SourceName);
                                connected = cache.PivotTables;
                                Assert.Equal(1, connected.Count);
                                connectedPivot = connected.Item[1];
                                Assert.Equal(pivotName, connectedPivot.Name);
                                shape = native.Shape;
                                anchor = shape.TopLeftCell;
                                sheet = (Microsoft.Office.Interop.Excel.Worksheet)native.Parent;
                                expectedAnchor = sheet.Range[position];
                                // Excel quantizes points; a boundary can be reported in the preceding cell.
                                Assert.Equal(Convert.ToDouble(expectedAnchor.Left), shape.Left, precision: 3);
                                Assert.Equal(Convert.ToDouble(expectedAnchor.Top), shape.Top, precision: 3);
                                Assert.Equal(slicer.Position, anchor.Address[false, false]);
                                Assert.Equal(_salesSheetName, sheet.Name);
                                items = cache.SlicerItems;
                                var selectedNative = new List<string>();
                                for (var itemIndex = 1; itemIndex <= items.Count; itemIndex++)
                                {
                                    Microsoft.Office.Interop.Excel.SlicerItem? item = null;
                                    try
                                    {
                                        item = items.Item[itemIndex];
                                        if (item.Selected)
                                        {
                                            selectedNative.Add(item.Name);
                                        }
                                    }
                                    finally
                                    {
                                        Sbroenne.ExcelMcp.ComInterop.ComUtilities.Release(ref item);
                                    }
                                }
                                Assert.Equal(selected.Order(StringComparer.Ordinal),
                                    selectedNative.Order(StringComparer.Ordinal));
                            }
                            finally
                            {
                                Sbroenne.ExcelMcp.ComInterop.ComUtilities.Release(ref connectedPivot);
                                Sbroenne.ExcelMcp.ComInterop.ComUtilities.Release(ref connected);
                                Sbroenne.ExcelMcp.ComInterop.ComUtilities.Release(ref items);
                                Sbroenne.ExcelMcp.ComInterop.ComUtilities.Release(ref expectedAnchor);
                                Sbroenne.ExcelMcp.ComInterop.ComUtilities.Release(ref sheet);
                                Sbroenne.ExcelMcp.ComInterop.ComUtilities.Release(ref anchor);
                                Sbroenne.ExcelMcp.ComInterop.ComUtilities.Release(ref shape);
                                Sbroenne.ExcelMcp.ComInterop.ComUtilities.Release(ref native);
                            }
                        }
                    }
                    finally
                    {
                        Sbroenne.ExcelMcp.ComInterop.ComUtilities.Release(ref slicers);
                        Sbroenne.ExcelMcp.ComInterop.ComUtilities.Release(ref cache);
                    }
                }
                Assert.Equal(1, matches);
            }
            finally
            {
                Sbroenne.ExcelMcp.ComInterop.ComUtilities.Release(ref caches);
            }
        });
        Assert.Equal("PivotTable", slicer.SourceType);
        AssertOriginalSales();
    }
}
