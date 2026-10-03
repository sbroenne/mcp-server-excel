using Sbroenne.ExcelMcp.Core.Commands.Range;
using Sbroenne.ExcelMcp.Core.Models;
using Xunit;

namespace Sbroenne.ExcelMcp.Service.Tests;

/// <summary>
/// Tests for PivotTable operations (List, GetInfo, Delete, Refresh, GetData)
/// Optimized: Single batch per test, no SaveAsync() unless testing persistence
/// </summary>
public sealed partial class PersistentServicePivotTableTests
{
    /// <inheritdoc/>
    [Fact]
    [Trait("Speed", "Medium")]
    public void List_EmptyWorkbook_ReturnsEmptyList()
    {
        // Arrange

        // Act
        var batch = _fixture.BatchToken;
        var result = _pivotCommands.List(batch);

        // Assert
        RequireSuccess(result);
        Assert.NotNull(result.PivotTables);
        Assert.Empty(result.PivotTables);
        AssertOriginalSales();
    }
    /// <inheritdoc/>

    [Fact]
    [Trait("Speed", "Medium")]
    public void List_WithPivotTable_ReturnsPivotTableInfo()
    {
        // Arrange

        var batch = _fixture.BatchToken;
        var createResult = _pivotCommands.CreateFromRange(
            batch, _salesSheetName, "A1:D6", _salesSheetName, "F1", "TestPivot");
        RequireSuccess(createResult);

        // Act - No save needed, same batch
        var result = _pivotCommands.List(batch);

        // Assert
        RequireSuccess(result);
        Assert.NotEmpty(result.PivotTables);
        var pivot = Assert.Single(result.PivotTables);
        Assert.Equal("TestPivot", pivot.Name);
        Assert.Equal(_salesSheetName, pivot.SheetName);
        RequireSuccess(result);
        Assert.Equal((0, 0, 0, 0), (pivot.RowFieldCount, pivot.ColumnFieldCount,
            pivot.ValueFieldCount, pivot.FilterFieldCount));
        var read = RequireSuccess(_pivotCommands.Read(batch, "TestPivot"));
        Assert.Equal(read.PivotTable.Range, pivot.Range);
        Assert.Equal(read.PivotTable.SourceData, pivot.SourceData);
        AssertOriginalSales();
    }
    /// <inheritdoc/>

    [Fact]
    [Trait("Speed", "Medium")]
    public void GetInfo_ExistingPivotTable_ReturnsCompleteInfo()
    {
        // Arrange

        var batch = _fixture.BatchToken;
        var createResult = _pivotCommands.CreateFromRange(
            batch, _salesSheetName, "A1:D6", _salesSheetName, "F1", "TestPivot");
        RequireSuccess(createResult);

        // Act - No save needed
        var result = _pivotCommands.Read(batch, "TestPivot");

        // Assert
        RequireSuccess(result);
        Assert.Equal("TestPivot", result.PivotTable.Name);
        Assert.NotEmpty(result.Fields);
        Assert.Equal(4, result.Fields.Count); // Region, Product, Sales, Date
        RequireSuccess(result);
        Assert.Equal(_salesSheetName, result.PivotTable.SheetName);
        Assert.Equal(["Date", "Product", "Region", "Sales"],
            result.Fields.Select(field => field.Name).Order(StringComparer.Ordinal));
        Assert.All(result.Fields, field => Assert.Equal(PivotFieldArea.Hidden, field.Area));
        AssertOriginalSales();
    }
    /// <inheritdoc/>

    [Fact]
    [Trait("Speed", "Medium")]
    public void GetInfo_NonExistentPivotTable_ReturnsError()
    {
        // Arrange

        // Act & Assert - expects exception when pivot table not found
        var batch = _fixture.BatchToken;
        RequireSuccess(_pivotCommands.CreateFromRange(
            batch, _salesSheetName, "A1:D6", _salesSheetName, "F1", "TestPivot"));
        var before = SnapshotPivot();
        var ex = Assert.Throws<InvalidOperationException>(() => _pivotCommands.Read(batch, "NonExistent"));
        Assert.Contains("not found", ex.Message, StringComparison.OrdinalIgnoreCase);
        Assert.Equal(before, SnapshotPivot());
        AssertOriginalSales();
    }
    /// <inheritdoc/>

    [Fact]
    [Trait("Speed", "Medium")]
    public void Delete_ExistingPivotTable_RemovesPivotTable()
    {
        // Arrange

        var batch = _fixture.BatchToken;
        var createResult = _pivotCommands.CreateFromRange(
            batch, _salesSheetName, "A1:D6", _salesSheetName, "F1", "TestPivot");
        RequireSuccess(createResult);
        RequireSuccess(_pivotCommands.CreateFromRange(
            batch, _salesSheetName, "A1:D6", _salesSheetName, "L1", "NeighborPivot"));
        var neighborBefore = SnapshotPivot("NeighborPivot");
        RequireSuccess(_commands.SetValues(batch, _salesSheetName, "K10", [["Unchanged"]]));

        // Act - Delete in same batch
        var deleteResult = _pivotCommands.Delete(batch, "TestPivot");

        // Assert
        RequireSuccess(deleteResult);

        // Verify pivot no longer exists
        var listResult = _pivotCommands.List(batch);
        RequireSuccess(listResult);
        RequireSuccess(deleteResult);
        RequireSuccess(listResult);
        Assert.Equal("NeighborPivot", Assert.Single(listResult.PivotTables).Name);
        Assert.Equal(neighborBefore, SnapshotPivot("NeighborPivot"));
        AssertOriginalSales();
        Assert.Equal("Unchanged", Assert.Single(Assert.Single(
            RequireSuccess(_commands.GetValues(batch, _salesSheetName, "K10")).Values)));
    }
    /// <inheritdoc/>

    [Fact]
    [Trait("Speed", "Medium")]
    public void Delete_NonExistentPivotTable_ReturnsError()
    {
        // Arrange

        // Act & Assert - expects exception when pivot table not found
        var batch = _fixture.BatchToken;
        RequireSuccess(_pivotCommands.CreateFromRange(
            batch, _salesSheetName, "A1:D6", _salesSheetName, "F1", "TestPivot"));
        var before = SnapshotPivot();
        var ex = Assert.Throws<InvalidOperationException>(() => _pivotCommands.Delete(batch, "NonExistent"));
        Assert.Contains("not found", ex.Message, StringComparison.OrdinalIgnoreCase);
        Assert.Equal(before, SnapshotPivot());
        Assert.Equal("TestPivot", Assert.Single(RequireSuccess(_pivotCommands.List(batch)).PivotTables).Name);
        AssertOriginalSales();
    }
    /// <inheritdoc/>

    [Fact]
    [Trait("Speed", "Medium")]
    public void Refresh_ExistingPivotTable_UpdatesData()
    {
        // Arrange

        var batch = _fixture.BatchToken;
        var createResult = _pivotCommands.CreateFromRange(
            batch, _salesSheetName, "A1:D6", _salesSheetName, "F1", "TestPivot");
        RequireSuccess(createResult);

        RequireSuccess(_pivotCommands.AddRowField(batch, "TestPivot", "Region"));
        RequireSuccess(_pivotCommands.AddValueField(batch, "TestPivot", "Sales", AggregationFunction.Sum, "Total Sales"));
        RequireSuccess(_pivotCommands.Refresh(batch, "TestPivot"));
        AssertPivotSales(325, 325);
        var changed = _commands.SetValues(batch, _salesSheetName, "C2", [[400]], overwritePolicy: OverwritePolicy.Allow);
        RequireSuccess(changed);
        AssertPivotSales(325, 325);

        var result = _pivotCommands.Refresh(batch, "TestPivot");

        // Assert
        RequireSuccess(result);
        Assert.Equal("TestPivot", result.PivotTableName);
        Assert.True(result.RefreshTime <= DateTime.Now);
        Assert.Equal(5, result.SourceRecordCount);
        AssertPivotSales(625, 325);
    }
    /// <inheritdoc/>

    [Fact]
    [Trait("Speed", "Medium")]
    public void GetData_ExistingPivotTable_ReturnsData()
    {
        // Arrange

        var batch = _fixture.BatchToken;

        // Create pivot with row field to generate data
        var createResult = _pivotCommands.CreateFromRange(
            batch, _salesSheetName, "A1:D6", _salesSheetName, "F1", "TestPivot");
        RequireSuccess(createResult);

        // Add Region to row area
        var addRowResult = _pivotCommands.AddRowField(batch, "TestPivot", "Region");
        RequireSuccess(addRowResult);

        // Act - GetData in same batch
        var result = _pivotCommands.GetData(batch, "TestPivot");

        // Assert
        RequireSuccess(result);
        Assert.Equal("TestPivot", result.PivotTableName);
        Assert.NotNull(result.Values);
        Assert.NotEmpty(result.Values);
        Assert.Equal(["North", "South"], result.Values
            .Select(row => row[0]?.ToString()).Where(value => value is "North" or "South"));
    }

    private void AssertPivotSales(int north, int south, string pivotName = "TestPivot")
    {
        var data = RequireSuccess(_pivotCommands.GetData(_fixture.BatchToken, pivotName));
        RequireSuccess(data);
        var expected = new Dictionary<string, int>
        {
            ["North"] = north,
            ["South"] = south
        };
        foreach (var pair in expected)
        {
            var row = Assert.Single(data.Values, values =>
                string.Equals(values[0]?.ToString(), pair.Key, StringComparison.Ordinal));
            Assert.Equal(pair.Value, Convert.ToInt32(row[1], System.Globalization.CultureInfo.InvariantCulture));
        }
        Assert.Equal(4, data.Values.Count);
        Assert.Equal(north + south, Convert.ToInt32(data.Values[^1][1], System.Globalization.CultureInfo.InvariantCulture));
    }
}
