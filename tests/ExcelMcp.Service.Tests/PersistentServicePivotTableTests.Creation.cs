using Sbroenne.ExcelMcp.Core.Models;
using Xunit;

namespace Sbroenne.ExcelMcp.Service.Tests;

/// <summary>
/// Tests for PivotTable creation operations
/// </summary>
public sealed partial class PersistentServicePivotTableTests
{
    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void CreateFromRange_MissingSheet_PreservesExistingPivot(bool missingSource)
    {
        RequireSuccess(_pivotCommands.CreateFromRange(
            _fixture.BatchToken, _salesSheetName, "A1:D6", _salesSheetName, "F1", "TestPivot"));
        RequireSuccess(_pivotCommands.AddRowField(_fixture.BatchToken, "TestPivot", "Region"));
        RequireSuccess(_pivotCommands.AddValueField(_fixture.BatchToken, "TestPivot", "Sales"));
        AssertPivotSales(325, 325);
        var before = SnapshotPivot();

        var error = Assert.Throws<InvalidOperationException>(() =>
            _pivotCommands.CreateFromRange(
                _fixture.BatchToken, missingSource ? "MissingSheet" : _salesSheetName, "A1:D6",
                missingSource ? _salesSheetName : "MissingSheet", "L1", "RejectedPivot"));

        Assert.Contains("pivottable.create-from-range failed [ComInterop/COMException]",
            error.Message, StringComparison.Ordinal);
        Assert.Equal(before, SnapshotPivot());
        Assert.Equal("TestPivot", Assert.Single(
            RequireSuccess(_pivotCommands.List(_fixture.BatchToken)).PivotTables).Name);
        AssertPivotSales(325, 325);
        AssertOriginalSales();
    }

    /// <inheritdoc/>
    [Fact]
    public void CreateFromRange_PopulatedRangeWithHeaders_CreatesCorrectPivotStructure()
    {
        // Arrange

        // Act
        var batch = _fixture.BatchToken;
        var result = _pivotCommands.CreateFromRange(
            batch,
            _salesSheetName, "A1:D6",
            _salesSheetName, "F1",
            "TestPivot");

        // Assert
        RequireSuccess(result);
        Assert.Equal("TestPivot", result.PivotTableName);
        Assert.Equal(_salesSheetName, result.SheetName);
        Assert.Equal(4, result.AvailableFields.Count);
        Assert.Equal($"'{_salesSheetName}'!A1:D6", result.SourceData);
        AssertCreatedPivot(result, _salesSheetName, "F1");
    }
    /// <inheritdoc/>

    [Fact]
    public void CreateFromTable_WithValidTable_CreatesCorrectPivotStructure()
    {
        // Arrange

        // Act - Use single batch for table creation and pivot creation
        var batch = _fixture.BatchToken;

        // Create table first
        RequireSuccess(_tableCommands.Create(batch, _salesSheetName, "SalesTable", "A1:D6", true, TableStylePresets.Medium2));

        // Create pivot from table
        var result = _pivotCommands.CreateFromTable(
            batch,
            "SalesTable",
            _salesSheetName, "F1",
            "TablePivot");

        // Assert
        RequireSuccess(result);
        Assert.Equal("TablePivot", result.PivotTableName);
        Assert.Equal(_salesSheetName, result.SheetName);
        Assert.Equal(4, result.AvailableFields.Count);
        Assert.Equal($"{_salesSheetName}!SalesTable", result.SourceData);
        AssertCreatedPivot(result, _salesSheetName, "F1");
    }
    /// <inheritdoc/>

    [Fact]
    public void CreateFromDataModel_NoDataModel_ReturnsError()
    {
        // Arrange - Use regular file without Data Model

        // Act & Assert - expects exception when Data Model is empty
        var batch = _fixture.BatchToken;
        RequireSuccess(_pivotCommands.CreateFromRange(
            batch, _salesSheetName, "A1:D6", _salesSheetName, "F1", "TestPivot"));
        var before = SnapshotPivot();
        var ex = Assert.Throws<InvalidOperationException>(() => _pivotCommands.CreateFromDataModel(
            batch,
            "AnyTable",
            _salesSheetName,
            "F1",
            "FailedPivot"));
        Assert.Contains("Data Model does not contain any tables", ex.Message);
        Assert.Equal(before, SnapshotPivot());
        Assert.Equal("TestPivot", Assert.Single(RequireSuccess(_pivotCommands.List(batch)).PivotTables).Name);
        AssertOriginalSales();
    }
    /// <inheritdoc/>

    [Fact]
    public void AddRowField_WithValidField_AddsFieldToRows()
    {
        // Arrange

        // Act - Use single batch for create and add field
        var batch = _fixture.BatchToken;

        // Create pivot
        var createResult = _pivotCommands.CreateFromRange(
            batch, _salesSheetName, "A1:D6", _salesSheetName, "F1", "TestPivot");
        RequireSuccess(createResult);

        // Add row field
        var result = _pivotCommands.AddRowField(batch, "TestPivot", "Region");

        // Assert
        RequireSuccess(result);
        Assert.Equal("Region", result.FieldName);
        RequireSuccess(result);
        AssertNativeField("TestPivot", "Region", PivotFieldArea.Row);
        AssertOriginalSales();
    }
    /// <inheritdoc/>

    [Fact]
    public void ListFields_AfterCreate_ReturnsAvailableFields()
    {
        // Arrange

        // Act - Use single batch for create and list fields
        var batch = _fixture.BatchToken;

        // Create pivot
        var createResult = _pivotCommands.CreateFromRange(
            batch, _salesSheetName, "A1:D6", _salesSheetName, "F1", "TestPivot");
        RequireSuccess(createResult);

        // List fields
        var result = _pivotCommands.ListFields(batch, "TestPivot");

        // Assert
        RequireSuccess(result);
        Assert.NotNull(result.Fields);
        RequireSuccess(result);
        Assert.Equal(["Date", "Product", "Region", "Sales"],
            result.Fields.Select(field => field.Name).Order(StringComparer.Ordinal));
        Assert.All(result.Fields, field => Assert.Equal(PivotFieldArea.Hidden, field.Area));
        AssertOriginalSales();
    }
}
