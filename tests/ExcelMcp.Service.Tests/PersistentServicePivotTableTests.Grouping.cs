using Sbroenne.ExcelMcp.Core.Commands.Range;
using Sbroenne.ExcelMcp.Core.Models;
using Xunit;

namespace Sbroenne.ExcelMcp.Service.Tests;

public sealed partial class PersistentServicePivotTableTests
{
    /// <summary>
    /// Tests date grouping by Months interval creates proper monthly groups in PivotTable.
    /// </summary>
    [Fact]
    [Trait("Speed", "Medium")]
    public void GroupByDate_MonthsInterval_CreatesMonthlyGroups()
    {
        // Arrange
        var batch = _fixture.BatchToken;

        // Create PivotTable
        var createResult = _pivotCommands.CreateFromRange(
            batch, _salesSheetName, "A1:D6", _salesSheetName, "F2", "MonthlySales");
        Assert.True(createResult.Success, $"Failed to create PivotTable: {createResult.ErrorMessage}");

        // Add Date to Row area
        var addDateResult = _pivotCommands.AddRowField(batch, "MonthlySales", "Date");
        Assert.True(addDateResult.Success, $"Failed to add Date field: {addDateResult.ErrorMessage}");

        // Add Sales to Value area
        var addValueResult = _pivotCommands.AddValueField(batch, "MonthlySales", "Sales");
        Assert.True(addValueResult.Success, $"Failed to add Sales field: {addValueResult.ErrorMessage}");

        // Act - Group Date by Months
        var groupResult = _pivotCommands.GroupByDate(batch, "MonthlySales", "Date", DateGroupingInterval.Months);

        // Assert
        Assert.True(groupResult.Success, $"GroupByDate failed: {groupResult.ErrorMessage}");
        Assert.Equal("Date", groupResult.FieldName);
        Assert.NotNull(groupResult.WorkflowHint);
        Assert.Contains("Months", groupResult.WorkflowHint);

        // Verify grouping created hierarchy by checking field list
        var listResult = _pivotCommands.ListFields(batch, "MonthlySales");
        Assert.True(listResult.Success, $"Failed to list fields: {listResult.ErrorMessage}");

        // DIAGNOSTIC: Print all field names to understand what Excel created
        var fieldNames = string.Join(", ", listResult.Fields?.Select(f => f.Name) ?? Array.Empty<string>());

        // Excel creates "Months" field when grouping by months
        var hasMonthsField = listResult.Fields?.Any(f => f.Name?.Contains("Month", StringComparison.OrdinalIgnoreCase) == true) == true;
        Assert.True(hasMonthsField, $"Expected to find Months field after grouping. Actual fields: {fieldNames}");
    }

    /// <summary>
    /// Tests date grouping by Days interval creates proper daily groups in PivotTable.
    /// </summary>
    [Fact]
    [Trait("Speed", "Medium")]
    public void GroupByDate_DaysInterval_CreatesDailyGroups()
    {
        // Arrange
        var batch = _fixture.BatchToken;

        // Create PivotTable
        var createResult = _pivotCommands.CreateFromRange(
            batch, _salesSheetName, "A1:D6", _salesSheetName, "F2", "DailySales");
        Assert.True(createResult.Success, $"Failed to create PivotTable: {createResult.ErrorMessage}");

        // Add Date to Row area
        var addDateResult = _pivotCommands.AddRowField(batch, "DailySales", "Date");
        Assert.True(addDateResult.Success, $"Failed to add Date field: {addDateResult.ErrorMessage}");

        // Add Sales to Value area
        var addValueResult = _pivotCommands.AddValueField(batch, "DailySales", "Sales");
        Assert.True(addValueResult.Success, $"Failed to add Sales field: {addValueResult.ErrorMessage}");

        // Act - Group Date by Days
        var groupResult = _pivotCommands.GroupByDate(batch, "DailySales", "Date", DateGroupingInterval.Days);

        // Assert
        Assert.True(groupResult.Success, $"GroupByDate failed: {groupResult.ErrorMessage}");
        Assert.Equal("Date", groupResult.FieldName);
        Assert.NotNull(groupResult.WorkflowHint);
        Assert.Contains("Days", groupResult.WorkflowHint);

        // Verify grouping created hierarchy by checking field list
        var listResult = _pivotCommands.ListFields(batch, "DailySales");
        Assert.True(listResult.Success, $"Failed to list fields: {listResult.ErrorMessage}");

        var fieldNames = string.Join(", ", listResult.Fields?.Select(f => f.Name) ?? Array.Empty<string>());

        // Excel creates "Days" field when grouping by days
        var hasDaysField = listResult.Fields?.Any(f => f.Name?.Contains("Day", StringComparison.OrdinalIgnoreCase) == true) == true;
        Assert.True(hasDaysField, $"Expected to find Days field after grouping. Actual fields: {fieldNames}");
    }

    /// <summary>
    /// Tests date grouping by Quarters interval creates proper quarterly groups in PivotTable.
    /// </summary>
    [Fact]
    [Trait("Speed", "Medium")]
    public void GroupByDate_QuartersInterval_CreatesQuarterlyGroups()
    {
        // Arrange
        var batch = _fixture.BatchToken;

        // Create PivotTable
        var createResult = _pivotCommands.CreateFromRange(
            batch, _salesSheetName, "A1:D6", _salesSheetName, "F2", "QuarterlySales");
        Assert.True(createResult.Success, $"Failed to create PivotTable: {createResult.ErrorMessage}");

        // Add Date to Row area
        var addDateResult = _pivotCommands.AddRowField(batch, "QuarterlySales", "Date");
        Assert.True(addDateResult.Success, $"Failed to add Date field: {addDateResult.ErrorMessage}");

        // Add Sales to Value area
        var addValueResult = _pivotCommands.AddValueField(batch, "QuarterlySales", "Sales");
        Assert.True(addValueResult.Success, $"Failed to add Sales field: {addValueResult.ErrorMessage}");

        // Act - Group Date by Quarters
        var groupResult = _pivotCommands.GroupByDate(batch, "QuarterlySales", "Date", DateGroupingInterval.Quarters);

        // Assert
        Assert.True(groupResult.Success, $"GroupByDate failed: {groupResult.ErrorMessage}");
        Assert.Equal("Date", groupResult.FieldName);
        Assert.NotNull(groupResult.WorkflowHint);
        Assert.Contains("Quarters", groupResult.WorkflowHint);

        // Verify grouping created hierarchy by checking field list
        var listResult = _pivotCommands.ListFields(batch, "QuarterlySales");
        Assert.True(listResult.Success, $"Failed to list fields: {listResult.ErrorMessage}");

        var fieldNames = string.Join(", ", listResult.Fields?.Select(f => f.Name) ?? Array.Empty<string>());

        // Excel creates "Quarters" field when grouping by quarters
        var hasQuartersField = listResult.Fields?.Any(f => f.Name?.Contains("Quarter", StringComparison.OrdinalIgnoreCase) == true) == true;
        Assert.True(hasQuartersField, $"Expected to find Quarters field after grouping. Actual fields: {fieldNames}");
    }

    /// <summary>
    /// Tests date grouping by Years interval creates proper yearly groups in PivotTable.
    /// </summary>
    [Fact]
    [Trait("Speed", "Medium")]
    public void GroupByDate_YearsInterval_CreatesYearlyGroups()
    {
        // Arrange
        var batch = _fixture.BatchToken;

        // Create PivotTable
        var createResult = _pivotCommands.CreateFromRange(
            batch, _salesSheetName, "A1:D6", _salesSheetName, "F2", "YearlySales");
        Assert.True(createResult.Success, $"Failed to create PivotTable: {createResult.ErrorMessage}");

        // Add Date to Row area
        var addDateResult = _pivotCommands.AddRowField(batch, "YearlySales", "Date");
        Assert.True(addDateResult.Success, $"Failed to add Date field: {addDateResult.ErrorMessage}");

        // Add Sales to Value area
        var addValueResult = _pivotCommands.AddValueField(batch, "YearlySales", "Sales");
        Assert.True(addValueResult.Success, $"Failed to add Sales field: {addValueResult.ErrorMessage}");

        // Act - Group Date by Years
        var groupResult = _pivotCommands.GroupByDate(batch, "YearlySales", "Date", DateGroupingInterval.Years);

        // Assert
        Assert.True(groupResult.Success, $"GroupByDate failed: {groupResult.ErrorMessage}");
        Assert.Equal("Date", groupResult.FieldName);
        Assert.NotNull(groupResult.WorkflowHint);
        Assert.Contains("Years", groupResult.WorkflowHint);

        // Verify grouping created hierarchy by checking field list
        var listResult = _pivotCommands.ListFields(batch, "YearlySales");
        Assert.True(listResult.Success, $"Failed to list fields: {listResult.ErrorMessage}");

        var fieldNames = string.Join(", ", listResult.Fields?.Select(f => f.Name) ?? Array.Empty<string>());

        // Excel creates "Years" field when grouping by years
        var hasYearsField = listResult.Fields?.Any(f => f.Name?.Contains("Year", StringComparison.OrdinalIgnoreCase) == true) == true;
        Assert.True(hasYearsField, $"Expected to find Years field after grouping. Actual fields: {fieldNames}");
    }

    /// <summary>
    /// Tests numeric grouping with auto-range (uses field min/max) creates proper numeric groups.
    /// </summary>
    [Fact]
    [Trait("Speed", "Medium")]
    public void GroupByNumeric_AutoRange_CreatesNumericGroups()
    {
        // Arrange
        var batch = _fixture.BatchToken;
        PrepareNumericSalesData(batch);

        // Create PivotTable
        var createResult = _pivotCommands.CreateFromRange(
            batch, _salesSheetName, "A1:D6", _salesSheetName, "F2", "SalesByRange");
        Assert.True(createResult.Success, $"Failed to create PivotTable: {createResult.ErrorMessage}");

        // Add Sales to Row area
        var addSalesResult = _pivotCommands.AddRowField(batch, "SalesByRange", "Sales");
        Assert.True(addSalesResult.Success, $"Failed to add Sales field: {addSalesResult.ErrorMessage}");

        // Add Region to Value area (Count)
        var addValueResult = _pivotCommands.AddValueField(batch, "SalesByRange", "Region", AggregationFunction.Count);
        Assert.True(addValueResult.Success, $"Failed to add Region field: {addValueResult.ErrorMessage}");

        // Act - Group Sales by 100 with auto-range
        var groupResult = _pivotCommands.GroupByNumeric(batch, "SalesByRange", "Sales", start: null, endValue: null, intervalSize: 100);

        // Assert
        Assert.True(groupResult.Success, $"GroupByNumeric failed: {groupResult.ErrorMessage}");
        Assert.Equal("Sales", groupResult.FieldName);
        Assert.NotNull(groupResult.WorkflowHint);
        Assert.Contains("100", groupResult.WorkflowHint);

        // Verify grouping created groups by checking field list
        var listResult = _pivotCommands.ListFields(batch, "SalesByRange");
        Assert.True(listResult.Success, $"Failed to list fields: {listResult.ErrorMessage}");

        var fieldNames = string.Join(", ", listResult.Fields?.Select(f => f.Name) ?? Array.Empty<string>());

        // After grouping, field should still be named "Sales" but contain grouped values
        var hasSalesField = listResult.Fields?.Any(f => f.Name == "Sales") == true;
        Assert.True(hasSalesField, $"Expected to find Sales field after grouping. Actual fields: {fieldNames}");
    }

    /// <summary>
    /// Tests numeric grouping with custom range creates proper numeric groups.
    /// </summary>
    [Fact]
    [Trait("Speed", "Medium")]
    public void GroupByNumeric_CustomRange_CreatesNumericGroups()
    {
        // Arrange
        var batch = _fixture.BatchToken;
        PrepareNumericSalesData(batch);

        // Create PivotTable
        var createResult = _pivotCommands.CreateFromRange(
            batch, _salesSheetName, "A1:D6", _salesSheetName, "F2", "SalesByCustomRange");
        Assert.True(createResult.Success, $"Failed to create PivotTable: {createResult.ErrorMessage}");

        // Add Sales to Row area
        var addSalesResult = _pivotCommands.AddRowField(batch, "SalesByCustomRange", "Sales");
        Assert.True(addSalesResult.Success, $"Failed to add Sales field: {addSalesResult.ErrorMessage}");

        // Add Region to Value area (Count)
        var addValueResult = _pivotCommands.AddValueField(batch, "SalesByCustomRange", "Region", AggregationFunction.Count);
        Assert.True(addValueResult.Success, $"Failed to add Region field: {addValueResult.ErrorMessage}");

        // Act - Group Sales 0-1000 by 200
        var groupResult = _pivotCommands.GroupByNumeric(batch, "SalesByCustomRange", "Sales", start: 0, endValue: 1000, intervalSize: 200);

        // Assert
        Assert.True(groupResult.Success, $"GroupByNumeric failed: {groupResult.ErrorMessage}");
        Assert.Equal("Sales", groupResult.FieldName);
        Assert.NotNull(groupResult.WorkflowHint);
        Assert.Contains("200", groupResult.WorkflowHint);

        // Verify grouping created groups
        var listResult = _pivotCommands.ListFields(batch, "SalesByCustomRange");
        Assert.True(listResult.Success, $"Failed to list fields: {listResult.ErrorMessage}");

        var fieldNames = string.Join(", ", listResult.Fields?.Select(f => f.Name) ?? Array.Empty<string>());

        var hasSalesField = listResult.Fields?.Any(f => f.Name == "Sales") == true;
        Assert.True(hasSalesField, $"Expected to find Sales field after grouping. Actual fields: {fieldNames}");
    }

    /// <summary>
    /// Tests numeric grouping with small interval creates fine-grained groups.
    /// </summary>
    [Fact]
    [Trait("Speed", "Medium")]
    public void GroupByNumeric_SmallInterval_CreatesFineGrainedGroups()
    {
        // Arrange
        var batch = _fixture.BatchToken;
        PrepareNumericSalesData(batch);

        // Create PivotTable
        var createResult = _pivotCommands.CreateFromRange(
            batch, _salesSheetName, "A1:D6", _salesSheetName, "F2", "SalesBySmallRange");
        Assert.True(createResult.Success, $"Failed to create PivotTable: {createResult.ErrorMessage}");

        // Add Sales to Row area
        var addSalesResult = _pivotCommands.AddRowField(batch, "SalesBySmallRange", "Sales");
        Assert.True(addSalesResult.Success, $"Failed to add Sales field: {addSalesResult.ErrorMessage}");

        // Add Region to Value area (Count)
        var addValueResult = _pivotCommands.AddValueField(batch, "SalesBySmallRange", "Region", AggregationFunction.Count);
        Assert.True(addValueResult.Success, $"Failed to add Region field: {addValueResult.ErrorMessage}");

        // Act - Group Sales by 50 for fine-grained analysis
        var groupResult = _pivotCommands.GroupByNumeric(batch, "SalesBySmallRange", "Sales", start: null, endValue: null, intervalSize: 50);

        // Assert
        Assert.True(groupResult.Success, $"GroupByNumeric failed: {groupResult.ErrorMessage}");
        Assert.Equal("Sales", groupResult.FieldName);
        Assert.NotNull(groupResult.WorkflowHint);
        Assert.Contains("50", groupResult.WorkflowHint);

        // Verify grouping created groups
        var listResult = _pivotCommands.ListFields(batch, "SalesBySmallRange");
        Assert.True(listResult.Success, $"Failed to list fields: {listResult.ErrorMessage}");

        var fieldNames = string.Join(", ", listResult.Fields?.Select(f => f.Name) ?? Array.Empty<string>());

        var hasSalesField = listResult.Fields?.Any(f => f.Name == "Sales") == true;
        Assert.True(hasSalesField, $"Expected to find Sales field after grouping. Actual fields: {fieldNames}");
    }
    private void PrepareNumericSalesData(
        Sbroenne.ExcelMcp.ComInterop.Session.IExcelBatch batch)
    {
        _commands.SetValues(
            batch,
            _salesSheetName,
            "A1:D6",
            [
                ["Region", "Product", "Sales", "Date"],
                ["North", "Widget", 150, "2025-01-15"],
                ["North", "Widget", 250, "2025-01-20"],
                ["South", "Gadget", 450, "2025-02-10"],
                ["North", "Gadget", 600, "2025-02-15"],
                ["South", "Widget", 850, "2025-03-05"],
            ],
            overwritePolicy: OverwritePolicy.Allow);
        _commands.SetNumberFormat(batch, _salesSheetName, "C2:C6", "0");
        _commands.SetNumberFormat(batch, _salesSheetName, "D2:D6", "m/d/yyyy");
    }
}
