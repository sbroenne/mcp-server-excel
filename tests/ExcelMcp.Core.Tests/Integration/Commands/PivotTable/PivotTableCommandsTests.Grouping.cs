using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Models;
using Microsoft.Extensions.Logging;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Commands.PivotTable;

public partial class PivotTableCommandsTests
{
    /// <summary>
    /// Tests date grouping by Months interval creates proper monthly groups in PivotTable.
    /// </summary>
    [Fact]
    [Trait("Speed", "Medium")]
    public void GroupByDate_MonthsInterval_CreatesMonthlyGroups()
    {
        // Arrange
        var testFile = CreateTestFileWithData(nameof(GroupByDate_MonthsInterval_CreatesMonthlyGroups));

        var logger = _loggerFactory.CreateLogger<ExcelBatch>();
        using var batch = new ExcelBatch(testFile, logger);

        // Create PivotTable
        var createResult = _pivotCommands.CreateFromRange(
            batch, "SalesData", "A1:D6", "SalesData", "F2", "MonthlySales");
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
        _output.WriteLine($"Fields after grouping: {fieldNames}");

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
        var testFile = CreateTestFileWithData(nameof(GroupByDate_DaysInterval_CreatesDailyGroups));

        var logger = _loggerFactory.CreateLogger<ExcelBatch>();
        using var batch = new ExcelBatch(testFile, logger);

        // Create PivotTable
        var createResult = _pivotCommands.CreateFromRange(
            batch, "SalesData", "A1:D6", "SalesData", "F2", "DailySales");
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
        _output.WriteLine($"Fields after grouping: {fieldNames}");

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
        var testFile = CreateTestFileWithData(nameof(GroupByDate_QuartersInterval_CreatesQuarterlyGroups));

        var logger = _loggerFactory.CreateLogger<ExcelBatch>();
        using var batch = new ExcelBatch(testFile, logger);

        // Create PivotTable
        var createResult = _pivotCommands.CreateFromRange(
            batch, "SalesData", "A1:D6", "SalesData", "F2", "QuarterlySales");
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
        _output.WriteLine($"Fields after grouping: {fieldNames}");

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
        var testFile = CreateTestFileWithData(nameof(GroupByDate_YearsInterval_CreatesYearlyGroups));

        var logger = _loggerFactory.CreateLogger<ExcelBatch>();
        using var batch = new ExcelBatch(testFile, logger);

        // Create PivotTable
        var createResult = _pivotCommands.CreateFromRange(
            batch, "SalesData", "A1:D6", "SalesData", "F2", "YearlySales");
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
        _output.WriteLine($"Fields after grouping: {fieldNames}");

        // Excel creates "Years" field when grouping by years
        var hasYearsField = listResult.Fields?.Any(f => f.Name?.Contains("Year", StringComparison.OrdinalIgnoreCase) == true) == true;
        Assert.True(hasYearsField, $"Expected to find Years field after grouping. Actual fields: {fieldNames}");
    }
}
