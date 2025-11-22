using Microsoft.Extensions.Logging;
using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Models;
using Sbroenne.ExcelMcp.Core.Tests.Helpers;
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
        using var batch = new ExcelBatch(new[] { testFile }, logger);

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
        using var batch = new ExcelBatch(new[] { testFile }, logger);

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
        using var batch = new ExcelBatch(new[] { testFile }, logger);

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
        using var batch = new ExcelBatch(new[] { testFile }, logger);

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

    /// <summary>
    /// Tests numeric grouping with auto-range (uses field min/max) creates proper numeric groups.
    /// </summary>
    [Fact]
    [Trait("Speed", "Medium")]
    public void GroupByNumeric_AutoRange_CreatesNumericGroups()
    {
        // Arrange
        var testFile = CreateTestFileWithNumericData(nameof(GroupByNumeric_AutoRange_CreatesNumericGroups));

        var logger = _loggerFactory.CreateLogger<ExcelBatch>();
        using var batch = new ExcelBatch(new[] { testFile }, logger);

        // Create PivotTable
        var createResult = _pivotCommands.CreateFromRange(
            batch, "SalesData", "A1:D6", "SalesData", "F2", "SalesByRange");
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
        _output.WriteLine($"Fields after numeric grouping: {fieldNames}");

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
        var testFile = CreateTestFileWithNumericData(nameof(GroupByNumeric_CustomRange_CreatesNumericGroups));

        var logger = _loggerFactory.CreateLogger<ExcelBatch>();
        using var batch = new ExcelBatch(new[] { testFile }, logger);

        // Create PivotTable
        var createResult = _pivotCommands.CreateFromRange(
            batch, "SalesData", "A1:D6", "SalesData", "F2", "SalesByCustomRange");
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
        _output.WriteLine($"Fields after custom range grouping: {fieldNames}");

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
        var testFile = CreateTestFileWithNumericData(nameof(GroupByNumeric_SmallInterval_CreatesFineGrainedGroups));

        var logger = _loggerFactory.CreateLogger<ExcelBatch>();
        using var batch = new ExcelBatch(new[] { testFile }, logger);

        // Create PivotTable
        var createResult = _pivotCommands.CreateFromRange(
            batch, "SalesData", "A1:D6", "SalesData", "F2", "SalesBySmallRange");
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
        _output.WriteLine($"Fields after small interval grouping: {fieldNames}");

        var hasSalesField = listResult.Fields?.Any(f => f.Name == "Sales") == true;
        Assert.True(hasSalesField, $"Expected to find Sales field after grouping. Actual fields: {fieldNames}");
    }

    /// <summary>
    /// Helper method to create test file with numeric Sales data for grouping tests.
    /// </summary>
    private string CreateTestFileWithNumericData(string testName)
    {
        var testFile = CoreTestHelper.CreateUniqueTestFile(
            nameof(PivotTableCommandsTests), testName, _tempDir);

        using var batch = ExcelSession.BeginBatch(testFile);

        batch.Execute((ctx, ct) =>
        {
            dynamic sheet = ctx.Book.Worksheets.Item(1);
            sheet.Name = "SalesData";

            sheet.Range["A1"].Value2 = "Region";
            sheet.Range["B1"].Value2 = "Product";
            sheet.Range["C1"].Value2 = "Sales";
            sheet.Range["D1"].Value2 = "Date";

            sheet.Range["A2"].Value2 = "North";
            sheet.Range["B2"].Value2 = "Widget";
            sheet.Range["C2"].Value2 = 150;
            sheet.Range["D2"].Value2 = new DateTime(2025, 1, 15);

            sheet.Range["A3"].Value2 = "North";
            sheet.Range["B3"].Value2 = "Widget";
            sheet.Range["C3"].Value2 = 250;
            sheet.Range["D3"].Value2 = new DateTime(2025, 1, 20);

            sheet.Range["A4"].Value2 = "South";
            sheet.Range["B4"].Value2 = "Gadget";
            sheet.Range["C4"].Value2 = 450;
            sheet.Range["D4"].Value2 = new DateTime(2025, 2, 10);

            sheet.Range["A5"].Value2 = "North";
            sheet.Range["B5"].Value2 = "Gadget";
            sheet.Range["C5"].Value2 = 600;
            sheet.Range["D5"].Value2 = new DateTime(2025, 2, 15);

            sheet.Range["A6"].Value2 = "South";
            sheet.Range["B6"].Value2 = "Widget";
            sheet.Range["C6"].Value2 = 850;
            sheet.Range["D6"].Value2 = new DateTime(2025, 3, 5);

            // Format Sales column with numeric format (similar to date formatting requirement)
            sheet.Range["C2:C6"].NumberFormat = "0";

            // Format Date column with date format
            sheet.Range["D2:D6"].NumberFormat = "m/d/yyyy";

            return 0;
        });

        batch.Save();

        return testFile;
    }
}
