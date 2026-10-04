using Sbroenne.ExcelMcp.Core.Commands.Range;
using Sbroenne.ExcelMcp.Core.Models;
using Xunit;

namespace Sbroenne.ExcelMcp.Service.Tests;

public sealed partial class PersistentServicePivotTableTests
{
    [Fact]
    public void GroupByDate_MissingField_PreservesConfiguredPivot()
    {
        RequireSuccess(_pivotCommands.CreateFromRange(
            _fixture.BatchToken, _salesSheetName, "A1:D6", _salesSheetName, "F1", "TestPivot"));
        RequireSuccess(_pivotCommands.AddRowField(_fixture.BatchToken, "TestPivot", "Region"));
        RequireSuccess(_pivotCommands.AddValueField(_fixture.BatchToken, "TestPivot", "Sales"));
        AssertPivotSales(325, 325);
        var before = RequireSuccess(_pivotCommands.Read(_fixture.BatchToken, "TestPivot"));
        var data = System.Text.Json.JsonSerializer.Serialize(
            RequireSuccess(_pivotCommands.GetData(_fixture.BatchToken, "TestPivot")));

        var result = _pivotCommands.GroupByDate(
            _fixture.BatchToken, "TestPivot", "MissingField", DateGroupingInterval.Months);

        Assert.False(result.Success);
        Assert.Equal("Failed to group field by date: Field 'MissingField' not found in PivotTable",
            result.ErrorMessage);
        var after = RequireSuccess(_pivotCommands.Read(_fixture.BatchToken, "TestPivot"));
        Assert.NotNull(before.PivotTable.LastRefresh);
        Assert.NotNull(after.PivotTable.LastRefresh);
        Assert.True(after.PivotTable.LastRefresh >= before.PivotTable.LastRefresh);
        // Grouping refreshes before looking up the field, even when that lookup fails.
        after.PivotTable.LastRefresh = before.PivotTable.LastRefresh;
        Assert.Equal(System.Text.Json.JsonSerializer.Serialize(before),
            System.Text.Json.JsonSerializer.Serialize(after));
        Assert.Equal(data, System.Text.Json.JsonSerializer.Serialize(
            RequireSuccess(_pivotCommands.GetData(_fixture.BatchToken, "TestPivot"))));
        AssertOriginalSales();
    }

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
        RequireSuccess(createResult);

        // Add Date to Row area
        var addDateResult = _pivotCommands.AddRowField(batch, "MonthlySales", "Date");
        RequireSuccess(addDateResult);

        // Add Sales to Value area
        var addValueResult = _pivotCommands.AddValueField(batch, "MonthlySales", "Sales");
        RequireSuccess(addValueResult);

        // Act - Group Date by Months
        var groupResult = _pivotCommands.GroupByDate(batch, "MonthlySales", "Date", DateGroupingInterval.Months);

        // Assert
        RequireSuccess(groupResult);
        Assert.Equal("Date", groupResult.FieldName);
        Assert.NotNull(groupResult.WorkflowHint);
        Assert.Contains("Months", groupResult.WorkflowHint);

        // Verify grouping created hierarchy by checking field list
        var listResult = _pivotCommands.ListFields(batch, "MonthlySales");
        RequireSuccess(listResult);

        // DIAGNOSTIC: Print all field names to understand what Excel created
        var fieldNames = string.Join(", ", listResult.Fields?.Select(f => f.Name) ?? Array.Empty<string>());

        // Excel creates "Months" field when grouping by months
        var hasMonthsField = listResult.Fields?.Any(f => f.Name?.Contains("Month", StringComparison.OrdinalIgnoreCase) == true) == true;
        Assert.True(hasMonthsField, $"Expected to find Months field after grouping. Actual fields: {fieldNames}");
        RequireSuccess(groupResult);
        AssertGroupedValues("MonthlySales", [650d, 250d, 275d, 125d], 650);
        Assert.Equal("2025", RequireSuccess(_pivotCommands.GetData(batch, "MonthlySales")).Values[1][0]);
        AssertOriginalSales();
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
        RequireSuccess(createResult);

        // Add Date to Row area
        var addDateResult = _pivotCommands.AddRowField(batch, "DailySales", "Date");
        RequireSuccess(addDateResult);

        // Add Sales to Value area
        var addValueResult = _pivotCommands.AddValueField(batch, "DailySales", "Sales");
        RequireSuccess(addValueResult);

        // Act - Group Date by Days
        var groupResult = _pivotCommands.GroupByDate(batch, "DailySales", "Date", DateGroupingInterval.Days);

        // Assert
        RequireSuccess(groupResult);
        Assert.Equal("Date", groupResult.FieldName);
        Assert.NotNull(groupResult.WorkflowHint);
        Assert.Contains("Days", groupResult.WorkflowHint);

        // Verify grouping created hierarchy by checking field list
        var listResult = _pivotCommands.ListFields(batch, "DailySales");
        RequireSuccess(listResult);

        var fieldNames = string.Join(", ", listResult.Fields?.Select(f => f.Name) ?? Array.Empty<string>());

        // Excel creates "Days" field when grouping by days
        var hasDaysField = listResult.Fields?.Any(f => f.Name?.Contains("Day", StringComparison.OrdinalIgnoreCase) == true) == true;
        Assert.True(hasDaysField, $"Expected to find Days field after grouping. Actual fields: {fieldNames}");
        RequireSuccess(groupResult);
        AssertGroupedValues("DailySales", [100d, 150d, 200d, 75d, 125d], 650);
        AssertOriginalSales();
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
        RequireSuccess(createResult);

        // Add Date to Row area
        var addDateResult = _pivotCommands.AddRowField(batch, "QuarterlySales", "Date");
        RequireSuccess(addDateResult);

        // Add Sales to Value area
        var addValueResult = _pivotCommands.AddValueField(batch, "QuarterlySales", "Sales");
        RequireSuccess(addValueResult);

        // Act - Group Date by Quarters
        var groupResult = _pivotCommands.GroupByDate(batch, "QuarterlySales", "Date", DateGroupingInterval.Quarters);

        // Assert
        RequireSuccess(groupResult);
        Assert.Equal("Date", groupResult.FieldName);
        Assert.NotNull(groupResult.WorkflowHint);
        Assert.Contains("Quarters", groupResult.WorkflowHint);

        // Verify grouping created hierarchy by checking field list
        var listResult = _pivotCommands.ListFields(batch, "QuarterlySales");
        RequireSuccess(listResult);

        var fieldNames = string.Join(", ", listResult.Fields?.Select(f => f.Name) ?? Array.Empty<string>());

        // Excel creates "Quarters" field when grouping by quarters
        var hasQuartersField = listResult.Fields?.Any(f => f.Name?.Contains("Quarter", StringComparison.OrdinalIgnoreCase) == true) == true;
        Assert.True(hasQuartersField, $"Expected to find Quarters field after grouping. Actual fields: {fieldNames}");
        RequireSuccess(groupResult);
        AssertGroupedValues("QuarterlySales", [650d, 650d], 650);
        var quarterly = RequireSuccess(_pivotCommands.GetData(batch, "QuarterlySales"));
        Assert.Equal("2025", quarterly.Values[1][0]);
        Assert.Equal("Qtr1", quarterly.Values[2][0]);
        AssertOriginalSales();
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
        RequireSuccess(createResult);

        // Add Date to Row area
        var addDateResult = _pivotCommands.AddRowField(batch, "YearlySales", "Date");
        RequireSuccess(addDateResult);

        // Add Sales to Value area
        var addValueResult = _pivotCommands.AddValueField(batch, "YearlySales", "Sales");
        RequireSuccess(addValueResult);

        // Act - Group Date by Years
        var groupResult = _pivotCommands.GroupByDate(batch, "YearlySales", "Date", DateGroupingInterval.Years);

        // Assert
        RequireSuccess(groupResult);
        Assert.Equal("Date", groupResult.FieldName);
        Assert.NotNull(groupResult.WorkflowHint);
        Assert.Contains("Years", groupResult.WorkflowHint);

        // Verify grouping created hierarchy by checking field list
        var listResult = _pivotCommands.ListFields(batch, "YearlySales");
        RequireSuccess(listResult);

        var fieldNames = string.Join(", ", listResult.Fields?.Select(f => f.Name) ?? Array.Empty<string>());

        // Excel creates "Years" field when grouping by years
        var hasYearsField = listResult.Fields?.Any(f => f.Name?.Contains("Year", StringComparison.OrdinalIgnoreCase) == true) == true;
        Assert.True(hasYearsField, $"Expected to find Years field after grouping. Actual fields: {fieldNames}");
        RequireSuccess(groupResult);
        AssertGroupedValues("YearlySales", [650d], 650);
        var yearly = RequireSuccess(_pivotCommands.GetData(batch, "YearlySales"));
        Assert.Equal("2025", Convert.ToString(yearly.Values[1][0],
            System.Globalization.CultureInfo.InvariantCulture));
        AssertOriginalSales();
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
        RequireSuccess(createResult);

        // Add Sales to Row area
        var addSalesResult = _pivotCommands.AddRowField(batch, "SalesByRange", "Sales");
        RequireSuccess(addSalesResult);

        // Add Region to Value area (Count)
        var addValueResult = _pivotCommands.AddValueField(batch, "SalesByRange", "Region", AggregationFunction.Count);
        RequireSuccess(addValueResult);

        // Act - Group Sales by 100 with auto-range
        var groupResult = _pivotCommands.GroupByNumeric(batch, "SalesByRange", "Sales", start: null, endValue: null, intervalSize: 100);

        // Assert
        RequireSuccess(groupResult);
        Assert.Equal("Sales", groupResult.FieldName);
        Assert.NotNull(groupResult.WorkflowHint);
        Assert.Contains("100", groupResult.WorkflowHint);

        // Verify grouping created groups by checking field list
        var listResult = _pivotCommands.ListFields(batch, "SalesByRange");
        RequireSuccess(listResult);

        var fieldNames = string.Join(", ", listResult.Fields?.Select(f => f.Name) ?? Array.Empty<string>());

        // After grouping, field should still be named "Sales" but contain grouped values
        var hasSalesField = listResult.Fields?.Any(f => f.Name == "Sales") == true;
        Assert.True(hasSalesField, $"Expected to find Sales field after grouping. Actual fields: {fieldNames}");
        RequireSuccess(groupResult);
        AssertNumericGroups("SalesByRange", ["150-249", "250-349", "450-549", "550-649", "750-850"]);
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
        RequireSuccess(createResult);

        // Add Sales to Row area
        var addSalesResult = _pivotCommands.AddRowField(batch, "SalesByCustomRange", "Sales");
        RequireSuccess(addSalesResult);

        // Add Region to Value area (Count)
        var addValueResult = _pivotCommands.AddValueField(batch, "SalesByCustomRange", "Region", AggregationFunction.Count);
        RequireSuccess(addValueResult);

        // Act - Group Sales 0-1000 by 200
        var groupResult = _pivotCommands.GroupByNumeric(batch, "SalesByCustomRange", "Sales", start: 0, endValue: 1000, intervalSize: 200);

        // Assert
        RequireSuccess(groupResult);
        Assert.Equal("Sales", groupResult.FieldName);
        Assert.NotNull(groupResult.WorkflowHint);
        Assert.Contains("200", groupResult.WorkflowHint);

        // Verify grouping created groups
        var listResult = _pivotCommands.ListFields(batch, "SalesByCustomRange");
        RequireSuccess(listResult);

        var fieldNames = string.Join(", ", listResult.Fields?.Select(f => f.Name) ?? Array.Empty<string>());

        var hasSalesField = listResult.Fields?.Any(f => f.Name == "Sales") == true;
        Assert.True(hasSalesField, $"Expected to find Sales field after grouping. Actual fields: {fieldNames}");
        RequireSuccess(groupResult);
        AssertNumericGroups("SalesByCustomRange", ["0-199", "200-399", "400-599", "600-799", "800-1000"]);
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
        RequireSuccess(createResult);

        // Add Sales to Row area
        var addSalesResult = _pivotCommands.AddRowField(batch, "SalesBySmallRange", "Sales");
        RequireSuccess(addSalesResult);

        // Add Region to Value area (Count)
        var addValueResult = _pivotCommands.AddValueField(batch, "SalesBySmallRange", "Region", AggregationFunction.Count);
        RequireSuccess(addValueResult);

        // Act - Group Sales by 50 for fine-grained analysis
        var groupResult = _pivotCommands.GroupByNumeric(batch, "SalesBySmallRange", "Sales", start: null, endValue: null, intervalSize: 50);

        // Assert
        RequireSuccess(groupResult);
        Assert.Equal("Sales", groupResult.FieldName);
        Assert.NotNull(groupResult.WorkflowHint);
        Assert.Contains("50", groupResult.WorkflowHint);

        // Verify grouping created groups
        var listResult = _pivotCommands.ListFields(batch, "SalesBySmallRange");
        RequireSuccess(listResult);

        var fieldNames = string.Join(", ", listResult.Fields?.Select(f => f.Name) ?? Array.Empty<string>());

        var hasSalesField = listResult.Fields?.Any(f => f.Name == "Sales") == true;
        Assert.True(hasSalesField, $"Expected to find Sales field after grouping. Actual fields: {fieldNames}");
        RequireSuccess(groupResult);
        AssertNumericGroups("SalesBySmallRange", ["150-199", "250-299", "450-499", "600-649", "800-850"]);
    }
    private void PrepareNumericSalesData(
        Sbroenne.ExcelMcp.ComInterop.Session.IExcelBatch batch)
    {
        RequireSuccess(_commands.SetValues(
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
            overwritePolicy: OverwritePolicy.Allow));
        RequireSuccess(_commands.SetNumberFormat(batch, _salesSheetName, "C2:C6", "0"));
        RequireSuccess(_commands.SetNumberFormat(batch, _salesSheetName, "D2:D6", "m/d/yyyy"));
    }

    private void AssertGroupedValues(string pivotName, double[] expectedValues, double total)
    {
        var values = RequireSuccess(_pivotCommands.GetData(_fixture.BatchToken, pivotName)).Values;
        Assert.True(values.Count == expectedValues.Length + 2,
            $"Unexpected grouped grid: {System.Text.Json.JsonSerializer.Serialize(values)}");
        Assert.All(values, row => Assert.Equal(2, row.Count));
        Assert.Equal(expectedValues, values.Skip(1).Take(expectedValues.Length)
            .Select(row => Convert.ToDouble(row[1], System.Globalization.CultureInfo.InvariantCulture)));
        Assert.Equal(total, Convert.ToDouble(values[^1][^1], System.Globalization.CultureInfo.InvariantCulture));
    }

    private void AssertNumericGroups(string pivotName, string[] expectedLabels)
    {
        AssertGroupedValues(pivotName, [1d, 1d, 1d, 1d, 1d], 5);
        var values = RequireSuccess(_pivotCommands.GetData(_fixture.BatchToken, pivotName)).Values;
        Assert.Equal(expectedLabels, values.Skip(1).Take(5).Select(row => row[0]?.ToString()));
        var source = RequireSuccess(_commands.GetValues(_fixture.BatchToken, _salesSheetName, "C2:C6"));
        Assert.Equal([150d, 250d, 450d, 600d, 850d],
            source.Values.Select(row => Convert.ToDouble(Assert.Single(row),
                System.Globalization.CultureInfo.InvariantCulture)));
    }
}
