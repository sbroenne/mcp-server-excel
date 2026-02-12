using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Commands;
using Sbroenne.ExcelMcp.Core.Models;
using Sbroenne.ExcelMcp.Core.Tests.Helpers;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Commands.PivotTable;

/// <summary>
/// Tests for OLAP/Data Model PivotTable field operations (Strategy Pattern: OlapPivotTableFieldStrategy).
/// Verifies that all field manipulation methods work correctly with Data Model PivotTables.
/// Uses CubeFields API via GetFieldForManipulation() helper.
/// Organized as partial class for consistency with Strategy Pattern architecture.
/// </summary>
public partial class PivotTableCommandsTests
{
    /// <summary>
    /// OLAP-specific tests use fixture to provide Data Model PivotTable.
    /// All OLAP tests marked with [Trait("Category", "OLAP")] for strategy classification.
    /// </summary>

    [Fact]
    [Trait("Speed", "Medium")]
    [Trait("Category", "OLAP")]
    public void AddRowField_OlapPivot_AddsFieldToRows()
    {
        // Arrange - Create OLAP test file with data model
        var olapTestFile = CreateOlapTestFile(nameof(AddRowField_OlapPivot_AddsFieldToRows));
        using var batch = ExcelSession.BeginBatch(olapTestFile);

        // Act - Remove existing Region field first, then add Quarter
        // Use exact CubeField names (LLM discovers via ListFields)
        _pivotCommands.RemoveField(batch, "DataModelPivot", "[RegionalSalesTable].[Region]");
        var result = _pivotCommands.AddRowField(batch, "DataModelPivot", "[RegionalSalesTable].[Quarter]", null);

        // Assert
        Assert.True(result.Success, $"Failed: {result.ErrorMessage}");
        Assert.Equal("[RegionalSalesTable].[Quarter]", result.FieldName);
        Assert.Equal(PivotFieldArea.Row, result.Area);
    }

    [Fact]
    [Trait("Speed", "Medium")]
    [Trait("Category", "OLAP")]
    public void AddColumnField_OlapPivot_AddsFieldToColumns()
    {
        // Arrange - Create OLAP test file with data model
        var olapTestFile = CreateOlapTestFile(nameof(AddColumnField_OlapPivot_AddsFieldToColumns));
        using var batch = ExcelSession.BeginBatch(olapTestFile);

        // Act - Remove existing Region field first to make room for Quarter
        // Use exact CubeField names (LLM discovers via ListFields)
        _pivotCommands.RemoveField(batch, "DataModelPivot", "[RegionalSalesTable].[Region]");
        var result = _pivotCommands.AddColumnField(batch, "DataModelPivot", "[RegionalSalesTable].[Quarter]", null);

        // Assert
        Assert.True(result.Success, $"Failed: {result.ErrorMessage}");
        Assert.Equal("[RegionalSalesTable].[Quarter]", result.FieldName);
        Assert.Equal(PivotFieldArea.Column, result.Area);
    }

    [Fact]
    [Trait("Speed", "Medium")]
    [Trait("Category", "OLAP")]
    public void SortField_OlapPivot_SortsFieldSuccessfully()
    {
        // Arrange - Create OLAP test file with data model
        var olapTestFile = CreateOlapTestFile(nameof(SortField_OlapPivot_SortsFieldSuccessfully));
        using var batch = ExcelSession.BeginBatch(olapTestFile);

        // Act - Region row field exists in fixture
        // Use exact CubeField name (LLM discovers via ListFields)
        var result = _pivotCommands.SortField(
            batch,
            "DataModelPivot",
            "[RegionalSalesTable].[Region]",
            SortDirection.Descending);

        // Assert
        Assert.True(result.Success, $"Failed: {result.ErrorMessage}");
        Assert.Equal("[RegionalSalesTable].[Region]", result.FieldName);
    }

    /// <summary>
    /// Regression test for Issue #217: Auto-create DAX measures when adding value fields to OLAP PivotTables.
    ///
    /// CURRENT BEHAVIOR: AddValueField on OLAP PivotTable always fails with:
    ///   "Cannot add value field to OLAP PivotTable. OLAP measures must be pre-defined..."
    ///
    /// EXPECTED BEHAVIOR: AddValueField should auto-create DAX measure and add to values area.
    ///
    /// This test is expected to FAIL initially, then PASS after implementing auto-DAX-creation.
    /// </summary>
    [Fact]
    [Trait("Speed", "Medium")]
    [Trait("Category", "OLAP")]
    public void AddValueField_OlapPivot_AutoCreatesDaxMeasure()
    {
        // Arrange - Create OLAP test file with Data Model PivotTable
        var olapTestFile = CreateOlapTestFile(nameof(AddValueField_OlapPivot_AutoCreatesDaxMeasure));
        using var batch = ExcelSession.BeginBatch(olapTestFile);

        // Act - Try to add Sales field as a Sum value field
        // Use exact CubeField name format [TableName].[ColumnName]
        // After implementation, should auto-create: [Regional Sales Total] = SUM('RegionalSalesTable'[Sales])
        // NOTE: Use unique name to avoid conflict with fixture's "Total Sales" measure on SalesTable
        var result = _pivotCommands.AddValueField(
            batch,
            "DataModelPivot",
            "[RegionalSalesTable].[Sales]",
            AggregationFunction.Sum,
            "Regional Sales Total");

        // Assert - Should succeed with auto-created DAX measure
        Assert.True(result.Success, $"AddValueField should auto-create DAX measure but failed: {result.ErrorMessage}");
        Assert.Equal("Regional Sales Total", result.FieldName); // Field name is the measure name
        Assert.Equal(PivotFieldArea.Value, result.Area);
        Assert.Equal("Regional Sales Total", result.CustomName);

        // Verify the DAX measure was created in Data Model
        var dataModelCommands = new DataModelCommands();
        var measuresResult = dataModelCommands.ListMeasures(batch, "RegionalSalesTable");
        Assert.True(measuresResult.Success, $"Failed to list measures: {measuresResult.ErrorMessage}");

        // Should contain either the auto-created measure or use the custom name
        var hasMeasure = measuresResult.Measures.Any(m =>
            m.Name.Contains("Regional Sales Total", StringComparison.OrdinalIgnoreCase));
        Assert.True(hasMeasure, "Auto-created DAX measure should exist in Data Model");
    }

    /// <summary>
    /// Test auto-creation of DAX measure with Count aggregation function.
    /// Verifies that different aggregation functions generate correct DAX formulas.
    /// </summary>
    [Fact]
    [Trait("Speed", "Medium")]
    [Trait("Category", "OLAP")]
    public void AddValueField_OlapPivot_AutoCreatesDaxMeasureWithCount()
    {
        // Arrange
        var olapTestFile = CreateOlapTestFile(nameof(AddValueField_OlapPivot_AutoCreatesDaxMeasureWithCount));
        using var batch = ExcelSession.BeginBatch(olapTestFile);

        // Act - Add Quarter field with Count aggregation
        // Use exact CubeField name format [TableName].[ColumnName]
        // Should auto-create: [Number of Quarters] = COUNT('RegionalSalesTable'[Quarter])
        var result = _pivotCommands.AddValueField(
            batch,
            "DataModelPivot",
            "[RegionalSalesTable].[Quarter]",
            AggregationFunction.Count,
            "Number of Quarters");

        // Assert
        Assert.True(result.Success, $"AddValueField with Count should auto-create DAX measure but failed: {result.ErrorMessage}");
        Assert.Equal("Number of Quarters", result.FieldName); // Field name is the measure name
        Assert.Equal(PivotFieldArea.Value, result.Area);
        Assert.Equal(AggregationFunction.Count, result.Function);

        // Verify the DAX measure was created with COUNT function
        var dataModelCommands = new DataModelCommands();
        var measuresResult = dataModelCommands.ListMeasures(batch, "RegionalSalesTable");
        Assert.True(measuresResult.Success, $"Failed to list measures: {measuresResult.ErrorMessage}");
        var hasCountMeasure = measuresResult.Measures.Any(m =>
            m.Name.Contains("Quarter", StringComparison.OrdinalIgnoreCase) &&
            (m.FormulaPreview?.Contains("COUNT", StringComparison.OrdinalIgnoreCase) ?? false));
        Assert.True(hasCountMeasure, "Auto-created COUNT measure should exist in Data Model");
    }

    /// <summary>
    /// Test adding a pre-existing measure to PivotTable values area.
    /// This is the core scenario from the issue: user has a measure in Data Model and wants to add it to PivotTable.
    /// Measure formats: "[Measures].[Name]", "Name", or CubeField name
    /// </summary>
    [Fact]
    [Trait("Speed", "Medium")]
    [Trait("Category", "OLAP")]
    public void AddValueField_OlapPivot_AddsPreExistingMeasure()
    {
        // Arrange - Create OLAP test file and add a measure first
        var olapTestFile = CreateOlapTestFile(nameof(AddValueField_OlapPivot_AddsPreExistingMeasure));
        using var batch = ExcelSession.BeginBatch(olapTestFile);

        // First, create a measure in the Data Model (not in PivotTable yet)
        var dataModelCommands = new DataModelCommands();
        dataModelCommands.CreateMeasure(
            batch,
            "RegionalSalesTable",
            "Total ACR",
            "SUM('RegionalSalesTable'[Sales])",
            null);  // CreateMeasure throws on error

        // Refresh PivotTable to pick up the new measure in CubeFields
        _pivotCommands.Refresh(batch, "DataModelPivot", null);

        // Act - Add the pre-existing measure to PivotTable values area
        // Should detect it's an existing measure and just set Orientation = xlDataField
        var result = _pivotCommands.AddValueField(
            batch,
            "DataModelPivot",
            "Total ACR", // Can use measure name directly
            AggregationFunction.Sum, // Ignored for pre-existing measures
            null);

        // Assert - Should succeed without creating a new measure
        Assert.True(result.Success, $"AddValueField should add existing measure but failed: {result.ErrorMessage}");
        Assert.Equal("Total ACR", result.FieldName);
        Assert.Equal(PivotFieldArea.Value, result.Area);

        // Verify only ONE measure with this name exists (not duplicated)
        var measuresResult = dataModelCommands.ListMeasures(batch, "RegionalSalesTable");
        Assert.True(measuresResult.Success, $"Failed to list measures: {measuresResult.ErrorMessage}");
        var measureCount = measuresResult.Measures.Count(m => m.Name == "Total ACR");
        Assert.Equal(1, measureCount); // Should still be 1, not 2
    }

    /// <summary>
    /// Test adding pre-existing measure using [Measures].[Name] format.
    /// This format is commonly used in OLAP/MDX contexts and should be supported.
    /// </summary>
    [Fact]
    [Trait("Speed", "Medium")]
    [Trait("Category", "OLAP")]
    public void AddValueField_OlapPivot_AddsPreExistingMeasureWithMeasuresPrefix()
    {
        // Arrange - Create measure first
        var olapTestFile = CreateOlapTestFile(nameof(AddValueField_OlapPivot_AddsPreExistingMeasureWithMeasuresPrefix));
        using var batch = ExcelSession.BeginBatch(olapTestFile);

        var dataModelCommands = new DataModelCommands();
        dataModelCommands.CreateMeasure(
            batch,
            "RegionalSalesTable",
            "Revenue Total",
            "SUM('RegionalSalesTable'[Sales])",
            null);  // CreateMeasure throws on error

        _pivotCommands.Refresh(batch, "DataModelPivot", null);

        // Act - Use [Measures].[Name] format (common in OLAP contexts)
        var result = _pivotCommands.AddValueField(
            batch,
            "DataModelPivot",
            "[Measures].[Revenue Total]", // MDX-style format
            AggregationFunction.Sum,
            null);

        // Assert
        Assert.True(result.Success, $"Should handle [Measures].[Name] format but failed: {result.ErrorMessage}");
        Assert.Equal("Revenue Total", result.FieldName);
        Assert.Equal(PivotFieldArea.Value, result.Area);
    }


    /// <summary>
    /// Helper to get the OLAP test file path from shared DataModelPivotTableFixture.
    /// Uses shared fixture instead of creating new one each time (massive performance improvement).
    /// </summary>
    private string CreateOlapTestFile(string _)
    {
        // Use the shared fixture from [Collection("DataModel")] - created ONCE for all test classes
        // This is initialized ONCE per test run, not per test method or per test class
        return _olapFixture.TestFilePath;
    }

    /// <summary>
    /// Test UPDATE: Change aggregation function for existing OLAP value field.
    /// Verifies that SetFieldFunction modifies the DAX measure formula in Data Model.
    /// </summary>
    [Fact]
    [Trait("Speed", "Medium")]
    [Trait("Category", "OLAP")]
    public void SetFieldFunction_OlapPivot_UpdatesDaxMeasureFormula()
    {
        // Arrange - Create measure with SUM first
        var olapTestFile = CreateOlapTestFile(nameof(SetFieldFunction_OlapPivot_UpdatesDaxMeasureFormula));
        using var batch = ExcelSession.BeginBatch(olapTestFile);

        // Use exact CubeField name format [TableName].[ColumnName]
        var addResult = _pivotCommands.AddValueField(
            batch,
            "DataModelPivot",
            "[RegionalSalesTable].[Sales]",
            AggregationFunction.Sum,
            "Sales Measure");
        Assert.True(addResult.Success, $"Setup failed: {addResult.ErrorMessage}");

        // Act - Change from SUM to COUNT
        // After the measure is created, reference it by its measure name or [Measures].[Name]
        var updateResult = _pivotCommands.SetFieldFunction(
            batch,
            "DataModelPivot",
            "[Measures].[Sales Measure]",
            AggregationFunction.Count);

        // Assert - Operation succeeded
        Assert.True(updateResult.Success, $"Update failed: {updateResult.ErrorMessage}");
        Assert.Contains("Sales Measure", updateResult.FieldName);
        Assert.Equal(AggregationFunction.Count, updateResult.Function);

        // Verify the DAX measure formula changed in Data Model
        var dataModelCommands = new DataModelCommands();
        var measuresResult = dataModelCommands.ListMeasures(batch, "RegionalSalesTable");
        Assert.True(measuresResult.Success, $"Failed to list measures: {measuresResult.ErrorMessage}");

        var measure = measuresResult.Measures.FirstOrDefault(m => m.Name == "Sales Measure");
        Assert.NotNull(measure);
        Assert.Contains("COUNT", measure.FormulaPreview, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("SUM", measure.FormulaPreview, StringComparison.OrdinalIgnoreCase);
    }

    /// <summary>
    /// Test UPDATE: Change number format for existing OLAP value field.
    /// Verifies that SetFieldFormat modifies the measure's format in Data Model.
    /// </summary>
    [Fact]
    [Trait("Speed", "Medium")]
    [Trait("Category", "OLAP")]
    public void SetFieldFormat_OlapPivot_UpdatesMeasureFormat()
    {
        // Arrange - Create measure first
        var olapTestFile = CreateOlapTestFile(nameof(SetFieldFormat_OlapPivot_UpdatesMeasureFormat));
        using var batch = ExcelSession.BeginBatch(olapTestFile);

        // Use exact CubeField name format [TableName].[ColumnName]
        var addResult = _pivotCommands.AddValueField(
            batch,
            "DataModelPivot",
            "[RegionalSalesTable].[Sales]",
            AggregationFunction.Sum,
            "Sales Total");
        Assert.True(addResult.Success, $"Setup failed: {addResult.ErrorMessage}");

        // Act - Set a simple format that Excel preserves exactly
        // After the measure is created, reference it by [Measures].[Name]
        // Use "0%" which is locale-independent
        var updateResult = _pivotCommands.SetFieldFormat(
            batch,
            "DataModelPivot",
            "[Measures].[Sales Total]",
            "0%");

        // Assert - Operation succeeded
        Assert.True(updateResult.Success, $"Update failed: {updateResult.ErrorMessage}");
        Assert.Contains("Sales Total", updateResult.FieldName);
        Assert.Equal("0%", updateResult.NumberFormat);
    }

    /// <summary>
    /// Test UPDATE: Format a PRE-EXISTING measure (not created in same test).
    /// This covers the bug scenario where SetFieldFormat failed for measures 
    /// created via excel_datamodel tool, which exist in CubeFields but not 
    /// in the same code path as AddValueField-created measures.
    /// 
    /// BUG REGRESSION TEST: The old SetFieldFormat searched model.ModelMeasures 
    /// but pre-existing measures may not be there in the expected format.
    /// The fix uses CubeField.PivotFields[1].NumberFormat directly.
    /// </summary>
    [Fact]
    [Trait("Speed", "Medium")]
    [Trait("Category", "OLAP")]
    public void SetFieldFormat_PreExistingMeasure_FormatsSuccessfully()
    {
        // Arrange - Use fixture which has pre-existing "ACR" measure on DisambiguationTable
        // This measure was created via dataModelCommands.CreateMeasure() during fixture init
        // NOT via AddValueField in this test - simulating real-world scenario
        var olapTestFile = CreateOlapTestFile(nameof(SetFieldFormat_PreExistingMeasure_FormatsSuccessfully));
        using var batch = ExcelSession.BeginBatch(olapTestFile);

        // First, add the pre-existing ACR measure to the DisambiguationTest PivotTable
        // The measure exists in Data Model but needs to be added to PivotTable's Values area
        var addResult = _pivotCommands.AddValueField(
            batch,
            "DisambiguationTest",
            "[Measures].[ACR]",  // Pre-existing measure from fixture
            AggregationFunction.Sum,  // Ignored for existing measures
            null);  // Keep existing name
        Assert.True(addResult.Success, $"AddValueField failed: {addResult.ErrorMessage}");

        // Act - Format the pre-existing measure (this was the bug scenario)
        // Use "0%" format which is locale-independent
        var formatResult = _pivotCommands.SetFieldFormat(
            batch,
            "DisambiguationTest",
            "[Measures].[ACR]",
            "0%");

        // Assert - Operation succeeded (was failing with "Measure not found in Data Model")
        Assert.True(formatResult.Success, $"SetFieldFormat failed: {formatResult.ErrorMessage}");
        Assert.Contains("ACR", formatResult.FieldName);
        Assert.Equal("0%", formatResult.NumberFormat);
    }
}




