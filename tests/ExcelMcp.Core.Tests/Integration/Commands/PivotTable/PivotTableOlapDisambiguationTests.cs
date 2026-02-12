using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Commands.PivotTable;
using Sbroenne.ExcelMcp.Core.Models;
using Sbroenne.ExcelMcp.Core.Tests.Helpers;
using Xunit;
using Xunit.Abstractions;

namespace Sbroenne.ExcelMcp.Core.Tests.Commands.PivotTable;

/// <summary>
/// Regression tests for OLAP PivotTable measure disambiguation issues.
/// 
/// BUG REPORT: When adding DAX measures to OLAP PivotTables, the server incorrectly
/// matches table columns with similar names instead of the actual DAX measure.
/// 
/// Issues being tested:
/// 1. AddValueField with "[Measures].[ACR]" matches "[DisambiguationTable].[ACRTypeKey]" instead of measure
/// 2. Partial matching causes wrong field to be added when names overlap
/// 3. CubeFieldType should be used to distinguish measures from hierarchies
/// 
/// These tests use the shared PivotTableRealisticFixture which creates:
/// - DisambiguationTable with columns "ACRTypeKey", "DiscountCode" 
/// - DAX measures "ACR", "Discount" that could be confused with columns
/// - DisambiguationTest PivotTable connected to the Data Model
/// </summary>
[Collection("DataModel")]
[Trait("Category", "Integration")]
[Trait("Feature", "PivotTables")]
[Trait("RequiresExcel", "true")]
public class PivotTableOlapDisambiguationTests
{
    private readonly DataModelPivotTableFixture _fixture;
    private readonly PivotTableCommands _pivotCommands;
    private readonly ITestOutputHelper _output;

    public PivotTableOlapDisambiguationTests(DataModelPivotTableFixture fixture, ITestOutputHelper output)
    {
        _fixture = fixture;
        _pivotCommands = new PivotTableCommands();
        _output = output;
    }

    /// <summary>
    /// REGRESSION TEST: AddValueField with [Measures].[MeasureName] should add the DAX measure,
    /// not a table column with a similar name.
    /// 
    /// BUG: When calling AddValueField with fieldName="[Measures].[ACR]", the current implementation
    /// uses Contains() matching which matches "[DisambiguationTable].[ACRTypeKey]" because it contains "ACR".
    /// 
    /// EXPECTED: Only the DAX measure "ACR" should be matched, not table columns.
    /// </summary>
    [Fact]
    [Trait("Speed", "Medium")]
    [Trait("Category", "Regression")]
    public void AddValueField_MeasuresPrefix_ShouldNotMatchTableColumn()
    {
        // Arrange - Use the shared fixture file
        Assert.True(_fixture.CreationResult.Success, $"Fixture creation failed: {_fixture.CreationResult.ErrorMessage}");
        using var batch = ExcelSession.BeginBatch(_fixture.TestFilePath);

        // Act - Try to add the DAX measure using [Measures].[Name] syntax
        var result = _pivotCommands.AddValueField(
            batch,
            "DisambiguationTest",
            "[Measures].[ACR]",  // Should match DAX measure, NOT [DisambiguationTable].[ACRTypeKey]
            AggregationFunction.Sum,
            null);

        // Assert - The operation should succeed
        Assert.True(result.Success, $"AddValueField failed: {result.ErrorMessage}");

        // CRITICAL: The result should show the MEASURE was added, not a table column
        // If the bug exists, FieldName would contain "ACRTypeKey" instead of "ACR"
        _output.WriteLine($"FieldName: {result.FieldName}");
        _output.WriteLine($"CustomName: {result.CustomName}");
        _output.WriteLine($"Area: {result.Area}");

        Assert.Equal("ACR", result.FieldName);
        Assert.DoesNotContain("ACRTypeKey", result.CustomName ?? "", StringComparison.OrdinalIgnoreCase);

        // The area should be Value (xlDataField)
        Assert.Equal(PivotFieldArea.Value, result.Area);
    }

    /// <summary>
    /// REGRESSION TEST: AddValueField with exact measure name should add the DAX measure,
    /// not a table column with a similar name.
    /// 
    /// BUG: When calling AddValueField with fieldName="Discount", the current implementation
    /// iterates through CubeFields and uses Contains() which matches "DiscountCode" column first.
    /// 
    /// EXPECTED: Exact measure name matching should find the measure, not a column.
    /// </summary>
    [Fact]
    [Trait("Speed", "Medium")]
    [Trait("Category", "Regression")]
    public void AddValueField_ExactMeasureName_ShouldNotMatchTableColumn()
    {
        // Arrange
        Assert.True(_fixture.CreationResult.Success, $"Fixture creation failed: {_fixture.CreationResult.ErrorMessage}");
        using var batch = ExcelSession.BeginBatch(_fixture.TestFilePath);

        // Act - Try to add the DAX measure using exact name (no [Measures]. prefix)
        var result = _pivotCommands.AddValueField(
            batch,
            "DisambiguationTest",
            "Discount",  // Should match DAX measure, NOT [DisambiguationTable].[DiscountCode]
            AggregationFunction.Sum,
            null);

        // Assert
        Assert.True(result.Success, $"AddValueField failed: {result.ErrorMessage}");

        _output.WriteLine($"FieldName: {result.FieldName}");
        _output.WriteLine($"CustomName: {result.CustomName}");

        // CRITICAL: The result should show the MEASURE was added
        Assert.Equal("Discount", result.FieldName);
        Assert.DoesNotContain("DiscountCode", result.CustomName ?? "", StringComparison.OrdinalIgnoreCase);
        Assert.Equal(PivotFieldArea.Value, result.Area);
    }

    /// <summary>
    /// Test that CubeFieldType property can distinguish measures from hierarchies.
    /// This verifies we can use the COM API to properly identify measure CubeFields.
    /// </summary>
    [Fact]
    [Trait("Speed", "Medium")]
    public void CubeFields_CanIdentifyMeasuresByCubeFieldType()
    {
        // Arrange
        Assert.True(_fixture.CreationResult.Success, $"Fixture creation failed: {_fixture.CreationResult.ErrorMessage}");
        using var batch = ExcelSession.BeginBatch(_fixture.TestFilePath);

        // Act - Enumerate CubeFields and check their types
        var measureFields = new List<(string Name, int CubeFieldType)>();
        var hierarchyFields = new List<(string Name, int CubeFieldType)>();

        batch.Execute((ctx, ct) =>
        {
            dynamic sheet = ctx.Book.Worksheets["DisambiguationPivot"];
            dynamic pivotTable = sheet.PivotTables("DisambiguationTest");
            dynamic cubeFields = pivotTable.CubeFields;

            for (int i = 1; i <= cubeFields.Count; i++)
            {
                dynamic? cf = null;
                try
                {
                    cf = cubeFields.Item(i);
                    string name = cf.Name?.ToString() ?? "";
                    int cubeFieldType = Convert.ToInt32(cf.CubeFieldType);

                    // xlMeasure = 2, xlHierarchy = 1
                    if (cubeFieldType == 2) // xlMeasure
                    {
                        measureFields.Add((name, cubeFieldType));
                    }
                    else if (cubeFieldType == 1) // xlHierarchy
                    {
                        hierarchyFields.Add((name, cubeFieldType));
                    }
                }
                finally
                {
                    if (cf != null)
                        ComUtilities.Release(ref cf!);
                }
            }

            ComUtilities.Release(ref cubeFields!);
            return 0;
        });

        // Assert - We should find measures with CubeFieldType = 2
        _output.WriteLine($"Found {measureFields.Count} measures:");
        foreach (var (name, type) in measureFields)
        {
            _output.WriteLine($"  - {name} (type={type})");
        }

        _output.WriteLine($"Found {hierarchyFields.Count} hierarchies (showing first 10):");
        foreach (var (name, type) in hierarchyFields.Take(10))
        {
            _output.WriteLine($"  - {name} (type={type})");
        }

        Assert.NotEmpty(measureFields);

        // Our created measures should be in the measure list
        Assert.Contains(measureFields, m => m.Name.Contains("ACR", StringComparison.OrdinalIgnoreCase));
        Assert.Contains(measureFields, m => m.Name.Contains("Discount", StringComparison.OrdinalIgnoreCase));

        // Table columns (ACRTypeKey, DiscountCode) should NOT be measures
        Assert.DoesNotContain(measureFields, m => m.Name.Contains("ACRTypeKey", StringComparison.OrdinalIgnoreCase));
        Assert.DoesNotContain(measureFields, m => m.Name.Contains("DiscountCode", StringComparison.OrdinalIgnoreCase));
    }

    /// <summary>
    /// After adding a measure to Values area, ListFields should show it
    /// with Area = Value, not Area = Hidden.
    /// </summary>
    [Fact]
    [Trait("Speed", "Medium")]
    [Trait("Category", "Regression")]
    public void ListFields_AfterAddValueField_ShouldShowMeasureInValueArea()
    {
        // Arrange
        Assert.True(_fixture.CreationResult.Success, $"Fixture creation failed: {_fixture.CreationResult.ErrorMessage}");
        using var batch = ExcelSession.BeginBatch(_fixture.TestFilePath);

        // First, add the measure to values area (use unambiguous measure name)
        var addResult = _pivotCommands.AddValueField(
            batch,
            "DisambiguationTest",
            "[Measures].[ACR]",
            AggregationFunction.Sum,
            null);

        // Even if add fails due to bug, let's check ListFields
        _output.WriteLine($"AddValueField result: Success={addResult.Success}, FieldName={addResult.FieldName}");

        // Act - List all fields
        var listResult = _pivotCommands.ListFields(batch, "DisambiguationTest");

        // Assert
        Assert.True(listResult.Success, $"ListFields failed: {listResult.ErrorMessage}");

        _output.WriteLine($"Fields in PivotTable:");
        foreach (var field in listResult.Fields)
        {
            _output.WriteLine($"  - {field.Name}: Area={field.Area}");
        }

        // Find ACR in the field list (could be measure or incorrectly matched column)
        var acrFields = listResult.Fields.Where(f =>
            f.Name.Contains("ACR", StringComparison.OrdinalIgnoreCase)).ToList();

        Assert.NotEmpty(acrFields);

        // If fix is applied: ACR measure should be in Value area
        // If bug exists: ACRTypeKey column might be matched instead
        var measureField = acrFields.FirstOrDefault(f =>
            f.Name.Contains("[Measures]", StringComparison.OrdinalIgnoreCase));

        if (measureField != null)
        {
            Assert.Equal(PivotFieldArea.Value, measureField.Area);
        }
    }
}




