using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.Core.Models;
using Xunit;
using Xunit.Abstractions;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sbroenne.ExcelMcp.Service.Tests;

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
[Collection("ServiceWorkflow")]
[Trait("Layer", "Service")]
[Trait("Category", "Integration")]
[Trait("Feature", "PivotTables")]
[Trait("RequiresExcel", "true")]
public class PersistentServicePivotTableOlapDisambiguationTests :
    PersistentServiceWorkbookTestBase,
    IClassFixture<PersistentServiceDataModelFixture>
{
    private readonly IPersistentPivotTableCommands _pivotCommands;
    private readonly ITestOutputHelper _output;
    private string _pivotName = string.Empty;
    private string _pivotSheet = string.Empty;

    public PersistentServicePivotTableOlapDisambiguationTests(
        PersistentServiceDataModelFixture fixture,
        ITestOutputHelper output) :
        base(fixture)
    {
        _pivotCommands =
            fixture.CreateCommands<IPersistentPivotTableCommands>();
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
        Assert.True(PersistentServiceDataModelFixture.CreationResult.Success, $"Fixture creation failed: {PersistentServiceDataModelFixture.CreationResult.ErrorMessage}");
        var batch = _fixture.BatchToken;
        PreparePivot();

        // Act - Try to add the DAX measure using [Measures].[Name] syntax
        var result = _pivotCommands.AddValueField(
            batch,
            _pivotName,
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
        AssertNativeMeasure("ACR", 8800);
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
        Assert.True(PersistentServiceDataModelFixture.CreationResult.Success, $"Fixture creation failed: {PersistentServiceDataModelFixture.CreationResult.ErrorMessage}");
        var batch = _fixture.BatchToken;
        PreparePivot();

        // Act - Try to add the DAX measure using exact name (no [Measures]. prefix)
        var result = _pivotCommands.AddValueField(
            batch,
            _pivotName,
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
        AssertNativeMeasure("Discount", 880);
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
        Assert.True(PersistentServiceDataModelFixture.CreationResult.Success, $"Fixture creation failed: {PersistentServiceDataModelFixture.CreationResult.ErrorMessage}");
        var batch = _fixture.BatchToken;

        // Act - Enumerate CubeFields and check their types
        var measureFields = new List<(string Name, int CubeFieldType)>();
        var hierarchyFields = new List<(string Name, int CubeFieldType)>();

        _fixture.ExecuteRawVerification((ctx, ct) =>
        {
            Excel.Sheets? sheets = null;
            Excel.Worksheet? sheet = null;
            Excel.PivotTables? pivots = null;
            Excel.PivotTable? pivotTable = null;
            Excel.CubeFields? cubeFields = null;
            try
            {
                sheets = ctx.Book.Worksheets;
                sheet = (Excel.Worksheet)sheets["DisambiguationPivot"];
                pivots = (Excel.PivotTables)sheet.PivotTables();
                pivotTable = pivots.Item("DisambiguationTest");
                cubeFields = pivotTable.CubeFields;
                for (int i = 1; i <= cubeFields.Count; i++)
                {
                    Excel.CubeField? field = null;
                    try
                    {
                        field = cubeFields[i];
                        if (field.CubeFieldType == Excel.XlCubeFieldType.xlMeasure)
                            measureFields.Add((field.Name, (int)field.CubeFieldType));
                        else if (field.CubeFieldType == Excel.XlCubeFieldType.xlHierarchy)
                            hierarchyFields.Add((field.Name, (int)field.CubeFieldType));
                    }
                    finally
                    {
                        ComUtilities.Release(ref field);
                    }
                }
            }
            finally
            {
                ComUtilities.Release(ref cubeFields);
                ComUtilities.Release(ref pivotTable);
                ComUtilities.Release(ref pivots);
                ComUtilities.Release(ref sheet);
                ComUtilities.Release(ref sheets);
            }
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
        Assert.Contains(("[Measures].[ACR]", 2), measureFields);
        Assert.Contains(("[Measures].[Discount]", 2), measureFields);
        Assert.Contains(hierarchyFields, field => field.Name == "[DisambiguationTable].[ACRTypeKey]");
        Assert.Contains(hierarchyFields, field => field.Name == "[DisambiguationTable].[DiscountCode]");

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
        Assert.True(PersistentServiceDataModelFixture.CreationResult.Success, $"Fixture creation failed: {PersistentServiceDataModelFixture.CreationResult.ErrorMessage}");
        var batch = _fixture.BatchToken;
        PreparePivot();

        // First, add the measure to values area (use unambiguous measure name)
        var addResult = _pivotCommands.AddValueField(
            batch,
            _pivotName,
            "[Measures].[ACR]",
            AggregationFunction.Sum,
            null);

        RequireSuccess(addResult);
        _output.WriteLine($"AddValueField result: Success={addResult.Success}, FieldName={addResult.FieldName}");

        // Act - List all fields
        var listResult = _pivotCommands.ListFields(batch, _pivotName);

        // Assert
        Assert.True(listResult.Success, $"ListFields failed: {listResult.ErrorMessage}");

        _output.WriteLine($"Fields in PivotTable:");
        foreach (var field in listResult.Fields)
        {
            _output.WriteLine($"  - {field.Name}: Area={field.Area}");
        }

        var measureField = Assert.Single(listResult.Fields,
            field => field.Name == "[Measures].[ACR]");
        Assert.Equal(PivotFieldArea.Value, measureField.Area);
        AssertNativeMeasure("ACR", 8800);
    }

    private void PreparePivot()
    {
        _pivotSheet = _fixture.CreateTestSheet(_fixture.BatchToken);
        _pivotName = $"Disambiguation_{Guid.NewGuid():N}";
        RequireSuccess(_pivotCommands.CreateFromDataModel(
            _fixture.BatchToken, "DisambiguationTable", _pivotSheet, "A1", _pivotName));
    }

    private void AssertNativeMeasure(string measureName, double expectedValue)
    {
        RequireSuccess(_pivotCommands.Refresh(_fixture.BatchToken, _pivotName, null));
        _fixture.ExecuteRawVerification((context, _) =>
        {
            Excel.Sheets? sheets = null;
            Excel.Worksheet? sheet = null;
            Excel.PivotTables? pivots = null;
            Excel.PivotTable? pivot = null;
            Excel.PivotCache? cache = null;
            Excel.PivotFields? fields = null;
            Excel.PivotField? field = null;
            Excel.CubeField? cubeField = null;
            Excel.Range? data = null;
            try
            {
                sheets = context.Book.Worksheets;
                sheet = (Excel.Worksheet)sheets[_pivotSheet];
                pivots = (Excel.PivotTables)sheet.PivotTables();
                pivot = pivots.Item(_pivotName);
                cache = pivot.PivotCache();
                Assert.True(cache.OLAP);
                fields = (Excel.PivotFields)pivot.DataFields;
                Assert.Equal(1, fields.Count);
                field = fields.Item(1);
                cubeField = field.CubeField;
                Assert.Equal($"[Measures].[{measureName}]", cubeField.Name);
                Assert.Equal(Excel.XlCubeFieldType.xlMeasure, cubeField.CubeFieldType);
                Assert.Equal(Excel.XlPivotFieldOrientation.xlDataField, field.Orientation);
                data = pivot.DataBodyRange;
                Assert.NotNull(data);
                Assert.Equal(expectedValue, Convert.ToDouble(data.Value2,
                    System.Globalization.CultureInfo.InvariantCulture), 8);
            }
            finally
            {
                ComUtilities.Release(ref data);
                ComUtilities.Release(ref cubeField);
                ComUtilities.Release(ref field);
                ComUtilities.Release(ref fields);
                ComUtilities.Release(ref cache);
                ComUtilities.Release(ref pivot);
                ComUtilities.Release(ref pivots);
                ComUtilities.Release(ref sheet);
                ComUtilities.Release(ref sheets);
            }
        });
    }
}
