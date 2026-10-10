using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.Core.Models;
using Xunit;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sbroenne.ExcelMcp.Service.Tests;

/// <summary>
/// Tests for OLAP/Data Model PivotTable field operations (Strategy Pattern: OlapPivotTableFieldStrategy).
/// Verifies that all field manipulation methods work correctly with Data Model PivotTables.
/// Uses CubeFields API via GetFieldForManipulation() helper.
/// Organized as partial class for consistency with Strategy Pattern architecture.
/// </summary>
public partial class PersistentServicePivotTableOlapFieldTests
{
    [Fact]
    [Trait("Speed", "Medium")]
    [Trait("Category", "OLAP")]
    public void CacheOptions_OlapCache_RejectsUnsupportedMutations()
    {
        Assert.True(_creationResult.Success, _creationResult.ErrorMessage);
        var batch = _fixture.BatchToken;
        var (_, pivotName) = CreateIsolatedPivot();
        var current = _pivotCommands.GetCacheOptions(batch, pivotName);
        Assert.True(current.Success, current.ErrorMessage);
        Assert.True(current.IsOlap);

        var exception = Assert.Throws<InvalidOperationException>(() =>
            _pivotCommands.SetCacheOptions(
                batch,
                pivotName,
                missingItemsLimit: PivotMissingItemsLimit.None));
        Assert.Contains("not available for OLAP", exception.Message);
        AssertCachePreserved();

        exception = Assert.Throws<InvalidOperationException>(() =>
            _pivotCommands.SetCacheOptions(
                batch,
                pivotName,
                optimizeCache: !current.OptimizeCache));
        Assert.Contains("read-only for external OLE DB/OLAP", exception.Message);
        AssertCachePreserved();

        exception = Assert.Throws<InvalidOperationException>(() =>
            _pivotCommands.SetCacheOptions(
                batch,
                pivotName,
                saveSourceData: true));
        Assert.Contains("cannot save source records", exception.Message);
        AssertCachePreserved();

        void AssertCachePreserved()
        {
            var retained = _pivotCommands.GetCacheOptions(batch, pivotName);
            Assert.True(retained.Success, retained.ErrorMessage);
            Assert.True(retained.IsOlap);
            Assert.Equal(current.SaveSourceData, retained.SaveSourceData);
            Assert.Equal(current.OptimizeCache, retained.OptimizeCache);
            Assert.Equal(current.MissingItemsLimit, retained.MissingItemsLimit);
            Assert.Equal(current.RefreshOnFileOpen, retained.RefreshOnFileOpen);
            Assert.Equal(current.EnableRefresh, retained.EnableRefresh);
        }
    }

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
        var batch = _fixture.BatchToken;
        var (sheetName, pivotName) = CreateIsolatedPivot();

        var placed = _pivotCommands.AddColumnField(batch, pivotName, "[RegionalSalesTable].[Quarter]", null);
        Assert.True(placed.Success, placed.ErrorMessage);
        AssertCubeFieldOrientation(sheetName, pivotName, "[RegionalSalesTable].[Quarter]", PivotFieldArea.Column);
        var removed = _pivotCommands.RemoveField(batch, pivotName, "[RegionalSalesTable].[Quarter]");
        Assert.True(removed.Success, removed.ErrorMessage);
        AssertCubeFieldOrientation(sheetName, pivotName, "[RegionalSalesTable].[Quarter]", PivotFieldArea.Hidden);
        var result = _pivotCommands.AddRowField(batch, pivotName, "[RegionalSalesTable].[Quarter]", null);

        // Assert
        Assert.True(result.Success, $"Failed: {result.ErrorMessage}");
        Assert.Equal("[RegionalSalesTable].[Quarter]", result.FieldName);
        Assert.Equal(PivotFieldArea.Row, result.Area);
        Assert.True(_pivotCommands.Refresh(batch, pivotName).Success);
        AssertCubeFieldOrientation(sheetName, pivotName, "[RegionalSalesTable].[Quarter]", PivotFieldArea.Row);
    }

    [Fact]
    [Trait("Speed", "Medium")]
    [Trait("Category", "OLAP")]
    public void AddColumnField_OlapPivot_AddsFieldToColumns()
    {
        // Arrange - Create OLAP test file with data model
        var batch = _fixture.BatchToken;
        var (sheetName, pivotName) = CreateIsolatedPivot();

        var placed = _pivotCommands.AddRowField(batch, pivotName, "[RegionalSalesTable].[Quarter]", null);
        Assert.True(placed.Success, placed.ErrorMessage);
        AssertCubeFieldOrientation(sheetName, pivotName, "[RegionalSalesTable].[Quarter]", PivotFieldArea.Row);
        var removed = _pivotCommands.RemoveField(batch, pivotName, "[RegionalSalesTable].[Quarter]");
        Assert.True(removed.Success, removed.ErrorMessage);
        AssertCubeFieldOrientation(sheetName, pivotName, "[RegionalSalesTable].[Quarter]", PivotFieldArea.Hidden);
        var result = _pivotCommands.AddColumnField(batch, pivotName, "[RegionalSalesTable].[Quarter]", null);

        // Assert
        Assert.True(result.Success, $"Failed: {result.ErrorMessage}");
        Assert.Equal("[RegionalSalesTable].[Quarter]", result.FieldName);
        Assert.Equal(PivotFieldArea.Column, result.Area);
        Assert.True(_pivotCommands.Refresh(batch, pivotName).Success);
        AssertCubeFieldOrientation(sheetName, pivotName, "[RegionalSalesTable].[Quarter]", PivotFieldArea.Column);
    }

    [Fact]
    [Trait("Speed", "Medium")]
    [Trait("Category", "OLAP")]
    public void SortField_OlapPivot_SortsFieldSuccessfully()
    {
        var batch = _fixture.BatchToken;
        var sheetName = _fixture.CreateTestSheet(batch);
        const string pivotName = "SortVerificationPivot";
        var created = _pivotCommands.CreateFromDataModel(
            batch, "RegionalSalesTable", sheetName, "A1", pivotName);
        Assert.True(created.Success, created.ErrorMessage);
        var row = _pivotCommands.AddRowField(
            batch, pivotName, "[RegionalSalesTable].[Region]", null);
        Assert.True(row.Success, row.ErrorMessage);
        var value = _pivotCommands.AddValueField(
            batch, pivotName, "[Measures].[TotalRevenue]", AggregationFunction.Sum, null);
        Assert.True(value.Success, value.ErrorMessage);
        Assert.True(_pivotCommands.Refresh(batch, pivotName).Success);
        var original = _pivotCommands.GetData(batch, pivotName);
        Assert.True(original.Success, original.ErrorMessage);
        Assert.Equal(["East", "North", "South", "West"],
            original.Values.Select(row => row[0]?.ToString())
                .Where(value => value is "West" or "South" or "North" or "East"));

        var result = _pivotCommands.SortField(
            batch,
            pivotName,
            "[RegionalSalesTable].[Region]",
            SortDirection.Descending);

        // Assert
        Assert.True(result.Success, $"Failed: {result.ErrorMessage}");
        Assert.Equal("[RegionalSalesTable].[Region]", result.FieldName);
        var sorted = _pivotCommands.GetData(batch, pivotName);
        Assert.True(sorted.Success, sorted.ErrorMessage);
        Assert.Equal(["West", "South", "North", "East"],
            sorted.Values.Select(row => row[0]?.ToString())
                .Where(value => value is "West" or "South" or "North" or "East"));
        Assert.True(_pivotCommands.Refresh(batch, pivotName).Success);
        var data = _pivotCommands.GetData(batch, pivotName);
        Assert.True(data.Success, data.ErrorMessage);
        Assert.Equal(["West", "South", "North", "East"],
            data.Values.Select(row => row[0]?.ToString())
                .Where(value => value is "West" or "South" or "North" or "East"));
        var ascending = _pivotCommands.SortField(
            batch, pivotName, "[RegionalSalesTable].[Region]", SortDirection.Ascending);
        Assert.True(ascending.Success, ascending.ErrorMessage);
        var restored = _pivotCommands.GetData(batch, pivotName);
        Assert.True(restored.Success, restored.ErrorMessage);
        Assert.Equal(["East", "North", "South", "West"],
            restored.Values.Select(row => row[0]?.ToString())
                .Where(value => value is "West" or "South" or "North" or "East"));
        var rejected = _pivotCommands.SortField(
            batch, pivotName, "[RegionalSalesTable].[MissingRegion]", SortDirection.Descending);
        Assert.False(rejected.Success);
        Assert.Contains("not found in OLAP PivotTable", rejected.ErrorMessage);
        var retained = _pivotCommands.GetData(batch, pivotName);
        Assert.True(retained.Success, retained.ErrorMessage);
        Assert.Equal(restored.Values.Count, retained.Values.Count);
        for (var index = 0; index < restored.Values.Count; index++)
        {
            Assert.Equal(restored.Values[index], retained.Values[index]);
        }
    }

    [Fact]
    [Trait("Speed", "Medium")]
    [Trait("Category", "OLAP")]
    public void ReportFilter_DataModel_ExcelSupportsSingleMultipleAndClearAll()
    {
        var (sheetName, pivotName) = CreateIsolatedPivot();
        const string fieldName = "[RegionalSalesTable].[Region]";
        var batch = _fixture.BatchToken;
        var row = _pivotCommands.AddRowField(batch, pivotName, fieldName);
        Assert.True(row.Success, row.ErrorMessage);
        var value = _pivotCommands.AddValueField(
            batch, pivotName, "[Measures].[TotalRevenue]", AggregationFunction.Sum, null);
        Assert.True(value.Success, value.ErrorMessage);
        var refreshed = _pivotCommands.Refresh(batch, pivotName);
        Assert.True(refreshed.Success, refreshed.ErrorMessage);
        var untouchedPivotName = $"Untouched_{Guid.NewGuid():N}";
        var untouchedCreated = _pivotCommands.CreateFromDataModel(
            batch, "RegionalSalesTable", sheetName, "G1", untouchedPivotName);
        Assert.True(untouchedCreated.Success, untouchedCreated.ErrorMessage);
        value = _pivotCommands.AddValueField(
            batch, untouchedPivotName, "[Measures].[TotalRevenue]", AggregationFunction.Sum, null);
        Assert.True(value.Success, value.ErrorMessage);
        refreshed = _pivotCommands.Refresh(batch, untouchedPivotName);
        Assert.True(refreshed.Success, refreshed.ErrorMessage);

        var memberNames = _fixture.ExecuteRawVerification((context, _) =>
        {
            Excel.Sheets? sheets = null;
            Excel.Worksheet? sheet = null;
            Excel.PivotTables? pivots = null;
            Excel.PivotTable? pivot = null;
            Excel.CubeFields? cubeFields = null;
            Excel.CubeField? cubeField = null;
            Excel.PivotFields? fields = null;
            Excel.PivotField? field = null;
            Excel.PivotItems? items = null;
            Excel.PivotItem? item = null;
            try
            {
                sheets = context.Book.Worksheets;
                sheet = (Excel.Worksheet)sheets[sheetName];
                pivots = (Excel.PivotTables)sheet.PivotTables();
                pivot = pivots.Item(pivotName);
                cubeFields = pivot.CubeFields;
                cubeField = cubeFields[fieldName];
                fields = cubeField.PivotFields;
                field = fields.Item(1);
                items = field.PivotItems();
                var names = new List<string>();
                for (int index = 1; index <= items.Count; index++)
                {
                    item = items.Item(index);
                    names.Add(item.Name);
                    ComUtilities.Release(ref item);
                }
                return names;
            }
            finally
            {
                ComUtilities.Release(ref item);
                ComUtilities.Release(ref items);
                ComUtilities.Release(ref field);
                ComUtilities.Release(ref fields);
                ComUtilities.Release(ref cubeField);
                ComUtilities.Release(ref cubeFields);
                ComUtilities.Release(ref pivot);
                ComUtilities.Release(ref pivots);
                ComUtilities.Release(ref sheet);
                ComUtilities.Release(ref sheets);
            }
        });
        Assert.True(memberNames.Count >= 2,
            $"Expected at least two region members; received {memberNames.Count}.");
        var removed = _pivotCommands.RemoveField(batch, pivotName, fieldName);
        Assert.True(removed.Success, removed.ErrorMessage);
        var placed = _pivotCommands.AddFilterField(batch, pivotName, fieldName);
        Assert.True(placed.Success, placed.ErrorMessage);
        refreshed = _pivotCommands.Refresh(batch, pivotName);
        Assert.True(refreshed.Success, refreshed.ErrorMessage);

        _fixture.ExecuteRawVerification((context, _) =>
        {
            Excel.Sheets? sheets = null;
            Excel.Worksheet? sheet = null;
            Excel.PivotTables? pivots = null;
            Excel.PivotTable? pivot = null;
            Excel.CubeFields? cubeFields = null;
            Excel.CubeField? cubeField = null;
            Excel.PivotFields? fields = null;
            Excel.PivotField? field = null;
            try
            {
                sheets = context.Book.Worksheets;
                sheet = (Excel.Worksheet)sheets[sheetName];
                pivots = (Excel.PivotTables)sheet.PivotTables();
                pivot = pivots.Item(pivotName);
                cubeFields = pivot.CubeFields;
                cubeField = cubeFields[fieldName];
                fields = cubeField.PivotFields;
                field = fields.Item(1);
                Assert.Equal((int)PivotFieldArea.Filter, (int)cubeField.Orientation);
            }
            finally
            {
                ComUtilities.Release(ref field);
                ComUtilities.Release(ref fields);
                ComUtilities.Release(ref cubeField);
                ComUtilities.Release(ref cubeFields);
                ComUtilities.Release(ref pivot);
                ComUtilities.Release(ref pivots);
                ComUtilities.Release(ref sheet);
                ComUtilities.Release(ref sheets);
            }
        });

        var selected = _pivotCommands.SetReportFilter(
            batch, sheetName, pivotName, fieldName, [memberNames[0], memberNames[1]]);
        Assert.True(selected.Success, selected.ErrorMessage);
        Assert.False(selected.ShowAll);
        Assert.Equal([memberNames[0], memberNames[1]], selected.SelectedItems);
        Assert.False(selected.MayHavePartiallyChanged);

        var multiSelected = _pivotCommands.GetData(batch, pivotName);
        Assert.True(multiSelected.Success, multiSelected.ErrorMessage);
        Assert.Equal(22000, ReadScalar(multiSelected));
        var untouched = _pivotCommands.GetData(batch, untouchedPivotName);
        Assert.True(untouched.Success, untouched.ErrorMessage);
        Assert.Equal(49000, ReadScalar(untouched));

        selected = _pivotCommands.SetReportFilter(
            batch, sheetName, pivotName, fieldName, [memberNames[0], memberNames[1]]);
        Assert.True(selected.Success, selected.ErrorMessage);
        Assert.Equal([memberNames[0], memberNames[1]], selected.SelectedItems);

        selected = _pivotCommands.SetReportFilter(batch, sheetName, pivotName, fieldName, [memberNames[0]]);
        Assert.True(selected.Success, selected.ErrorMessage);
        Assert.False(selected.ShowAll);
        Assert.Equal([memberNames[0]], selected.SelectedItems);
        var singleSelected = _pivotCommands.GetData(batch, pivotName);
        Assert.True(singleSelected.Success, singleSelected.ErrorMessage);
        Assert.Equal(11500, ReadScalar(singleSelected));
        untouched = _pivotCommands.GetData(batch, untouchedPivotName);
        Assert.True(untouched.Success, untouched.ErrorMessage);
        Assert.Equal(49000, ReadScalar(untouched));

        selected = _pivotCommands.SetReportFilter(
            batch, sheetName, pivotName, fieldName, [memberNames[0], memberNames[1]]);
        Assert.True(selected.Success, selected.ErrorMessage);
        Assert.Equal([memberNames[0], memberNames[1]], selected.SelectedItems);

        selected = _pivotCommands.SetReportFilter(batch, sheetName, pivotName, fieldName, []);
        Assert.True(selected.Success, selected.ErrorMessage);
        Assert.True(selected.ShowAll);
        Assert.Empty(selected.SelectedItems);
        Assert.False(selected.MayHavePartiallyChanged);
        var allSelected = _pivotCommands.GetData(batch, pivotName);
        Assert.True(allSelected.Success, allSelected.ErrorMessage);
        Assert.Equal(49000, ReadScalar(allSelected));
        untouched = _pivotCommands.GetData(batch, untouchedPivotName);
        Assert.True(untouched.Success, untouched.ErrorMessage);
        Assert.Equal(49000, ReadScalar(untouched));

        var rejected = _pivotCommands.SetReportFilter(
            batch, sheetName, pivotName, fieldName, ["[RegionalSalesTable].[Region].&[Missing]"]);
        Assert.False(rejected.Success);
        Assert.True(rejected.MayHavePartiallyChanged);
        Assert.False(rejected.RollbackAttempted);
        Assert.False(string.IsNullOrWhiteSpace(rejected.ErrorMessage));
        allSelected = _pivotCommands.GetData(batch, pivotName);
        Assert.True(allSelected.Success, allSelected.ErrorMessage);
        Assert.Equal(49000, ReadScalar(allSelected));

        rejected = _pivotCommands.SetReportFilter(batch, sheetName, pivotName, fieldName, [" "]);
        Assert.False(rejected.Success);
        Assert.False(rejected.MayHavePartiallyChanged);
        rejected = _pivotCommands.SetReportFilter(
            batch, sheetName, pivotName, fieldName, [memberNames[0], memberNames[0]]);
        Assert.False(rejected.Success);
        Assert.False(rejected.MayHavePartiallyChanged);

        rejected = _pivotCommands.SetReportFilter(
            batch, $"{sheetName}_missing", pivotName, fieldName, [memberNames[0]]);
        Assert.False(rejected.Success);
        Assert.False(rejected.MayHavePartiallyChanged);

        removed = _pivotCommands.RemoveField(batch, pivotName, fieldName);
        Assert.True(removed.Success, removed.ErrorMessage);
        row = _pivotCommands.AddRowField(batch, pivotName, fieldName);
        Assert.True(row.Success, row.ErrorMessage);
        rejected = _pivotCommands.SetReportFilter(
            batch, sheetName, pivotName, fieldName, [memberNames[0]]);
        Assert.False(rejected.Success);
        Assert.False(rejected.MayHavePartiallyChanged);
        Assert.Contains("report-filter area", rejected.ErrorMessage);
        AssertCubeFieldOrientation(sheetName, pivotName, fieldName, PivotFieldArea.Row);

        static double ReadScalar(PivotTableDataResult result)
        {
            var values = result.Values
                .SelectMany(row => row)
                .Where(value => value is not null &&
                    double.TryParse(value.ToString(), System.Globalization.NumberStyles.Float,
                        System.Globalization.CultureInfo.InvariantCulture, out _))
                .Select(value => Convert.ToDouble(value, System.Globalization.CultureInfo.InvariantCulture));
            return Assert.Single(values);
        }
    }

    /// <summary>
    /// Regression test for Issue #217: Auto-create DAX measures when adding value fields to OLAP PivotTables.
    ///
    /// AddValueField creates a DAX measure and places it in Values.
    /// </summary>
    [Fact]
    [Trait("Speed", "Medium")]
    [Trait("Category", "OLAP")]
    public void AddValueField_OlapPivot_AutoCreatesDaxMeasure()
    {
        // Arrange - Create OLAP test file with Data Model PivotTable
        var batch = _fixture.BatchToken;
        var (sheetName, pivotName) = CreateIsolatedPivot();

        // Act - Try to add Sales field as a Sum value field
        // Use exact CubeField name format [TableName].[ColumnName]
        // After implementation, should auto-create: [Regional Sales Total] = SUM('RegionalSalesTable'[Sales])
        // NOTE: Use unique name to avoid conflict with fixture's "Total Sales" measure on SalesTable
        var result = _pivotCommands.AddValueField(
            batch,
            pivotName,
            "[RegionalSalesTable].[Sales]",
            AggregationFunction.Sum,
            "Regional Sales Total");

        // Assert - Should succeed with auto-created DAX measure
        Assert.True(result.Success, $"AddValueField should auto-create DAX measure but failed: {result.ErrorMessage}");
        Assert.Equal("Regional Sales Total", result.FieldName); // Field name is the measure name
        Assert.Equal(PivotFieldArea.Value, result.Area);
        Assert.Equal("Regional Sales Total", result.CustomName);

        // Verify the DAX measure was created in Data Model
        var dataModelCommands = _dataModelCommands;
        var measuresResult = dataModelCommands.ListMeasures(batch, "RegionalSalesTable");
        Assert.True(measuresResult.Success, $"Failed to list measures: {measuresResult.ErrorMessage}");

        var measure = Assert.Single(measuresResult.Measures, m => m.Name == "Regional Sales Total");
        Assert.Equal("SUM('RegionalSalesTable'[Sales])", measure.FormulaPreview);
        AssertMeasureValue("Regional Sales Total", 49000);
        Assert.True(_pivotCommands.Refresh(batch, pivotName).Success);
        AssertCubeFieldOrientation(sheetName, pivotName, "[Measures].[Regional Sales Total]", PivotFieldArea.Value);
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
        var batch = _fixture.BatchToken;
        var (sheetName, pivotName) = CreateIsolatedPivot();

        // Act - Add Quarter field with Count aggregation
        // Use exact CubeField name format [TableName].[ColumnName]
        // Should auto-create: [Number of Quarters] = COUNT('RegionalSalesTable'[Quarter])
        var result = _pivotCommands.AddValueField(
            batch,
            pivotName,
            "[RegionalSalesTable].[Quarter]",
            AggregationFunction.Count,
            "Number of Quarters");

        // Assert
        Assert.True(result.Success, $"AddValueField with Count should auto-create DAX measure but failed: {result.ErrorMessage}");
        Assert.Equal("Number of Quarters", result.FieldName); // Field name is the measure name
        Assert.Equal(PivotFieldArea.Value, result.Area);
        Assert.Equal(AggregationFunction.Count, result.Function);

        // Verify the DAX measure was created with COUNT function
        var dataModelCommands = _dataModelCommands;
        var measuresResult = dataModelCommands.ListMeasures(batch, "RegionalSalesTable");
        Assert.True(measuresResult.Success, $"Failed to list measures: {measuresResult.ErrorMessage}");
        var measure = Assert.Single(measuresResult.Measures, m => m.Name == "Number of Quarters");
        Assert.Equal("COUNT('RegionalSalesTable'[Quarter])", measure.FormulaPreview);
        AssertMeasureValue("Number of Quarters", 8);
        Assert.True(_pivotCommands.Refresh(batch, pivotName).Success);
        AssertCubeFieldOrientation(sheetName, pivotName, "[Measures].[Number of Quarters]", PivotFieldArea.Value);
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
        var batch = _fixture.BatchToken;
        var (sheetName, pivotName) = CreateIsolatedPivot();

        // First, create a measure in the Data Model (not in PivotTable yet)
        var dataModelCommands = _dataModelCommands;
        dataModelCommands.CreateMeasure(
            batch,
            "RegionalSalesTable",
            "Total ACR",
            "SUM('RegionalSalesTable'[Sales])",
            null);  // CreateMeasure throws on error

        // Refresh PivotTable to pick up the new measure in CubeFields
        Assert.True(_pivotCommands.Refresh(batch, pivotName, null).Success);
        AssertCubeFieldOrientation(sheetName, pivotName, "[Measures].[Total ACR]", PivotFieldArea.Hidden);

        // Act - Add the pre-existing measure to PivotTable values area
        // Should detect it's an existing measure and just set Orientation = xlDataField
        var result = _pivotCommands.AddValueField(
            batch,
            pivotName,
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
        AssertMeasureValue("Total ACR", 49000);
        Assert.True(_pivotCommands.Refresh(batch, pivotName).Success);
        AssertCubeFieldOrientation(sheetName, pivotName, "[Measures].[Total ACR]", PivotFieldArea.Value);
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
        var batch = _fixture.BatchToken;
        var (sheetName, pivotName) = CreateIsolatedPivot();

        var dataModelCommands = _dataModelCommands;
        dataModelCommands.CreateMeasure(
            batch,
            "RegionalSalesTable",
            "Revenue Total",
            "SUM('RegionalSalesTable'[Sales])",
            null);  // CreateMeasure throws on error

        Assert.True(_pivotCommands.Refresh(batch, pivotName, null).Success);
        AssertCubeFieldOrientation(sheetName, pivotName, "[Measures].[Revenue Total]", PivotFieldArea.Hidden);

        // Act - Use [Measures].[Name] format (common in OLAP contexts)
        var result = _pivotCommands.AddValueField(
            batch,
            pivotName,
            "[Measures].[Revenue Total]", // MDX-style format
            AggregationFunction.Sum,
            null);

        // Assert
        Assert.True(result.Success, $"Should handle [Measures].[Name] format but failed: {result.ErrorMessage}");
        Assert.Equal("Revenue Total", result.FieldName);
        Assert.Equal(PivotFieldArea.Value, result.Area);
        AssertMeasureValue("Revenue Total", 49000);
        Assert.True(_pivotCommands.Refresh(batch, pivotName).Success);
        AssertCubeFieldOrientation(sheetName, pivotName, "[Measures].[Revenue Total]", PivotFieldArea.Value);
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
        var batch = _fixture.BatchToken;
        var (sheetName, pivotName) = CreateIsolatedPivot();

        // Use exact CubeField name format [TableName].[ColumnName]
        var addResult = _pivotCommands.AddValueField(
            batch,
            pivotName,
            "[RegionalSalesTable].[Sales]",
            AggregationFunction.Sum,
            "Sales Measure");
        Assert.True(addResult.Success, $"Setup failed: {addResult.ErrorMessage}");
        AssertMeasureValue("Sales Measure", 49000);

        // Act - Change from SUM to COUNT
        // After the measure is created, reference it by its measure name or [Measures].[Name]
        var updateResult = _pivotCommands.SetFieldFunction(
            batch,
            pivotName,
            "[Measures].[Sales Measure]",
            AggregationFunction.Count);

        // Assert - Operation succeeded
        Assert.True(updateResult.Success, $"Update failed: {updateResult.ErrorMessage}");
        Assert.Contains("Sales Measure", updateResult.FieldName);
        Assert.Equal(AggregationFunction.Count, updateResult.Function);

        // Verify the DAX measure formula changed in Data Model
        var dataModelCommands = _dataModelCommands;
        var measuresResult = dataModelCommands.ListMeasures(batch, "RegionalSalesTable");
        Assert.True(measuresResult.Success, $"Failed to list measures: {measuresResult.ErrorMessage}");

        var measure = Assert.Single(measuresResult.Measures, m => m.Name == "Sales Measure");
        Assert.Equal("COUNT('RegionalSalesTable'[Sales])", measure.FormulaPreview);
        AssertMeasureValue("Sales Measure", 8);
        Assert.True(_pivotCommands.Refresh(batch, pivotName).Success);
        AssertCubeFieldOrientation(sheetName, pivotName, "[Measures].[Sales Measure]", PivotFieldArea.Value);
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
        var batch = _fixture.BatchToken;
        var (sheetName, pivotName) = CreateIsolatedPivot();

        // Use exact CubeField name format [TableName].[ColumnName]
        var addResult = _pivotCommands.AddValueField(
            batch,
            pivotName,
            "[RegionalSalesTable].[Sales]",
            AggregationFunction.Sum,
            "Sales Total");
        Assert.True(addResult.Success, $"Setup failed: {addResult.ErrorMessage}");

        // Act - Set a simple format that Excel preserves exactly
        // After the measure is created, reference it by [Measures].[Name]
        // Use "0%" which is locale-independent
        var updateResult = _pivotCommands.SetFieldFormat(
            batch,
            pivotName,
            "[Measures].[Sales Total]",
            "0%");

        // Assert - Operation succeeded
        Assert.True(updateResult.Success, $"Update failed: {updateResult.ErrorMessage}");
        Assert.Contains("Sales Total", updateResult.FieldName);
        Assert.Equal("0%", updateResult.NumberFormat);
        AssertPivotFieldFormat(sheetName, pivotName, "Sales Total", "0%");
        AssertMeasureValue("Sales Total", 49000);
    }

    /// <summary>
    /// Test UPDATE: Format a PRE-EXISTING measure (not created in same test).
    /// This covers the bug scenario where SetFieldFormat failed for measures
    /// created via datamodel tool, which exist in CubeFields but not
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
        var batch = _fixture.BatchToken;
        var (sheetName, pivotName) = CreateIsolatedPivot("DisambiguationTable");

        // First, add the pre-existing ACR measure to the DisambiguationTest PivotTable
        // The measure exists in Data Model but needs to be added to PivotTable's Values area
        var addResult = _pivotCommands.AddValueField(
            batch,
            pivotName,
            "[Measures].[ACR]",  // Pre-existing measure from fixture
            AggregationFunction.Sum,  // Ignored for existing measures
            null);  // Keep existing name
        Assert.True(addResult.Success, $"AddValueField failed: {addResult.ErrorMessage}");

        // Act - Format the pre-existing measure (this was the bug scenario)
        // Use "0%" format which is locale-independent
        var formatResult = _pivotCommands.SetFieldFormat(
            batch,
            pivotName,
            "[Measures].[ACR]",
            "0%");

        // Assert - Operation succeeded (was failing with "Measure not found in Data Model")
        Assert.True(formatResult.Success, $"SetFieldFormat failed: {formatResult.ErrorMessage}");
        Assert.Contains("ACR", formatResult.FieldName);
        Assert.Equal("0%", formatResult.NumberFormat);
        AssertPivotFieldFormat(sheetName, pivotName, "ACR", "0%");
    }

    private (string SheetName, string PivotName) CreateIsolatedPivot(string tableName = "RegionalSalesTable")
    {
        var sheetName = _fixture.CreateTestSheet(_fixture.BatchToken);
        var pivotName = $"FieldVerification_{Guid.NewGuid():N}";
        var created = _pivotCommands.CreateFromDataModel(
            _fixture.BatchToken, tableName, sheetName, "A1", pivotName);
        Assert.True(created.Success, created.ErrorMessage);
        return (sheetName, pivotName);
    }

    private void AssertPivotFieldFormat(string sheetName, string pivotName, string measureName, string expected) =>
        _fixture.ExecuteRawVerification((context, _) =>
        {
            Excel.Sheets? sheets = null;
            Excel.Worksheet? sheet = null;
            Excel.PivotTables? pivots = null;
            Excel.PivotTable? pivot = null;
            Excel.CubeFields? cubes = null;
            Excel.CubeField? cube = null;
            Excel.PivotFields? fields = null;
            Excel.PivotField? field = null;
            try
            {
                sheets = context.Book.Worksheets;
                sheet = (Excel.Worksheet)sheets[sheetName];
                pivots = (Excel.PivotTables)sheet.PivotTables();
                pivot = pivots.Item(pivotName);
                cubes = pivot.CubeFields;
                cube = cubes[$"[Measures].[{measureName}]"];
                fields = cube.PivotFields;
                field = fields.Item(1);
                Assert.Equal(expected, field.NumberFormat);
            }
            finally
            {
                ComUtilities.Release(ref field);
                ComUtilities.Release(ref fields);
                ComUtilities.Release(ref cube);
                ComUtilities.Release(ref cubes);
                ComUtilities.Release(ref pivot);
                ComUtilities.Release(ref pivots);
                ComUtilities.Release(ref sheet);
                ComUtilities.Release(ref sheets);
            }
        });

    private void AssertMeasureValue(string measureName, double expected)
    {
        var result = _dataModelCommands.Evaluate(
            _fixture.BatchToken, $"EVALUATE ROW(\"Value\", [{measureName}])");
        Assert.True(result.Success, result.ErrorMessage);
        Assert.Equal("[Value]", Assert.Single(result.Columns));
        Assert.Equal(expected, Convert.ToDouble(
            Assert.Single(Assert.Single(result.Rows)), System.Globalization.CultureInfo.InvariantCulture));
    }

    private void AssertCubeFieldOrientation(
        string sheetName, string pivotName, string fieldName, PivotFieldArea expected) =>
        _fixture.ExecuteRawVerification((context, _) =>
        {
            Excel.Sheets? sheets = null;
            Excel.Worksheet? sheet = null;
            Excel.PivotTables? pivots = null;
            Excel.PivotTable? pivot = null;
            Excel.CubeFields? cubes = null;
            Excel.CubeField? cube = null;
            try
            {
                sheets = context.Book.Worksheets;
                sheet = (Excel.Worksheet)sheets[sheetName];
                pivots = (Excel.PivotTables)sheet.PivotTables();
                pivot = pivots.Item(pivotName);
                cubes = pivot.CubeFields;
                cube = cubes[fieldName];
                Assert.Equal((int)expected, (int)cube.Orientation);
            }
            finally
            {
                ComUtilities.Release(ref cube);
                ComUtilities.Release(ref cubes);
                ComUtilities.Release(ref pivot);
                ComUtilities.Release(ref pivots);
                ComUtilities.Release(ref sheet);
                ComUtilities.Release(ref sheets);
            }
        });
}
