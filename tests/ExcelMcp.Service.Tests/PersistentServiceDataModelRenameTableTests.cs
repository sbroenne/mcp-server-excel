// Copyright (c) Sbroenne. All rights reserved.
// Licensed under the MIT License.

using Sbroenne.ExcelMcp.Core.Commands;
using Sbroenne.ExcelMcp.Core.Models;
using System.Globalization;
using System.Text.Json;
using Sbroenne.ExcelMcp.ComInterop;
using Xunit;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sbroenne.ExcelMcp.Service.Tests;

/// <summary>
/// Integration tests for RenameTable operation in the Data Model.
///
/// KNOWN EXCEL LIMITATION: Data Model table names (ModelTable.Name) are IMMUTABLE after creation.
/// The table name is cached from the source connection at creation time and CANNOT be changed
/// via the COM API - not through direct property assignment, Model.Refresh(), or even save/reopen.
///
/// These tests verify:
/// 1. The implementation correctly attempts the rename via COM
/// 2. The implementation returns a clear failure when the rename cannot be performed
/// 3. Rollback preserves the original state (PQ + connection names are restored)
/// 4. Validation rules (empty names, conflicts, non-existent tables) work correctly
/// </summary>
[Trait("Layer", "Service")]
[Trait("Category", "Integration")]
[Collection("ServiceWorkflow")]
[Trait("RequiresExcel", "true")]
[Trait("Feature", "DataModel")]
[Trait("Speed", "Slow")]
public class PersistentServiceDataModelRenameTableTests(
    PersistentServiceWorkbookFixture fixture) :
    PersistentServiceWorkbookTestBase(fixture),
    IClassFixture<PersistentServiceWorkbookFixture>
{
    private readonly IDataModelCommands _dataModelCommands =
        fixture.CreateCommands<IDataModelCommands>();
    private readonly IPowerQueryCommands _powerQueryCommands =
        fixture.CreateCommands<IPowerQueryCommands>();
    private readonly IDataModelRelCommands _relationships =
        fixture.CreateCommands<IDataModelRelCommands>();
    private readonly List<QueryState> _states = [];
    private string _baseline = "";
    private string _guardSheet = "";
    private string _measure = "";
    private int _expectedTotal;

    private sealed record QueryState(string Name, string Source, string[] Columns, object[][] Rows);

    /// <summary>
    /// Creates a test file with a PQ-backed Data Model table that can be renamed.
    /// PQ-backed tables are created by loading a Power Query with LoadToDataModel mode,
    /// which creates a "Query - {QueryName}" connection with Microsoft.Mashup.OleDb provider.
    /// </summary>
    private void CreateDataModelTable(string tableName = "TestTable")
    {
        string mCode = $@"let
    Source = #table(
        type table [ID = Int64.Type, Value = Int64.Type, Category = text],
        {{{{1, 100, ""A""}}, {{2, 200, ""B""}}, {{3, 300, ""A""}}}}
    )
in
    Source";

        // Create PQ with LoadToDataModel - this creates the "Query - {tableName}" connection
        RequireSuccess(_powerQueryCommands.Create(
            _fixture.BatchToken,
            tableName,
            mCode,
            PowerQueryLoadMode.LoadToDataModel));
        _fixture.RegisterPowerQueryForCleanup(tableName);
        _states.Add(new(tableName, mCode, ["ID", "Value", "Category"],
            [[1, 100, "A"], [2, 200, "B"], [3, 300, "A"]]));
        SeedDependencies(tableName, 600);
    }

    /// <summary>
    /// Creates a test file with two PQ-backed Data Model tables for conflict testing.
    /// </summary>
    private void CreateTwoDataModelTables(
        string table1Name = "Table1",
        string table2Name = "Table2")
    {
        // Create first Power Query → Data Model
        string mCode1 = $@"let
    Source = #table(
        type table [ID = Int64.Type, Value = Int64.Type],
        {{{{1, 100}}, {{2, 200}}}}
    )
in
    Source";
        RequireSuccess(_powerQueryCommands.Create(
            _fixture.BatchToken,
            table1Name,
            mCode1,
            PowerQueryLoadMode.LoadToDataModel));
        _fixture.RegisterPowerQueryForCleanup(table1Name);
        _states.Add(new(table1Name, mCode1, ["ID", "Value"], [[1, 100], [2, 200]]));

        // Create second Power Query → Data Model
        string mCode2 = $@"let
    Source = #table(
        type table [Category = text, Name = text],
        {{{{""A"", ""Alpha""}}, {{""B"", ""Beta""}}}}
    )
in
    Source";
        RequireSuccess(_powerQueryCommands.Create(
            _fixture.BatchToken,
            table2Name,
            mCode2,
            PowerQueryLoadMode.LoadToDataModel));
        _fixture.RegisterPowerQueryForCleanup(table2Name);
        _states.Add(new(table2Name, mCode2, ["Category", "Name"], [["A", "Alpha"], ["B", "Beta"]]));
        SeedDependencies(table1Name, 300);
    }

    // ==========================================
    // EXCEL LIMITATION CASES
    // These tests verify that the implementation correctly handles
    // the Excel limitation where ModelTable.Name is immutable.
    // ==========================================

    /// <summary>
    /// Tests that attempting to rename a PQ-backed Data Model table fails
    /// with a clear error message about the Excel limitation.
    /// LLM use case: "rename data model table from 'SalesData' to 'SalesTable'"
    /// </summary>
    [Fact]
    public void RenameTable_PqBackedTable_FailsDueToExcelLimitation()
    {
        // Arrange
        CreateDataModelTable("OriginalTable");

        var batch = _fixture.BatchToken;

        // Verify table exists in Data Model
        var listBefore = RequireSuccess(_dataModelCommands.ListTables(batch));
        Assert.True(listBefore.Success);
        Assert.Contains(listBefore.Tables, t => t.Name == "OriginalTable");

        // Act
        var result = _dataModelCommands.RenameTable(batch, "OriginalTable", "RenamedTable");

        // Assert - Rename fails due to Excel limitation
        Assert.False(result.Success);
        Assert.Contains("immutable", result.ErrorMessage, StringComparison.OrdinalIgnoreCase);
        Assert.Equal("data-model-table", result.ObjectType);
        Assert.Equal("OriginalTable", result.OldName);
        Assert.Equal("RenamedTable", result.NewName);

        // Verify original table is preserved (rollback worked)
        var listAfter = RequireSuccess(_dataModelCommands.ListTables(batch));
        Assert.True(listAfter.Success);
        Assert.Contains(listAfter.Tables, t => t.Name == "OriginalTable");
        Assert.DoesNotContain(listAfter.Tables, t => t.Name == "RenamedTable");
        AssertPreserved();
    }

    /// <summary>
    /// Tests that rename attempts with whitespace are normalized but still fail
    /// due to the Excel limitation on Data Model table names.
    /// </summary>
    [Fact]
    public void RenameTable_WithLeadingTrailingSpaces_FailsDueToExcelLimitation()
    {
        // Arrange
        CreateDataModelTable("TestTable");

        var batch = _fixture.BatchToken;

        // Act - rename with whitespace in new name
        var result = _dataModelCommands.RenameTable(batch, "  TestTable  ", "  TrimmedName  ");

        // Assert - Names are normalized but rename fails
        Assert.False(result.Success);
        Assert.Contains("immutable", result.ErrorMessage, StringComparison.OrdinalIgnoreCase);
        Assert.Equal("TestTable", result.NormalizedOldName);      // Normalized (trimmed)
        Assert.Equal("TrimmedName", result.NormalizedNewName);    // Normalized (trimmed)

        // Verify original table is preserved
        var list = RequireSuccess(_dataModelCommands.ListTables(batch));
        Assert.Contains(list.Tables, t => t.Name == "TestTable");
        AssertPreserved();
    }

    // ==========================================
    // NO-OP CASES
    // ==========================================

    /// <summary>
    /// Tests that renaming to the same name (after trim) is a no-op success.
    /// </summary>
    [Fact]
    public void RenameTable_SameNameAfterTrim_ReturnsNoOpSuccess()
    {
        // Arrange
        CreateDataModelTable("TestTable");

        var batch = _fixture.BatchToken;

        // Act - rename to same name (with extra spaces that get trimmed)
        var result = _dataModelCommands.RenameTable(batch, "TestTable", "  TestTable  ");

        // Assert - should be success (no-op)
        Assert.True(result.Success, $"No-op should succeed: {result.ErrorMessage}");
        Assert.True(string.IsNullOrEmpty(result.ErrorMessage));
        Assert.Equal("TestTable", result.NormalizedOldName);
        Assert.Equal("TestTable", result.NormalizedNewName);  // Same after normalization
        AssertPreserved();
    }

    // ==========================================
    // CASE-ONLY RENAME CASES
    // ==========================================

    /// <summary>
    /// Tests that case-only rename also fails due to the Excel limitation.
    /// Even though case-only changes are technically "the same" table, Excel
    /// still cannot change the ModelTable.Name property.
    /// </summary>
    [Fact]
    public void RenameTable_CaseOnlyChange_FailsDueToExcelLimitation()
    {
        // Arrange
        CreateDataModelTable("testtable");

        var batch = _fixture.BatchToken;

        // Act - rename to same name with different casing
        var result = _dataModelCommands.RenameTable(batch, "testtable", "TestTable");

        // Assert - Rename fails due to Excel limitation
        Assert.False(result.Success);
        Assert.Contains("immutable", result.ErrorMessage, StringComparison.OrdinalIgnoreCase);
        Assert.Equal("testtable", result.OldName);
        Assert.Equal("TestTable", result.NewName);

        // Verify original table is preserved
        var list = RequireSuccess(_dataModelCommands.ListTables(batch));
        Assert.Contains(list.Tables, t => t.Name.Equals("testtable", StringComparison.OrdinalIgnoreCase));
        AssertPreserved();
    }

    // ==========================================
    // CONFLICT CASES
    // ==========================================

    /// <summary>
    /// Tests that renaming to an existing table name (case-insensitive) fails.
    /// </summary>
    [Fact]
    public void RenameTable_ConflictWithExistingTable_ReturnsFailure()
    {
        // Arrange
        CreateTwoDataModelTables("SourceTable", "TargetTable");

        var batch = _fixture.BatchToken;

        // Act - try to rename to existing table name
        var result = _dataModelCommands.RenameTable(batch, "SourceTable", "TargetTable");

        // Assert
        Assert.False(result.Success);
        Assert.Contains("already exists", result.ErrorMessage, StringComparison.OrdinalIgnoreCase);
        AssertPreserved();
    }

    /// <summary>
    /// Tests that case-insensitive conflict detection works.
    /// </summary>
    [Fact]
    public void RenameTable_CaseInsensitiveConflict_ReturnsFailure()
    {
        // Arrange
        CreateTwoDataModelTables("SourceTable", "TARGETTABLE");

        var batch = _fixture.BatchToken;

        // Act - try to rename with different case of existing name
        var result = _dataModelCommands.RenameTable(batch, "SourceTable", "targettable");

        // Assert
        Assert.False(result.Success);
        Assert.Contains("already exists", result.ErrorMessage, StringComparison.OrdinalIgnoreCase);
        AssertPreserved();
    }

    // ==========================================
    // MISSING TABLE CASES
    // ==========================================

    /// <summary>
    /// Tests that renaming a non-existent table fails with clear error.
    /// </summary>
    [Fact]
    public void RenameTable_NonExistentTable_ReturnsFailure()
    {
        // Arrange
        CreateDataModelTable("ExistingTable");

        var batch = _fixture.BatchToken;

        // Act
        var result = _dataModelCommands.RenameTable(batch, "NonExistentTable", "NewName");

        // Assert
        Assert.False(result.Success);
        Assert.Contains("not found", result.ErrorMessage, StringComparison.OrdinalIgnoreCase);
        AssertPreserved();
    }

    // ==========================================
    // INVALID NAME CASES
    // ==========================================

    /// <summary>
    /// Tests that empty new name fails validation.
    /// </summary>
    [Fact]
    public void RenameTable_EmptyNewName_ReturnsFailure()
    {
        // Arrange
        CreateDataModelTable("TestTable");

        var batch = _fixture.BatchToken;

        // Act
        var exception = Assert.Throws<ArgumentException>(() =>
            _dataModelCommands.RenameTable(batch, "TestTable", ""));

        Assert.Contains("required", exception.Message, StringComparison.OrdinalIgnoreCase);
        AssertPreserved();
    }

    /// <summary>
    /// Tests that whitespace-only new name fails validation.
    /// </summary>
    [Fact]
    public void RenameTable_WhitespaceOnlyNewName_ReturnsFailure()
    {
        // Arrange
        CreateDataModelTable("TestTable");

        var batch = _fixture.BatchToken;

        // Act
        var exception = Assert.Throws<ArgumentException>(() =>
            _dataModelCommands.RenameTable(batch, "TestTable", "   "));

        Assert.Contains("required", exception.Message, StringComparison.OrdinalIgnoreCase);
        AssertPreserved();
    }

    /// <summary>
    /// Tests that empty old name fails validation.
    /// </summary>
    [Fact]
    public void RenameTable_EmptyOldName_ReturnsFailure()
    {
        // Arrange
        CreateDataModelTable("TestTable");

        var batch = _fixture.BatchToken;

        // Act
        var exception = Assert.Throws<ArgumentException>(() =>
            _dataModelCommands.RenameTable(batch, "", "NewName"));

        Assert.Contains("required", exception.Message, StringComparison.OrdinalIgnoreCase);
        AssertPreserved();
    }

    // ==========================================
    // ROUND-TRIP PERSISTENCE TEST
    // ==========================================

    /// <summary>
    /// Tests that when rename fails, the original table name is preserved across save/reopen.
    /// This verifies the rollback mechanism works correctly.
    /// </summary>
    [Fact]
    public async Task RenameTable_FailureThenSaveAndReopen_PreservesOriginalTable()
    {
        // Arrange
        CreateDataModelTable("OriginalName");

        // Act - Attempt rename (will fail), then save
        var result = _dataModelCommands.RenameTable(
            _fixture.BatchToken,
            "OriginalName",
            "PersistedName");
        Assert.False(result.Success);
        Assert.Contains(
            "immutable",
            result.ErrorMessage,
            StringComparison.OrdinalIgnoreCase);
        AssertPreserved();
        await _fixture.SaveAndReopenAsync();

        // Assert - Reopen and verify original table is preserved
        var batch2 = _fixture.BatchToken;
        var list = RequireSuccess(_dataModelCommands.ListTables(batch2));
        Assert.True(list.Success);
        Assert.Contains(list.Tables, t => t.Name == "OriginalName");  // Original preserved
        Assert.DoesNotContain(list.Tables, t => t.Name == "PersistedName");  // New name not present
        AssertPreserved();
    }

    private void SeedDependencies(string tableName, int expectedTotal)
    {
        const string source = """
            let Source = #table(type table [ID = Int64.Type, Label = text],
                {{1, "Lookup one"}, {2, "Lookup two"}, {3, "Lookup three"}})
            in Source
            """;
        var lookup = $"RenameLookup_{Guid.NewGuid():N}"[..26];
        RequireSuccess(_powerQueryCommands.Create(_fixture.BatchToken, lookup, source,
            PowerQueryLoadMode.LoadToDataModel));
        _fixture.RegisterPowerQueryForCleanup(lookup);
        _states.Add(new(lookup, source, ["ID", "Label"],
            [[1, "Lookup one"], [2, "Lookup two"], [3, "Lookup three"]]));
        RequireSuccess(_relationships.CreateRelationship(_fixture.BatchToken, tableName, "ID", lookup, "ID"));
        _measure = $"RenameTotal_{Guid.NewGuid():N}"[..26];
        RequireSuccess(_dataModelCommands.CreateMeasure(_fixture.BatchToken, tableName, _measure,
            $"SUM('{tableName}'[Value])", formatType: "WholeNumber"));
        _fixture.RegisterDataModelMeasureForCleanup(_measure);
        _expectedTotal = expectedTotal;
        _guardSheet = _fixture.CreateTestSheet(_fixture.BatchToken);
        RequireSuccess(_commands.SetValues(_fixture.BatchToken, _guardSheet, "A1:B2",
            [["Rename guard", 37], ["Retained cells", 91]]));
        _baseline = CaptureState();
    }

    private void AssertPreserved() => Assert.Equal(_baseline, CaptureState());

    private string CaptureState()
    {
        var batch = _fixture.BatchToken;
        foreach (var state in _states)
        {
            PowerQueryStateAssertions.AssertStored(_fixture, state.Name, state.Source,
                PowerQueryLoadMode.LoadToDataModel, null, state.Columns, state.Rows);
        }
        var queries = RequireSuccess(_powerQueryCommands.List(batch)).Queries;
        var tables = RequireSuccess(_dataModelCommands.ListTables(batch)).Tables;
        Assert.Equal(_states.Count, queries.Count);
        Assert.Equal(_states.Count, tables.Count);
        var relationships = RequireSuccess(_relationships.ListRelationships(batch)).Relationships;
        Assert.Single(relationships);
        var measures = RequireSuccess(_dataModelCommands.ListMeasures(batch)).Measures;
        Assert.Equal(_measure, Assert.Single(measures).Name);
        var total = RequireSuccess(_dataModelCommands.Evaluate(batch, $"EVALUATE ROW(\"Total\", [{_measure}])"));
        PowerQueryStateAssertions.AssertRows([[_expectedTotal]], total.Rows);
        var cells = RequireSuccess(_commands.GetValues(batch, _guardSheet, "A1:B2")).Values;
        PowerQueryStateAssertions.AssertRows([["Rename guard", 37], ["Retained cells", 91]], cells);
        var connections = _fixture.ExecuteRawVerification((context, _) =>
        {
            Excel.Connections? collection = null;
            try
            {
                collection = context.Book.Connections;
                var snapshots = new List<object>();
                for (var index = 1; index <= collection.Count; index++)
                {
                    Excel.WorkbookConnection? connection = null;
                    Excel.OLEDBConnection? oledb = null;
                    try
                    {
                        connection = collection.Item(index);
                        if (connection.Type == Excel.XlConnectionType.xlConnectionTypeOLEDB)
                        {
                            oledb = connection.OLEDBConnection;
                            Assert.False(oledb.Refreshing);
                            snapshots.Add(new
                            {
                                connection.Name,
                                connection.Description,
                                connection.Type,
                                connection.InModel,
                                Source = Convert.ToString(oledb.Connection, CultureInfo.InvariantCulture),
                                Command = JsonSerializer.Serialize(oledb.CommandText),
                                oledb.CommandType,
                                oledb.BackgroundQuery,
                                oledb.RefreshOnFileOpen,
                                oledb.RefreshPeriod
                            });
                        }
                        else
                        {
                            snapshots.Add(new { connection.Name, connection.Description, connection.Type });
                        }
                    }
                    finally
                    {
                        ComUtilities.Release(ref oledb);
                        ComUtilities.Release(ref connection);
                    }
                }
                return snapshots;
            }
            finally { ComUtilities.Release(ref collection); }
        });
        return JsonSerializer.Serialize(new
        {
            Queries = queries,
            Tables = tables,
            Columns = tables.Select(table => RequireSuccess(_dataModelCommands.ListColumns(batch, table.Name))).ToList(),
            Relationships = relationships,
            Measure = RequireSuccess(_dataModelCommands.Read(batch, _measure)),
            Connections = connections,
            Cells = cells
        });
    }
}
