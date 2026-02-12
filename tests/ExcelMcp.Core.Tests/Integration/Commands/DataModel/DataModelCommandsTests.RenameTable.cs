// Copyright (c) Sbroenne. All rights reserved.
// Licensed under the MIT License.

using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Commands;
using Sbroenne.ExcelMcp.Core.Models;
using Sbroenne.ExcelMcp.Core.Tests.Helpers;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Commands.DataModel;

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
[Trait("Layer", "Core")]
[Trait("Category", "Integration")]
[Trait("RequiresExcel", "true")]
[Trait("Feature", "DataModel")]
[Trait("Speed", "Slow")]
public class DataModelRenameTableTests : IClassFixture<DataModelTestsFixture>
{
    private readonly DataModelCommands _dataModelCommands;
    private readonly PowerQueryCommands _powerQueryCommands;
    private readonly DataModelTestsFixture _fixture;

    public DataModelRenameTableTests(DataModelTestsFixture fixture)
    {
        _dataModelCommands = new DataModelCommands();
        _powerQueryCommands = new PowerQueryCommands(_dataModelCommands);
        _fixture = fixture;
    }

    /// <summary>
    /// Creates a test file with a PQ-backed Data Model table that can be renamed.
    /// PQ-backed tables are created by loading a Power Query with LoadToDataModel mode,
    /// which creates a "Query - {QueryName}" connection with Microsoft.Mashup.OleDb provider.
    /// </summary>
    private string CreateTestFileWithDataModelTable(string testName, string tableName = "TestTable")
    {
        var testFile = _fixture.CreateTestFile(testName);

        using var batch = ExcelSession.BeginBatch(testFile);

        // Create a Power Query that loads to Data Model
        // The M code creates a simple table with sample data
        string mCode = $@"let
    Source = #table(
        type table [ID = Int64.Type, Value = Int64.Type, Category = text],
        {{{{1, 100, ""A""}}, {{2, 200, ""B""}}, {{3, 300, ""A""}}}}
    )
in
    Source";

        // Create PQ with LoadToDataModel - this creates the "Query - {tableName}" connection
        _powerQueryCommands.Create(batch, tableName, mCode, PowerQueryLoadMode.LoadToDataModel);

        batch.Save();

        return testFile;
    }

    /// <summary>
    /// Creates a test file with two PQ-backed Data Model tables for conflict testing.
    /// </summary>
    private string CreateTestFileWithTwoDataModelTables(string testName, string table1Name = "Table1", string table2Name = "Table2")
    {
        var testFile = _fixture.CreateTestFile(testName);

        using var batch = ExcelSession.BeginBatch(testFile);

        // Create first Power Query → Data Model
        string mCode1 = $@"let
    Source = #table(
        type table [ID = Int64.Type, Value = Int64.Type],
        {{{{1, 100}}, {{2, 200}}}}
    )
in
    Source";
        _powerQueryCommands.Create(batch, table1Name, mCode1, PowerQueryLoadMode.LoadToDataModel);

        // Create second Power Query → Data Model
        string mCode2 = $@"let
    Source = #table(
        type table [Category = text, Name = text],
        {{{{""A"", ""Alpha""}}, {{""B"", ""Beta""}}}}
    )
in
    Source";
        _powerQueryCommands.Create(batch, table2Name, mCode2, PowerQueryLoadMode.LoadToDataModel);

        batch.Save();

        return testFile;
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
        var testFile = CreateTestFileWithDataModelTable(
            nameof(RenameTable_PqBackedTable_FailsDueToExcelLimitation),
            "OriginalTable");

        using var batch = ExcelSession.BeginBatch(testFile);

        // Verify table exists in Data Model
        var listBefore = _dataModelCommands.ListTables(batch);
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
        var listAfter = _dataModelCommands.ListTables(batch);
        Assert.True(listAfter.Success);
        Assert.Contains(listAfter.Tables, t => t.Name == "OriginalTable");
        Assert.DoesNotContain(listAfter.Tables, t => t.Name == "RenamedTable");
    }

    /// <summary>
    /// Tests that rename attempts with whitespace are normalized but still fail
    /// due to the Excel limitation on Data Model table names.
    /// </summary>
    [Fact]
    public void RenameTable_WithLeadingTrailingSpaces_FailsDueToExcelLimitation()
    {
        // Arrange
        var testFile = CreateTestFileWithDataModelTable(
            nameof(RenameTable_WithLeadingTrailingSpaces_FailsDueToExcelLimitation),
            "TestTable");

        using var batch = ExcelSession.BeginBatch(testFile);

        // Act - rename with whitespace in new name
        var result = _dataModelCommands.RenameTable(batch, "  TestTable  ", "  TrimmedName  ");

        // Assert - Names are normalized but rename fails
        Assert.False(result.Success);
        Assert.Contains("immutable", result.ErrorMessage, StringComparison.OrdinalIgnoreCase);
        Assert.Equal("TestTable", result.NormalizedOldName);      // Normalized (trimmed)
        Assert.Equal("TrimmedName", result.NormalizedNewName);    // Normalized (trimmed)

        // Verify original table is preserved
        var list = _dataModelCommands.ListTables(batch);
        Assert.Contains(list.Tables, t => t.Name == "TestTable");
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
        var testFile = CreateTestFileWithDataModelTable(
            nameof(RenameTable_SameNameAfterTrim_ReturnsNoOpSuccess),
            "TestTable");

        using var batch = ExcelSession.BeginBatch(testFile);

        // Act - rename to same name (with extra spaces that get trimmed)
        var result = _dataModelCommands.RenameTable(batch, "TestTable", "  TestTable  ");

        // Assert - should be success (no-op)
        Assert.True(result.Success, $"No-op should succeed: {result.ErrorMessage}");
        Assert.Equal("TestTable", result.NormalizedOldName);
        Assert.Equal("TestTable", result.NormalizedNewName);  // Same after normalization
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
        var testFile = CreateTestFileWithDataModelTable(
            nameof(RenameTable_CaseOnlyChange_FailsDueToExcelLimitation),
            "testtable");

        using var batch = ExcelSession.BeginBatch(testFile);

        // Act - rename to same name with different casing
        var result = _dataModelCommands.RenameTable(batch, "testtable", "TestTable");

        // Assert - Rename fails due to Excel limitation
        Assert.False(result.Success);
        Assert.Contains("immutable", result.ErrorMessage, StringComparison.OrdinalIgnoreCase);
        Assert.Equal("testtable", result.OldName);
        Assert.Equal("TestTable", result.NewName);

        // Verify original table is preserved
        var list = _dataModelCommands.ListTables(batch);
        Assert.Contains(list.Tables, t => t.Name.Equals("testtable", StringComparison.OrdinalIgnoreCase));
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
        var testFile = CreateTestFileWithTwoDataModelTables(
            nameof(RenameTable_ConflictWithExistingTable_ReturnsFailure),
            "SourceTable",
            "TargetTable");

        using var batch = ExcelSession.BeginBatch(testFile);

        // Act - try to rename to existing table name
        var result = _dataModelCommands.RenameTable(batch, "SourceTable", "TargetTable");

        // Assert
        Assert.False(result.Success);
        Assert.Contains("already exists", result.ErrorMessage, StringComparison.OrdinalIgnoreCase);
    }

    /// <summary>
    /// Tests that case-insensitive conflict detection works.
    /// </summary>
    [Fact]
    public void RenameTable_CaseInsensitiveConflict_ReturnsFailure()
    {
        // Arrange
        var testFile = CreateTestFileWithTwoDataModelTables(
            nameof(RenameTable_CaseInsensitiveConflict_ReturnsFailure),
            "SourceTable",
            "TARGETTABLE");

        using var batch = ExcelSession.BeginBatch(testFile);

        // Act - try to rename with different case of existing name
        var result = _dataModelCommands.RenameTable(batch, "SourceTable", "targettable");

        // Assert
        Assert.False(result.Success);
        Assert.Contains("already exists", result.ErrorMessage, StringComparison.OrdinalIgnoreCase);
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
        var testFile = CreateTestFileWithDataModelTable(
            nameof(RenameTable_NonExistentTable_ReturnsFailure),
            "ExistingTable");

        using var batch = ExcelSession.BeginBatch(testFile);

        // Act
        var result = _dataModelCommands.RenameTable(batch, "NonExistentTable", "NewName");

        // Assert
        Assert.False(result.Success);
        Assert.Contains("not found", result.ErrorMessage, StringComparison.OrdinalIgnoreCase);
    }

    /// <summary>
    /// Tests that renaming fails clearly when Data Model has no tables.
    /// </summary>
    [Fact]
    public void RenameTable_EmptyDataModel_ReturnsFailure()
    {
        // Arrange - Create file without Data Model
        var testFile = _fixture.CreateTestFile(nameof(RenameTable_EmptyDataModel_ReturnsFailure));

        using var batch = ExcelSession.BeginBatch(testFile);

        // Act
        var result = _dataModelCommands.RenameTable(batch, "AnyTable", "NewName");

        // Assert
        Assert.False(result.Success);
        Assert.Contains("no tables", result.ErrorMessage, StringComparison.OrdinalIgnoreCase);
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
        var testFile = CreateTestFileWithDataModelTable(
            nameof(RenameTable_EmptyNewName_ReturnsFailure),
            "TestTable");

        using var batch = ExcelSession.BeginBatch(testFile);

        // Act
        var result = _dataModelCommands.RenameTable(batch, "TestTable", "");

        // Assert
        Assert.False(result.Success);
        Assert.Contains("empty", result.ErrorMessage, StringComparison.OrdinalIgnoreCase);
    }

    /// <summary>
    /// Tests that whitespace-only new name fails validation.
    /// </summary>
    [Fact]
    public void RenameTable_WhitespaceOnlyNewName_ReturnsFailure()
    {
        // Arrange
        var testFile = CreateTestFileWithDataModelTable(
            nameof(RenameTable_WhitespaceOnlyNewName_ReturnsFailure),
            "TestTable");

        using var batch = ExcelSession.BeginBatch(testFile);

        // Act
        var result = _dataModelCommands.RenameTable(batch, "TestTable", "   ");

        // Assert
        Assert.False(result.Success);
        Assert.Contains("empty", result.ErrorMessage, StringComparison.OrdinalIgnoreCase);
    }

    /// <summary>
    /// Tests that empty old name fails validation.
    /// </summary>
    [Fact]
    public void RenameTable_EmptyOldName_ReturnsFailure()
    {
        // Arrange
        var testFile = CreateTestFileWithDataModelTable(
            nameof(RenameTable_EmptyOldName_ReturnsFailure),
            "TestTable");

        using var batch = ExcelSession.BeginBatch(testFile);

        // Act
        var result = _dataModelCommands.RenameTable(batch, "", "NewName");

        // Assert
        Assert.False(result.Success);
        Assert.Contains("empty", result.ErrorMessage, StringComparison.OrdinalIgnoreCase);
    }

    // ==========================================
    // ROUND-TRIP PERSISTENCE TEST
    // ==========================================

    /// <summary>
    /// Tests that when rename fails, the original table name is preserved across save/reopen.
    /// This verifies the rollback mechanism works correctly.
    /// </summary>
    [Fact]
    public void RenameTable_FailureThenSaveAndReopen_PreservesOriginalTable()
    {
        // Arrange
        var testFile = CreateTestFileWithDataModelTable(
            nameof(RenameTable_FailureThenSaveAndReopen_PreservesOriginalTable),
            "OriginalName");

        // Act - Attempt rename (will fail), then save
        using (var batch1 = ExcelSession.BeginBatch(testFile))
        {
            var result = _dataModelCommands.RenameTable(batch1, "OriginalName", "PersistedName");
            Assert.False(result.Success);  // Rename fails due to Excel limitation
            Assert.Contains("immutable", result.ErrorMessage, StringComparison.OrdinalIgnoreCase);
            batch1.Save();
        }

        // Assert - Reopen and verify original table is preserved
        using var batch2 = ExcelSession.BeginBatch(testFile);
        var list = _dataModelCommands.ListTables(batch2);
        Assert.True(list.Success);
        Assert.Contains(list.Tables, t => t.Name == "OriginalName");  // Original preserved
        Assert.DoesNotContain(list.Tables, t => t.Name == "PersistedName");  // New name not present
    }
}




