// Copyright (c) Sbroenne. All rights reserved.
// Licensed under the MIT License.

using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Commands;
using Sbroenne.ExcelMcp.Core.Commands.Table;
using Sbroenne.ExcelMcp.Core.Tests.Helpers;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Commands.DataModel;

/// <summary>
/// Integration tests for DeleteTable operation.
/// Uses isolated test files since DeleteTable is destructive and would affect shared fixtures.
/// </summary>
[Trait("Layer", "Core")]
[Trait("Category", "Integration")]
[Trait("RequiresExcel", "true")]
[Trait("Feature", "DataModel")]
[Trait("Speed", "Slow")]
public class DataModelDeleteTableTests : IClassFixture<DataModelTestsFixture>
{
    private readonly DataModelCommands _dataModelCommands;
    private readonly TableCommands _tableCommands;
    private readonly FileCommands _fileCommands;
    private readonly DataModelTestsFixture _fixture;

    public DataModelDeleteTableTests(DataModelTestsFixture fixture)
    {
        _dataModelCommands = new DataModelCommands();
        _tableCommands = new TableCommands();
        _fileCommands = new FileCommands();
        _fixture = fixture;
    }

    /// <summary>
    /// Creates a test file with a Data Model table that can be deleted.
    /// </summary>
    private string CreateTestFileWithDataModelTable(string testName)
    {
        var testFile = _fixture.CreateTestFile(testName);

        using var batch = ExcelSession.BeginBatch(testFile);

        // Create a simple table with data
        batch.Execute((ctx, ct) =>
        {
            dynamic? sheet = null;
            dynamic? range = null;
            dynamic? listObject = null;
            try
            {
                dynamic sheets = ctx.Book.Worksheets;
                sheet = sheets.Item(1);

                // Headers and data
                sheet.Range["A1"].Value2 = "ID";
                sheet.Range["B1"].Value2 = "Value";
                sheet.Range["A2"].Value2 = 1;
                sheet.Range["B2"].Value2 = 100;
                sheet.Range["A3"].Value2 = 2;
                sheet.Range["B3"].Value2 = 200;

                // Format as Excel Table
                range = sheet.Range["A1:B3"];
                listObject = sheet.ListObjects.Add(
                    SourceType: 1, // xlSrcRange
                    Source: range,
                    XlListObjectHasHeaders: 1 // xlYes
                );
                listObject.Name = "TestTable";
            }
            finally
            {
                ComInterop.ComUtilities.Release(ref listObject);
                ComInterop.ComUtilities.Release(ref range);
                ComInterop.ComUtilities.Release(ref sheet);
            }
            return 0;
        });

        // Add table to Data Model
        _tableCommands.AddToDataModel(batch, "TestTable");

        batch.Save();

        return testFile;
    }

    /// <summary>
    /// Tests deleting a table from the Data Model.
    /// LLM use case: "delete orphaned table from data model"
    /// </summary>
    [Fact]
    public async Task DeleteTable_ExistingTable_RemovesFromDataModel()
    {
        // Arrange
        var testFile = CreateTestFileWithDataModelTable(nameof(DeleteTable_ExistingTable_RemovesFromDataModel));

        using var batch = ExcelSession.BeginBatch(testFile);

        // Verify table exists in Data Model
        var listBefore = await _dataModelCommands.ListTables(batch);
        Assert.True(listBefore.Success);
        Assert.Contains(listBefore.Tables, t => t.Name == "TestTable");

        // Act - Delete the table from Data Model
        _ = _dataModelCommands.DeleteTable(batch, "TestTable");

        // Assert - Table should be gone from Data Model
        var listAfter = await _dataModelCommands.ListTables(batch);
        Assert.True(listAfter.Success);
        Assert.DoesNotContain(listAfter.Tables, t => t.Name == "TestTable");
    }

    /// <summary>
    /// Tests that DeleteTable throws when table doesn't exist.
    /// </summary>
    [Fact]
    public void DeleteTable_NonExistentTable_ThrowsInvalidOperationException()
    {
        // Arrange
        var testFile = CreateTestFileWithDataModelTable(nameof(DeleteTable_NonExistentTable_ThrowsInvalidOperationException));

        using var batch = ExcelSession.BeginBatch(testFile);

        // Act & Assert
        var ex = Assert.Throws<InvalidOperationException>(() =>
            _dataModelCommands.DeleteTable(batch, "NonExistentTable"));

        Assert.Contains("not found", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    /// <summary>
    /// Tests that DeleteTable throws when Data Model has no tables.
    /// </summary>
    [Fact]
    public void DeleteTable_EmptyDataModel_ThrowsInvalidOperationException()
    {
        // Arrange - Create file without Data Model
        var testFile = _fixture.CreateTestFile();

        using var batch = ExcelSession.BeginBatch(testFile);

        // Act & Assert
        var ex = Assert.Throws<InvalidOperationException>(() =>
            _dataModelCommands.DeleteTable(batch, "AnyTable"));

        Assert.Contains("no tables", ex.Message, StringComparison.OrdinalIgnoreCase);
    }
}




