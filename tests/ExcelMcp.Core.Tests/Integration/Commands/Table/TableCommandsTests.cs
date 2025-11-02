using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Commands;
using Sbroenne.ExcelMcp.Core.Commands.Range;
using Sbroenne.ExcelMcp.Core.Commands.Table;
using Sbroenne.ExcelMcp.Core.Tests.Helpers;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Commands.Table;

/// <summary>
/// Integration tests for Table commands (Phase 1 & Phase 2).
/// Uses TableTestsFixture which creates ONE Table file per test class (~5-10s setup).
/// Fixture initialization IS the test for Table creation - validates CreateAsync command.
/// Each test gets its own batch for isolation.
///
/// Phase 1: Lifecycle, Structure, Filters, Columns, Data, DataModel
/// Phase 2: Structured References, Sorting
/// </summary>
[Trait("Layer", "Core")]
[Trait("Category", "Integration")]
[Trait("RequiresExcel", "true")]
[Trait("Feature", "Tables")]
public partial class TableCommandsTests : IClassFixture<TableTestsFixture>
{
    protected readonly ITableCommands _tableCommands;
    protected readonly IRangeCommands _rangeCommands;
    protected readonly string _tableFile;
    protected readonly TableCreationResult _creationResult;
    protected readonly string _tempDir;

    public TableCommandsTests(TableTestsFixture fixture)
    {
        _tableCommands = new TableCommands();
        _rangeCommands = new RangeCommands();
        _tableFile = fixture.TestFilePath;
        _creationResult = fixture.CreationResult;
        _tempDir = Path.GetDirectoryName(fixture.TestFilePath)!;
    }

    /// <summary>
    /// Helper to create unique test file with SalesTable for modification tests.
    /// Used when tests need to modify the table (delete, rename, resize, etc.) 
    /// without affecting the shared fixture file.
    /// </summary>
    protected async Task<string> CreateTestFileWithTableAsync(string testName)
    {
        var testFile = await CoreTestHelper.CreateUniqueTestFileAsync(
            nameof(TableCommandsTests), testName, _tempDir);

        await using var batch = await ExcelSession.BeginBatchAsync(testFile);

        // Create worksheet with sample data
        await batch.Execute<int>((ctx, ct) =>
        {
            dynamic sheet = ctx.Book.Worksheets.Item(1);
            sheet.Name = "Sales";

            // Add headers
            sheet.Range["A1"].Value2 = "Region";
            sheet.Range["B1"].Value2 = "Product";
            sheet.Range["C1"].Value2 = "Amount";
            sheet.Range["D1"].Value2 = "Date";

            // Add sample data
            sheet.Range["A2"].Value2 = "North";
            sheet.Range["B2"].Value2 = "Widget";
            sheet.Range["C2"].Value2 = 100;
            sheet.Range["D2"].Value2 = new DateTime(2025, 1, 15);

            sheet.Range["A3"].Value2 = "South";
            sheet.Range["B3"].Value2 = "Gadget";
            sheet.Range["C3"].Value2 = 250;
            sheet.Range["D3"].Value2 = new DateTime(2025, 2, 20);

            sheet.Range["A4"].Value2 = "East";
            sheet.Range["B4"].Value2 = "Widget";
            sheet.Range["C4"].Value2 = 150;
            sheet.Range["D4"].Value2 = new DateTime(2025, 3, 10);

            sheet.Range["A5"].Value2 = "West";
            sheet.Range["B5"].Value2 = "Gadget";
            sheet.Range["C5"].Value2 = 300;
            sheet.Range["D5"].Value2 = new DateTime(2025, 1, 25);

            return 0;
        });

        // Create table from range A1:D5
        var createResult = await _tableCommands.CreateAsync(batch, "Sales", "SalesTable", "A1:D5", true, "TableStyleMedium2");
        if (!createResult.Success)
        {
            throw new InvalidOperationException($"Failed to create test table: {createResult.ErrorMessage}");
        }
        
        await batch.SaveAsync();
        
        return testFile;
    }

    /// <summary>
    /// Explicit test that validates the fixture creation results.
    /// This makes the creation test visible in test results and validates:
    /// - FileCommands.CreateEmptyAsync()
    /// - TableCommands.CreateAsync() with sample data
    /// - Batch.SaveAsync() persistence
    /// </summary>
    [Fact]
    [Trait("Speed", "Fast")]
    public void TableCreation_ViaFixture_CreatesSalesTable()
    {
        // Assert the fixture creation succeeded
        Assert.True(_creationResult.Success, 
            $"Table creation failed during fixture initialization: {_creationResult.ErrorMessage}");
        
        Assert.True(_creationResult.FileCreated, "File creation failed");
        Assert.Equal(1, _creationResult.TablesCreated);
        Assert.True(_creationResult.CreationTimeSeconds > 0);
        
        // This test appears in test results as proof that creation was tested
        Console.WriteLine($"✅ Table created successfully in {_creationResult.CreationTimeSeconds:F1}s");
    }

    /// <summary>
    /// Tests that Table persists correctly after file close/reopen.
    /// Validates that SaveAsync() properly persisted the table.
    /// </summary>
    [Fact]
    [Trait("Speed", "Medium")]
    public async Task TableCreation_Persists_AfterReopenFile()
    {
        // Close and reopen to verify persistence (new batch = new session)
        await using var batch = await ExcelSession.BeginBatchAsync(_tableFile);
        
        // Verify table persisted
        var result = await _tableCommands.ListAsync(batch);
        Assert.True(result.Success, $"ListAsync failed: {result.ErrorMessage}");
        Assert.Single(result.Tables);
        Assert.Contains(result.Tables, t => t.Name == "SalesTable");
        
        // Verify table info
        var infoResult = await _tableCommands.GetInfoAsync(batch, "SalesTable");
        Assert.True(infoResult.Success, $"GetInfoAsync failed: {infoResult.ErrorMessage}");
        Assert.Equal(4, infoResult.Table!.Columns?.Count); // Region, Product, Amount, Date
        
        // This proves creation + save worked correctly
    }
}
