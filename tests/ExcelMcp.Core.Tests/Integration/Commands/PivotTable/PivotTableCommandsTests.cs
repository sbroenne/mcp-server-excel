using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Commands;
using Sbroenne.ExcelMcp.Core.Commands.PivotTable;
using Sbroenne.ExcelMcp.Core.Tests.Helpers;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Commands.PivotTable;

/// <summary>
/// Integration tests for PivotTable commands.
/// Uses PivotTableTestsFixture which creates ONE data file per test class (~5-10s setup).
/// Fixture initialization IS the test for data preparation.
/// Each test gets its own batch for isolation.
/// </summary>
[Trait("Layer", "Core")]
[Trait("Category", "Integration")]
[Trait("RequiresExcel", "true")]
[Trait("Feature", "PivotTables")]
public partial class PivotTableCommandsTests : IClassFixture<PivotTableTestsFixture>
{
    protected readonly IPivotTableCommands _pivotCommands;
    protected readonly string _pivotFile;
    protected readonly PivotTableCreationResult _creationResult;
    protected readonly string _tempDir;

    public PivotTableCommandsTests(PivotTableTestsFixture fixture)
    {
        _pivotCommands = new PivotTableCommands();
        _pivotFile = fixture.TestFilePath;
        _creationResult = fixture.CreationResult;
        _tempDir = Path.GetDirectoryName(fixture.TestFilePath)!;
    }

    /// <summary>
    /// Helper to create unique test file with sales data for pivot table tests.
    /// Used when tests need unique files for specific scenarios.
    /// </summary>
    protected async Task<string> CreateTestFileWithDataAsync(string testName)
    {
        var testFile = await CoreTestHelper.CreateUniqueTestFileAsync(
            nameof(PivotTableCommandsTests), testName, _tempDir);

        await using var batch = await ExcelSession.BeginBatchAsync(testFile);

        await batch.Execute<int>((ctx, ct) =>
        {
            dynamic sheet = ctx.Book.Worksheets.Item(1);
            sheet.Name = "SalesData";

            sheet.Range["A1"].Value2 = "Region";
            sheet.Range["B1"].Value2 = "Product";
            sheet.Range["C1"].Value2 = "Sales";
            sheet.Range["D1"].Value2 = "Date";

            sheet.Range["A2"].Value2 = "North";
            sheet.Range["B2"].Value2 = "Widget";
            sheet.Range["C2"].Value2 = 100;
            sheet.Range["D2"].Value2 = new DateTime(2025, 1, 15);

            sheet.Range["A3"].Value2 = "North";
            sheet.Range["B3"].Value2 = "Widget";
            sheet.Range["C3"].Value2 = 150;
            sheet.Range["D3"].Value2 = new DateTime(2025, 1, 20);

            sheet.Range["A4"].Value2 = "South";
            sheet.Range["B4"].Value2 = "Gadget";
            sheet.Range["C4"].Value2 = 200;
            sheet.Range["D4"].Value2 = new DateTime(2025, 2, 10);

            sheet.Range["A5"].Value2 = "North";
            sheet.Range["B5"].Value2 = "Gadget";
            sheet.Range["C5"].Value2 = 75;
            sheet.Range["D5"].Value2 = new DateTime(2025, 2, 15);

            sheet.Range["A6"].Value2 = "South";
            sheet.Range["B6"].Value2 = "Widget";
            sheet.Range["C6"].Value2 = 125;
            sheet.Range["D6"].Value2 = new DateTime(2025, 3, 5);

            return 0;
        });
        
        await batch.SaveAsync();
        
        return testFile;
    }

    /// <summary>
    /// Explicit test that validates the fixture creation results.
    /// This makes the data preparation test visible in test results and validates:
    /// - FileCommands.CreateEmptyAsync()
    /// - Sales data creation
    /// - Batch.SaveAsync() persistence
    /// </summary>
    [Fact]
    [Trait("Speed", "Fast")]
    public void DataPreparation_ViaFixture_CreatesSalesData()
    {
        // Assert the fixture creation succeeded
        Assert.True(_creationResult.Success, 
            $"Data preparation failed during fixture initialization: {_creationResult.ErrorMessage}");
        
        Assert.True(_creationResult.FileCreated, "File creation failed");
        Assert.Equal(5, _creationResult.DataRowsCreated);
        Assert.True(_creationResult.CreationTimeSeconds > 0);
        
        // This test appears in test results as proof that creation was tested
        Console.WriteLine($"✅ Data prepared successfully in {_creationResult.CreationTimeSeconds:F1}s");
    }

    /// <summary>
    /// Tests that sales data persists correctly after file close/reopen.
    /// Validates that SaveAsync() properly persisted the data.
    /// </summary>
    [Fact]
    [Trait("Speed", "Medium")]
    public async Task DataPreparation_Persists_AfterReopenFile()
    {
        // Close and reopen to verify persistence (new batch = new session)
        await using var batch = await ExcelSession.BeginBatchAsync(_pivotFile);
        
        // Verify data persisted by reading range
        await batch.Execute<int>((ctx, ct) =>
        {
            dynamic sheet = ctx.Book.Worksheets.Item("SalesData");
            
            // Verify headers
            Assert.Equal("Region", sheet.Range["A1"].Value2?.ToString());
            Assert.Equal("Product", sheet.Range["B1"].Value2?.ToString());
            Assert.Equal("Sales", sheet.Range["C1"].Value2?.ToString());
            Assert.Equal("Date", sheet.Range["D1"].Value2?.ToString());
            
            // Verify first data row
            Assert.Equal("North", sheet.Range["A2"].Value2?.ToString());
            Assert.Equal("Widget", sheet.Range["B2"].Value2?.ToString());
            Assert.Equal(100.0, Convert.ToDouble(sheet.Range["C2"].Value2));
            
            return 0;
        });
        
        // This proves data creation + save worked correctly
    }
}
