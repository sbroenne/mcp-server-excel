using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Commands;
using Sbroenne.ExcelMcp.Core.Commands.Table;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Commands.DataModel;

/// <summary>
/// Shared fixture for Data Model WRITE tests.
/// Creates ONE fresh Data Model file that is shared across all write tests in the class.
/// This reduces setup time from 60-120s per test to 60-120s total for all write tests.
/// </summary>
public class DataModelWriteTestsFixture : IAsyncLifetime
{
    private readonly string _tempDir;
    public string TestFilePath { get; private set; } = null!;

    public DataModelWriteTestsFixture()
    {
        // Create temp directory for this fixture
        _tempDir = Path.Combine(Path.GetTempPath(), $"DataModelWriteTests_{Guid.NewGuid():N}");
        Directory.CreateDirectory(_tempDir);
    }

    /// <summary>
    /// Called ONCE before any tests in the class run.
    /// Creates a fresh Data Model file that all write tests will use.
    /// </summary>
    public async Task InitializeAsync()
    {
        Console.WriteLine("Creating shared Data Model file for write tests (60-120 seconds)...");
        var sw = System.Diagnostics.Stopwatch.StartNew();

        TestFilePath = Path.Combine(_tempDir, "SharedDataModelForWriteTests.xlsx");

        var fileCommands = new FileCommands();
        var tableCommands = new TableCommands();
        var dataModelCommands = new DataModelCommands();

        // Create empty workbook
        var result = await fileCommands.CreateEmptyAsync(TestFilePath, overwriteIfExists: false);
        if (!result.Success)
        {
            throw new InvalidOperationException($"Failed to create test Excel file: {result.ErrorMessage}");
        }

        // Create Data Model with tables, relationships, and base measures
        await using var batch = await ExcelSession.BeginBatchAsync(TestFilePath);

        // Create worksheets and tables
        await CreateSalesWorksheetAsync(batch);
        await CreateCustomersWorksheetAsync(batch);
        await CreateProductsWorksheetAsync(batch);

        // Add to Data Model
        await tableCommands.AddToDataModelAsync(batch, "SalesTable");
        await tableCommands.AddToDataModelAsync(batch, "CustomersTable");
        await tableCommands.AddToDataModelAsync(batch, "ProductsTable");

        // Create relationships
        await dataModelCommands.CreateRelationshipAsync(batch,
            "SalesTable", "CustomerID", "CustomersTable", "CustomerID", active: true);
        await dataModelCommands.CreateRelationshipAsync(batch,
            "SalesTable", "ProductID", "ProductsTable", "ProductID", active: true);

        // DON'T create base measures here - let tests create their own
        // Creating measures in fixture and then trying to create more in tests causes COM errors

        await batch.SaveAsync();

        sw.Stop();
        Console.WriteLine($"Shared Data Model created in {sw.Elapsed.TotalSeconds:F1}s: {TestFilePath}");
    }

    /// <summary>
    /// Called ONCE after all tests in the class complete.
    /// Cleans up the temp directory.
    /// </summary>
    public Task DisposeAsync()
    {
        try
        {
            if (Directory.Exists(_tempDir))
            {
                Directory.Delete(_tempDir, recursive: true);
            }
        }
        catch
        {
            // Ignore cleanup errors
        }
        return Task.CompletedTask;
    }

    private async Task CreateSalesWorksheetAsync(IExcelBatch batch)
    {
        await batch.Execute<int>((ctx, ct) =>
        {
            dynamic sheet = ctx.Book.Worksheets.Add();
            sheet.Name = "Sales";

            sheet.Range["A1"].Value2 = "SalesID";
            sheet.Range["B1"].Value2 = "CustomerID";
            sheet.Range["C1"].Value2 = "ProductID";
            sheet.Range["D1"].Value2 = "Amount";
            sheet.Range["E1"].Value2 = "Date";

            sheet.Range["A2"].Value2 = 1; sheet.Range["B2"].Value2 = 101; sheet.Range["C2"].Value2 = 201; sheet.Range["D2"].Value2 = 1500.50; sheet.Range["E2"].Value2 = new DateTime(2024, 1, 15);
            sheet.Range["A3"].Value2 = 2; sheet.Range["B3"].Value2 = 102; sheet.Range["C3"].Value2 = 202; sheet.Range["D3"].Value2 = 2300.75; sheet.Range["E3"].Value2 = new DateTime(2024, 1, 16);
            sheet.Range["A4"].Value2 = 3; sheet.Range["B4"].Value2 = 101; sheet.Range["C4"].Value2 = 201; sheet.Range["D4"].Value2 = 800.00; sheet.Range["E4"].Value2 = new DateTime(2024, 1, 17);

            dynamic range = sheet.Range["A1:E4"];
            dynamic tables = sheet.ListObjects;
            dynamic table = tables.Add(1, range, Type.Missing, 1);
            table.Name = "SalesTable";

            return 0;
        });
    }

    private async Task CreateCustomersWorksheetAsync(IExcelBatch batch)
    {
        await batch.Execute<int>((ctx, ct) =>
        {
            dynamic sheet = ctx.Book.Worksheets.Add();
            sheet.Name = "Customers";

            sheet.Range["A1"].Value2 = "CustomerID"; sheet.Range["B1"].Value2 = "Name"; sheet.Range["C1"].Value2 = "Region";
            sheet.Range["A2"].Value2 = 101; sheet.Range["B2"].Value2 = "Customer A"; sheet.Range["C2"].Value2 = "North";
            sheet.Range["A3"].Value2 = 102; sheet.Range["B3"].Value2 = "Customer B"; sheet.Range["C3"].Value2 = "South";

            dynamic range = sheet.Range["A1:C3"];
            dynamic tables = sheet.ListObjects;
            dynamic table = tables.Add(1, range, Type.Missing, 1);
            table.Name = "CustomersTable";

            return 0;
        });
    }

    private async Task CreateProductsWorksheetAsync(IExcelBatch batch)
    {
        await batch.Execute<int>((ctx, ct) =>
        {
            dynamic sheet = ctx.Book.Worksheets.Add();
            sheet.Name = "Products";

            sheet.Range["A1"].Value2 = "ProductID"; sheet.Range["B1"].Value2 = "ProductName"; sheet.Range["C1"].Value2 = "Category";
            sheet.Range["A2"].Value2 = 201; sheet.Range["B2"].Value2 = "Product X"; sheet.Range["C2"].Value2 = "Electronics";
            sheet.Range["A3"].Value2 = 202; sheet.Range["B3"].Value2 = "Product Y"; sheet.Range["C3"].Value2 = "Furniture";

            dynamic range = sheet.Range["A1:C3"];
            dynamic tables = sheet.ListObjects;
            dynamic table = tables.Add(1, range, Type.Missing, 1);
            table.Name = "ProductsTable";

            return 0;
        });
    }
}
