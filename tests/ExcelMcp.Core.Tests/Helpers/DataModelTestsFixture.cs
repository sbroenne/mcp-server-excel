using System.Diagnostics;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Commands;
using Sbroenne.ExcelMcp.Core.Commands.Table;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Helpers;

/// <summary>
/// Fixture that creates ONE Data Model file per test CLASS.
/// The fixture initialization IS the test for Data Model creation.
/// - Created ONCE before any tests run (~60-120s)
/// - Shared READ-ONLY by all tests in the class
/// - Each test gets its own batch (isolation at batch level)
/// - No file sharing between test classes
/// - Creation results exposed for validation tests
/// </summary>
public class DataModelTestsFixture : IAsyncLifetime
{
    private readonly string _tempDir;

    /// <summary>
    /// Path to the test Data Model file
    /// </summary>
    public string TestFilePath { get; private set; } = null!;

    /// <summary>
    /// Results of Data Model creation (exposed for validation)
    /// </summary>
    public DataModelCreationResult CreationResult { get; private set; } = null!;
    /// <inheritdoc/>

    public DataModelTestsFixture()
    {
        _tempDir = Path.Join(Path.GetTempPath(), $"DataModelTests_{Guid.NewGuid():N}");
        Directory.CreateDirectory(_tempDir);
    }

    /// <summary>
    /// Called ONCE before any tests in the class run.
    /// This IS the test for Data Model creation - if it fails, all tests fail (correct behavior).
    /// Tests: file creation, table creation, AddToDataModel, CreateRelationship, CreateMeasure, persistence.
    /// </summary>
    public async Task InitializeAsync()
    {
        var sw = Stopwatch.StartNew();

        TestFilePath = Path.Join(_tempDir, "DataModel.xlsx");
        CreationResult = new DataModelCreationResult();

        try
        {
            var fileCommands = new FileCommands();
            var createFileResult = await fileCommands.CreateEmptyAsync(TestFilePath);
            if (!createFileResult.Success)
                throw new InvalidOperationException(
                    $"CREATION TEST FAILED: File creation failed: {createFileResult.ErrorMessage}");

            CreationResult.FileCreated = true;

            await using var batch = await ExcelSession.BeginBatchAsync(TestFilePath);

            await CreateSalesTableAsync(batch);
            await CreateCustomersTableAsync(batch);
            await CreateProductsTableAsync(batch);
            CreationResult.TablesCreated = 3;

            var tableCommands = new TableCommands();
            var dataModelCommands = new DataModelCommands();

            var addSales = await tableCommands.AddToDataModelAsync(batch, "SalesTable");
            if (!addSales.Success)
                throw new InvalidOperationException(
                    $"CREATION TEST FAILED: AddToDataModel(SalesTable) failed: {addSales.ErrorMessage}");

            var addCustomers = await tableCommands.AddToDataModelAsync(batch, "CustomersTable");
            if (!addCustomers.Success)
                throw new InvalidOperationException(
                    $"CREATION TEST FAILED: AddToDataModel(CustomersTable) failed: {addCustomers.ErrorMessage}");

            var addProducts = await tableCommands.AddToDataModelAsync(batch, "ProductsTable");
            if (!addProducts.Success)
                throw new InvalidOperationException(
                    $"CREATION TEST FAILED: AddToDataModel(ProductsTable) failed: {addProducts.ErrorMessage}");

            CreationResult.TablesLoadedToModel = 3;

            var rel1 = await dataModelCommands.CreateRelationshipAsync(batch,
                "SalesTable", "CustomerID", "CustomersTable", "CustomerID", active: true);
            if (!rel1.Success)
                throw new InvalidOperationException(
                    $"CREATION TEST FAILED: CreateRelationship(Sales→Customers) failed: {rel1.ErrorMessage}");

            var rel2 = await dataModelCommands.CreateRelationshipAsync(batch,
                "SalesTable", "ProductID", "ProductsTable", "ProductID", active: true);
            if (!rel2.Success)
                throw new InvalidOperationException(
                    $"CREATION TEST FAILED: CreateRelationship(Sales→Products) failed: {rel2.ErrorMessage}");

            CreationResult.RelationshipsCreated = 2;

            var m1 = await dataModelCommands.CreateMeasureAsync(batch, "SalesTable", "Total Sales",
                "SUM(SalesTable[Amount])", "Currency", "Total sales revenue");
            if (!m1.Success)
                throw new InvalidOperationException(
                    $"CREATION TEST FAILED: CreateMeasure(Total Sales) failed: {m1.ErrorMessage}");

            var m2 = await dataModelCommands.CreateMeasureAsync(batch, "SalesTable", "Average Sale",
                "AVERAGE(SalesTable[Amount])", "Currency", "Average sale amount");
            if (!m2.Success)
                throw new InvalidOperationException(
                    $"CREATION TEST FAILED: CreateMeasure(Average Sale) failed: {m2.ErrorMessage}");

            var m3 = await dataModelCommands.CreateMeasureAsync(batch, "SalesTable", "Total Customers",
                "DISTINCTCOUNT(SalesTable[CustomerID])", "WholeNumber", "Unique customer count");
            if (!m3.Success)
                throw new InvalidOperationException(
                    $"CREATION TEST FAILED: CreateMeasure(Total Customers) failed: {m3.ErrorMessage}");

            CreationResult.MeasuresCreated = 3;

            // ═══════════════════════════════════════════════════════
            // TEST 6: Persistence (Save)
            // ═══════════════════════════════════════════════════════
            await batch.SaveAsync();

            sw.Stop();
            CreationResult.Success = true;
            CreationResult.CreationTimeSeconds = sw.Elapsed.TotalSeconds;

        }
        catch (Exception ex)
        {
            CreationResult.Success = false;
            CreationResult.ErrorMessage = ex.Message;

            sw.Stop();

            throw; // Fail all tests in class (correct behavior - no point testing if creation failed)
        }
    }

    /// <summary>
    /// Called ONCE after all tests in the class complete.
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
            // Cleanup is best-effort
        }
        return Task.CompletedTask;
    }

    /// <summary>
    /// Creates Sales worksheet with sample data and formats as Excel Table
    /// </summary>
    private static async Task CreateSalesTableAsync(IExcelBatch batch)
    {
        await batch.Execute((ctx, ct) =>
        {
            dynamic? sheet = null;
            dynamic? range = null;
            dynamic? listObject = null;
            try
            {
                dynamic sheets = ctx.Book.Worksheets;
                sheet = sheets.Add();
                sheet.Name = "Sales";

                // Headers
                sheet.Range["A1"].Value2 = "SalesID";
                sheet.Range["B1"].Value2 = "Date";
                sheet.Range["C1"].Value2 = "CustomerID";
                sheet.Range["D1"].Value2 = "ProductID";
                sheet.Range["E1"].Value2 = "Amount";
                sheet.Range["F1"].Value2 = "Quantity";

                // Sample data (10 rows)
                var salesData = new object[,]
                {
                    { 1, new DateTime(2024, 1, 15), 101, 1001, 1500.00, 3 },
                    { 2, new DateTime(2024, 1, 20), 102, 1002, 2200.00, 2 },
                    { 3, new DateTime(2024, 2, 10), 103, 1003, 750.00, 1 },
                    { 4, new DateTime(2024, 2, 15), 101, 1001, 3000.00, 6 },
                    { 5, new DateTime(2024, 3, 5), 104, 1004, 1200.00, 2 },
                    { 6, new DateTime(2024, 3, 12), 102, 1002, 4400.00, 4 },
                    { 7, new DateTime(2024, 4, 8), 105, 1005, 980.00, 1 },
                    { 8, new DateTime(2024, 4, 22), 103, 1003, 1500.00, 2 },
                    { 9, new DateTime(2024, 5, 10), 104, 1001, 2500.00, 5 },
                    { 10, new DateTime(2024, 5, 25), 101, 1004, 2400.00, 4 }
                };

                range = sheet.Range["A2:F11"];
                range.Value2 = salesData;

                // Format as Excel Table
                range = sheet.Range["A1:F11"];
                listObject = sheet.ListObjects.Add(
                    SourceType: 1, // xlSrcRange
                    Source: range,
                    XlListObjectHasHeaders: 1 // xlYes
                );
                listObject.Name = "SalesTable";
                listObject.TableStyle = "TableStyleMedium2";
            }
            finally
            {
                ComInterop.ComUtilities.Release(ref listObject);
                ComInterop.ComUtilities.Release(ref range);
                ComInterop.ComUtilities.Release(ref sheet);
            }
            return 0;
        });
    }

    /// <summary>
    /// Creates Customers worksheet with sample data and formats as Excel Table
    /// </summary>
    private static async Task CreateCustomersTableAsync(IExcelBatch batch)
    {
        await batch.Execute((ctx, ct) =>
        {
            dynamic? sheet = null;
            dynamic? range = null;
            dynamic? listObject = null;
            try
            {
                dynamic sheets = ctx.Book.Worksheets;
                sheet = sheets.Add();
                sheet.Name = "Customers";

                // Headers
                sheet.Range["A1"].Value2 = "CustomerID";
                sheet.Range["B1"].Value2 = "Name";
                sheet.Range["C1"].Value2 = "Region";
                sheet.Range["D1"].Value2 = "Country";

                // Sample data
                var customersData = new object[,]
                {
                    { 101, "Acme Corp", "North", "USA" },
                    { 102, "TechStart Inc", "South", "USA" },
                    { 103, "Global Solutions", "East", "UK" },
                    { 104, "Innovation Labs", "West", "Canada" },
                    { 105, "Digital Ventures", "North", "USA" }
                };

                range = sheet.Range["A2:D6"];
                range.Value2 = customersData;

                // Format as Excel Table
                range = sheet.Range["A1:D6"];
                listObject = sheet.ListObjects.Add(
                    SourceType: 1, // xlSrcRange
                    Source: range,
                    XlListObjectHasHeaders: 1 // xlYes
                );
                listObject.Name = "CustomersTable";
                listObject.TableStyle = "TableStyleMedium2";
            }
            finally
            {
                ComInterop.ComUtilities.Release(ref listObject);
                ComInterop.ComUtilities.Release(ref range);
                ComInterop.ComUtilities.Release(ref sheet);
            }
            return 0;
        });
    }

    /// <summary>
    /// Creates Products worksheet with sample data and formats as Excel Table
    /// </summary>
    private static async Task CreateProductsTableAsync(IExcelBatch batch)
    {
        await batch.Execute((ctx, ct) =>
        {
            dynamic? sheet = null;
            dynamic? range = null;
            dynamic? listObject = null;
            try
            {
                dynamic sheets = ctx.Book.Worksheets;
                sheet = sheets.Add();
                sheet.Name = "Products";

                // Headers
                sheet.Range["A1"].Value2 = "ProductID";
                sheet.Range["B1"].Value2 = "Name";
                sheet.Range["C1"].Value2 = "Category";
                sheet.Range["D1"].Value2 = "Price";

                // Sample data
                var productsData = new object[,]
                {
                    { 1001, "Laptop Pro", "Electronics", 1200.00 },
                    { 1002, "Desktop Elite", "Electronics", 1500.00 },
                    { 1003, "Tablet Max", "Electronics", 800.00 },
                    { 1004, "Monitor 4K", "Accessories", 450.00 },
                    { 1005, "Keyboard RGB", "Accessories", 120.00 }
                };

                range = sheet.Range["A2:D6"];
                range.Value2 = productsData;

                // Format as Excel Table
                range = sheet.Range["A1:D6"];
                listObject = sheet.ListObjects.Add(
                    SourceType: 1, // xlSrcRange
                    Source: range,
                    XlListObjectHasHeaders: 1 // xlYes
                );
                listObject.Name = "ProductsTable";
                listObject.TableStyle = "TableStyleMedium2";
            }
            finally
            {
                ComInterop.ComUtilities.Release(ref listObject);
                ComInterop.ComUtilities.Release(ref range);
                ComInterop.ComUtilities.Release(ref sheet);
            }
            return 0;
        });
    }
}

/// <summary>
/// Results of Data Model creation (exposed by fixture for validation tests)
/// </summary>
public class DataModelCreationResult
{
    /// <inheritdoc/>
    public bool Success { get; set; }
    /// <inheritdoc/>
    public bool FileCreated { get; set; }
    /// <inheritdoc/>
    public int TablesCreated { get; set; }
    /// <inheritdoc/>
    public int TablesLoadedToModel { get; set; }
    /// <inheritdoc/>
    public int RelationshipsCreated { get; set; }
    /// <inheritdoc/>
    public int MeasuresCreated { get; set; }
    /// <inheritdoc/>
    public double CreationTimeSeconds { get; set; }
    /// <inheritdoc/>
    public string? ErrorMessage { get; set; }
}
