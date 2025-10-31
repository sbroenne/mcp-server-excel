using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Commands;
using Sbroenne.ExcelMcp.Core.Commands.Table;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Commands.DataModel;

/// <summary>
/// Base class for Data Model Core operations integration tests.
/// These tests require Excel installation and validate Core Data Model operations.
/// Tests use Core commands directly (not through CLI wrapper).
/// </summary>
[Trait("Layer", "Core")]
[Trait("Category", "Integration")]
[Trait("RequiresExcel", "true")]
[Trait("Feature", "DataModel")]
public partial class DataModelCommandsTests : IDisposable
{
    protected readonly IDataModelCommands _dataModelCommands;
    protected readonly IFileCommands _fileCommands;
    protected readonly ITableCommands _tableCommands;
    protected readonly string _tempDir;
    private bool _disposed;

    public DataModelCommandsTests()
    {
        _dataModelCommands = new DataModelCommands();
        _fileCommands = new FileCommands();
        _tableCommands = new TableCommands();

        // Create temp directory for test files
        _tempDir = Path.Combine(Path.GetTempPath(), $"ExcelCore_DM_Tests_{Guid.NewGuid():N}");
        Directory.CreateDirectory(_tempDir);
    }

    /// <summary>
    /// Creates a unique test file with Data Model for each test.
    /// Each test gets its own isolated Excel file to prevent test pollution.
    /// Uses PRODUCTION commands to create realistic Data Model structure.
    /// </summary>
    protected async Task<string> CreateTestFileAsync(string fileName)
    {
        var filePath = Path.Combine(_tempDir, fileName);

        // Create an empty workbook first
        var result = await _fileCommands.CreateEmptyAsync(filePath, overwriteIfExists: false);
        if (!result.Success)
        {
            throw new InvalidOperationException($"Failed to create test Excel file: {result.ErrorMessage}. Excel may not be installed.");
        }

        // Create realistic Data Model using PRODUCTION commands
        await using var batch = await ExcelSession.BeginBatchAsync(filePath);

        // Create Sales worksheet with data
        await CreateSalesWorksheetAsync(batch);

        // Create Customers worksheet with data
        await CreateCustomersWorksheetAsync(batch);

        // Create Products worksheet with data
        await CreateProductsWorksheetAsync(batch);

        // Add tables to Data Model using PRODUCTION command
        var addSales = await _tableCommands.AddToDataModelAsync(batch, "SalesTable");
        var addCustomers = await _tableCommands.AddToDataModelAsync(batch, "CustomersTable");
        var addProducts = await _tableCommands.AddToDataModelAsync(batch, "ProductsTable");

        // Create relationships using PRODUCTION command
        if (addSales.Success && addCustomers.Success)
        {
            await _dataModelCommands.CreateRelationshipAsync(batch,
                "SalesTable", "CustomerID", "CustomersTable", "CustomerID", active: true);
        }

        if (addSales.Success && addProducts.Success)
        {
            await _dataModelCommands.CreateRelationshipAsync(batch,
                "SalesTable", "ProductID", "ProductsTable", "ProductID", active: true);
        }

        // Create sample measures using PRODUCTION command
        if (addSales.Success)
        {
            await _dataModelCommands.CreateMeasureAsync(batch, "SalesTable", "Total Sales",
                "SUM(SalesTable[Amount])", "Currency", "Total sales revenue");
            await _dataModelCommands.CreateMeasureAsync(batch, "SalesTable", "Average Sale",
                "AVERAGE(SalesTable[Amount])", "Currency", "Average sale amount");
            await _dataModelCommands.CreateMeasureAsync(batch, "SalesTable", "Total Customers",
                "DISTINCTCOUNT(SalesTable[CustomerID])", "General", "Unique customer count");
        }

        await batch.SaveAsync();

        return filePath;
    }

    /// <summary>
    /// Creates Sales worksheet with sample data and formats as Excel Table
    /// </summary>
    private async Task CreateSalesWorksheetAsync(IExcelBatch batch)
    {
        await batch.Execute<int>((ctx, ct) =>
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
    private async Task CreateCustomersWorksheetAsync(IExcelBatch batch)
    {
        await batch.Execute<int>((ctx, ct) =>
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
    private async Task CreateProductsWorksheetAsync(IExcelBatch batch)
    {
        await batch.Execute<int>((ctx, ct) =>
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

    public void Dispose()
    {
        if (_disposed) return;

        try
        {
            if (Directory.Exists(_tempDir))
            {
                // Give Excel time to release file locks
                System.Threading.Thread.Sleep(100);

                // Retry cleanup a few times if needed
                for (int i = 0; i < 3; i++)
                {
                    try
                    {
                        Directory.Delete(_tempDir, recursive: true);
                        break;
                    }
                    catch (IOException) when (i < 2)
                    {
                        System.Threading.Thread.Sleep(500);
                    }
                }
            }
        }
        catch
        {
            // Best effort cleanup
        }

        _disposed = true;
        GC.SuppressFinalize(this);
    }
}
