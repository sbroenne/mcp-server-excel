using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Commands;
using Sbroenne.ExcelMcp.Core.Commands.Table;
using Sbroenne.ExcelMcp.Core.Tests.Helpers;
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
public partial class DataModelCommandsTests : IClassFixture<TempDirectoryFixture>
{
    protected readonly IDataModelCommands _dataModelCommands;
    protected readonly IFileCommands _fileCommands;
    protected readonly ITableCommands _tableCommands;
    protected readonly string _tempDir;

    public DataModelCommandsTests(TempDirectoryFixture fixture)
    {
        _dataModelCommands = new DataModelCommands();
        _fileCommands = new FileCommands();
        _tableCommands = new TableCommands();
        _tempDir = fixture.TempDir;
    }

    /// <summary>
    /// Creates a unique test file with Data Model for each test.
    /// Each test gets its own isolated Excel file to prevent test pollution.
    /// Uses pre-built template for READ-ONLY tests (fast), builds fresh Data Model for WRITE tests (slower but necessary).
    /// </summary>
    /// <param name="fileName">Name of the test file to create</param>
    /// <param name="requiresWritableDataModel">If true, creates fresh Data Model instead of using template (needed for Create/Update/Delete tests)</param>
    protected async Task<string> CreateTestFileAsync(string fileName, bool requiresWritableDataModel = false)
    {
        var filePath = Path.Join(_tempDir, fileName);

        if (requiresWritableDataModel)
        {
            // WRITE tests: Build fresh Data Model (slower but supports modifications)
            return await CreateFreshDataModelFileAsync(filePath);
        }
        else
        {
            // READ tests: Use pre-built template (fast - just file copy)
            return await CreateFromTemplateAsync(filePath);
        }
    }

    /// <summary>
    /// Creates a test file by copying the pre-built template (fast, for READ-ONLY tests)
    /// </summary>
    private async Task<string> CreateFromTemplateAsync(string filePath)
    {
        // Path to pre-built Data Model template
        var templatePath = Path.Join(
            Path.GetDirectoryName(typeof(DataModelCommandsTests).Assembly.Location)!,
            "TestAssets",
            "DataModelTemplate.xlsx");

        // If template doesn't exist, create it once (one-time setup)
        if (!File.Exists(templatePath))
        {
            Directory.CreateDirectory(Path.GetDirectoryName(templatePath)!);
            await CreateFreshDataModelFileAsync(templatePath);
        }

        // Copy template to test file location (fast - just file copy ~100ms)
        File.Copy(templatePath, filePath, overwrite: true);
        
        // Ensure the copied file is writable
        var fileInfo = new FileInfo(filePath);
        if (fileInfo.IsReadOnly)
        {
            fileInfo.IsReadOnly = false;
        }

        return filePath;
    }

    /// <summary>
    /// Creates a test file with fresh Data Model built from scratch (slower, but supports WRITE operations)
    /// </summary>
    private async Task<string> CreateFreshDataModelFileAsync(string filePath)
    {
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
            var measure1 = await _dataModelCommands.CreateMeasureAsync(batch, "SalesTable", "Total Sales",
                "SUM(SalesTable[Amount])", "Currency", "Total sales revenue");
            if (!measure1.Success)
                throw new InvalidOperationException($"Failed to create 'Total Sales' measure: {measure1.ErrorMessage}");

            var measure2 = await _dataModelCommands.CreateMeasureAsync(batch, "SalesTable", "Average Sale",
                "AVERAGE(SalesTable[Amount])", "Currency", "Average sale amount");
            if (!measure2.Success)
                throw new InvalidOperationException($"Failed to create 'Average Sale' measure: {measure2.ErrorMessage}");

            var measure3 = await _dataModelCommands.CreateMeasureAsync(batch, "SalesTable", "Total Customers",
                "DISTINCTCOUNT(SalesTable[CustomerID])", "WholeNumber", "Unique customer count");
            if (!measure3.Success)
                throw new InvalidOperationException($"Failed to create 'Total Customers' measure: {measure3.ErrorMessage}");
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
}
