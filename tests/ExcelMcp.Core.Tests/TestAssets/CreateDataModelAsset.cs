using System;
using System.IO;
using System.Threading.Tasks;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Commands;
using Sbroenne.ExcelMcp.Core.Commands.Table;

namespace Sbroenne.ExcelMcp.Core.Tests.TestAssets;

/// <summary>
/// Builder for pre-configured Data Model test asset.
/// Run this once to create DataModelTemplate.xlsx with tables, relationships, and measures.
/// </summary>
public static class DataModelAssetBuilder
{
    public static async Task<string> CreateDataModelAssetAsync(string targetPath)
    {
        Console.WriteLine($"Creating Data Model test asset: {targetPath}");
        
        var fileCommands = new FileCommands();
        var tableCommands = new TableCommands();
        var dataModelCommands = new DataModelCommands();

        // Delete if exists
        if (File.Exists(targetPath))
        {
            File.Delete(targetPath);
            Console.WriteLine("  Deleted existing asset");
        }

        // Create empty workbook
        var createResult = await fileCommands.CreateEmptyAsync(targetPath);
        if (!createResult.Success)
        {
            throw new InvalidOperationException($"Failed to create asset: {createResult.ErrorMessage}");
        }

        var sw = System.Diagnostics.Stopwatch.StartNew();
        await using var batch = await ExcelSession.BeginBatchAsync(targetPath);

        // Create SalesTable
        Console.WriteLine("  Creating SalesTable...");
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

        // Create CustomersTable
        Console.WriteLine("  Creating CustomersTable...");
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

        // Create ProductsTable
        Console.WriteLine("  Creating ProductsTable...");
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

        // Add to Data Model (SLOW: 30-90 seconds)
        Console.WriteLine("  Adding to Data Model (30-90 seconds)...");
        var addSales = await tableCommands.AddToDataModelAsync(batch, "SalesTable");
        var addCustomers = await tableCommands.AddToDataModelAsync(batch, "CustomersTable");
        var addProducts = await tableCommands.AddToDataModelAsync(batch, "ProductsTable");

        // Create relationships
        if (addSales.Success && addCustomers.Success)
        {
            Console.WriteLine("  Creating relationship: Sales-Customers...");
            await dataModelCommands.CreateRelationshipAsync(batch, "SalesTable", "CustomerID", "CustomersTable", "CustomerID", active: true);
        }

        if (addSales.Success && addProducts.Success)
        {
            Console.WriteLine("  Creating relationship: Sales-Products...");
            await dataModelCommands.CreateRelationshipAsync(batch, "SalesTable", "ProductID", "ProductsTable", "ProductID", active: true);
        }

        // Create measures
        Console.WriteLine("  Creating measures...");
        await dataModelCommands.CreateMeasureAsync(batch, "SalesTable", "Total Sales", "SUM(SalesTable[Amount])", "Currency", "Total sales revenue");
        await dataModelCommands.CreateMeasureAsync(batch, "SalesTable", "Average Sale", "AVERAGE(SalesTable[Amount])", "Currency", "Average sale amount");
        await dataModelCommands.CreateMeasureAsync(batch, "SalesTable", "Total Customers", "DISTINCTCOUNT(SalesTable[CustomerID])", "WholeNumber", "Unique customer count");

        await batch.SaveAsync();
        sw.Stop();

        Console.WriteLine($"✅ Asset created in {sw.Elapsed.TotalSeconds:F1}s: {targetPath}");
        return targetPath;
    }
}
