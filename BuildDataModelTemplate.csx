#r "tests/ExcelMcp.Core.Tests/bin/Release/net8.0/Sbroenne.ExcelMcp.Core.dll"
#r "tests/ExcelMcp.Core.Tests/bin/Release/net8.0/Sbroenne.ExcelMcp.ComInterop.dll"

using System;
using System.IO;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Commands;
using Sbroenne.ExcelMcp.Core.Commands.Table;

var targetPath = "tests/ExcelMcp.Core.Tests/TestAssets/DataModelTemplate.xlsx";
Console.WriteLine($"Creating: {targetPath}");

if (File.Exists(targetPath)) File.Delete(targetPath);

var fileCommands = new FileCommands();
var tableCommands = new TableCommands();
var dataModelCommands = new DataModelCommands();

// Create workbook
var createResult = await fileCommands.CreateEmptyAsync(targetPath);
if (!createResult.Success)
{
    Console.WriteLine($"ERROR: {createResult.ErrorMessage}");
    return 1;
}

Console.WriteLine("Created workbook");
var sw = System.Diagnostics.Stopwatch.StartNew();

await using (var batch = await ExcelSession.BeginBatchAsync(targetPath))
{
    Console.WriteLine("Creating Sales table...");
    await batch.Execute<int>((ctx, ct) =>
    {
        dynamic sheet = ctx.Book.Worksheets.Add();
        sheet.Name = "Sales";
        sheet.Range["A1"].Value2 = "SalesID";
        sheet.Range["B1"].Value2 = "CustomerID";
        sheet.Range["C1"].Value2 = "ProductID";
        sheet.Range["D1"].Value2 = "Amount";
        
        sheet.Range["A2"].Value2 = 1; sheet.Range["B2"].Value2 = 101; sheet.Range["C2"].Value2 = 201; sheet.Range["D2"].Value2 = 1500.50;
        sheet.Range["A3"].Value2 = 2; sheet.Range["B3"].Value2 = 102; sheet.Range["C3"].Value2 = 202; sheet.Range["D3"].Value2 = 2300.75;
        
        dynamic range = sheet.Range["A1:D3"];
        dynamic table = sheet.ListObjects.Add(1, range, Type.Missing, 1);
        table.Name = "SalesTable";
        return 0;
    });
    
    Console.WriteLine("Creating Customers table...");
    await batch.Execute<int>((ctx, ct) =>
    {
        dynamic sheet = ctx.Book.Worksheets.Add();
        sheet.Name = "Customers";
        sheet.Range["A1"].Value2 = "CustomerID"; sheet.Range["B1"].Value2 = "Name";
        sheet.Range["A2"].Value2 = 101; sheet.Range["B2"].Value2 = "Customer A";
        sheet.Range["A3"].Value2 = 102; sheet.Range["B3"].Value2 = "Customer B";
        
        dynamic range = sheet.Range["A1:B3"];
        dynamic table = sheet.ListObjects.Add(1, range, Type.Missing, 1);
        table.Name = "CustomersTable";
        return 0;
    });
    
    Console.WriteLine("Creating Products table...");
    await batch.Execute<int>((ctx, ct) =>
    {
        dynamic sheet = ctx.Book.Worksheets.Add();
        sheet.Name = "Products";
        sheet.Range["A1"].Value2 = "ProductID"; sheet.Range["B1"].Value2 = "ProductName";
        sheet.Range["A2"].Value2 = 201; sheet.Range["B2"].Value2 = "Product X";
        sheet.Range["A3"].Value2 = 202; sheet.Range["B3"].Value2 = "Product Y";
        
        dynamic range = sheet.Range["A1:B3"];
        dynamic table = sheet.ListObjects.Add(1, range, Type.Missing, 1);
        table.Name = "ProductsTable";
        return 0;
    });
    
    Console.WriteLine("Adding to Data Model (this takes 30-90 seconds)...");
    var r1 = await tableCommands.AddToDataModelAsync(batch, "SalesTable");
    Console.WriteLine($"  Sales: {r1.Success}");
    var r2 = await tableCommands.AddToDataModelAsync(batch, "CustomersTable");
    Console.WriteLine($"  Customers: {r2.Success}");
    var r3 = await tableCommands.AddToDataModelAsync(batch, "ProductsTable");
    Console.WriteLine($"  Products: {r3.Success}");
    
    Console.WriteLine("Creating relationships...");
    await dataModelCommands.CreateRelationshipAsync(batch, "SalesTable", "CustomerID", "CustomersTable", "CustomerID", true);
    await dataModelCommands.CreateRelationshipAsync(batch, "SalesTable", "ProductID", "ProductsTable", "ProductID", true);
    
    Console.WriteLine("Creating measures...");
    await dataModelCommands.CreateMeasureAsync(batch, "SalesTable", "Total Sales", "SUM(SalesTable[Amount])", "Currency", "Total revenue");
    
    await batch.SaveAsync();
}

sw.Stop();
Console.WriteLine($"✅ Template created in {sw.Elapsed.TotalSeconds:F1}s");
return 0;
