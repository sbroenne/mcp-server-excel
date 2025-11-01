using System;
using System.IO;
using System.Threading.Tasks;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Commands;

var templatePath = Path.Combine(Directory.GetCurrentDirectory(), "tests", "ExcelMcp.Core.Tests", "bin", "Debug", "net8.0", "TestAssets", "DataModelTemplate.xlsx");

if (!File.Exists(templatePath))
{
    Console.WriteLine($"Template not found: {templatePath}");
    return 1;
}

Console.WriteLine($"Inspecting template: {templatePath}");

var dataModelCommands = new DataModelCommands();
await using var batch = await ExcelSession.BeginBatchAsync(templatePath);

// List tables
var tablesResult = await dataModelCommands.ListTablesAsync(batch);
Console.WriteLine($"\nTables ({tablesResult.Tables.Count}):");
foreach (var table in tablesResult.Tables)
{
    Console.WriteLine($"  - {table.Name}");
}

// List measures  
var measuresResult = await dataModelCommands.ListMeasuresAsync(batch);
Console.WriteLine($"\nMeasures ({measuresResult.Measures.Count}):");
foreach (var measure in measuresResult.Measures)
{
    Console.WriteLine($"  - {measure.Name} (Table: {measure.TableName})");
}

return 0;
