using Sbroenne.ExcelMcp.Core.Commands.Prototypes;

// Test TOM API capabilities
Console.WriteLine("=== TOM API Prototype Test ===");
Console.WriteLine();

TomPrototype.ListTomCapabilities();

Console.WriteLine();
Console.WriteLine("=== Testing with Excel File ===");
Console.WriteLine();

// User should provide path to an Excel file with Data Model
if (args.Length == 0)
{
    Console.WriteLine("Usage: dotnet run <path-to-excel-file-with-data-model>");
    Console.WriteLine();
    Console.WriteLine("The Excel file should:");
    Console.WriteLine("  - Be .xlsx or .xlsm format");
    Console.WriteLine("  - Have Power Pivot / Data Model enabled");
    Console.WriteLine("  - Contain at least one table in the Data Model");
    Console.WriteLine();
    return 1;
}

string excelFile = args[0];

if (!File.Exists(excelFile))
{
    Console.WriteLine($"❌ File not found: {excelFile}");
    return 1;
}

Console.WriteLine($"Testing with: {excelFile}");
Console.WriteLine();

// Test 1: Connection
Console.WriteLine("Test 1: Can Connect to Data Model?");
bool connected = TomPrototype.CanConnectToExcelDataModel(excelFile);
Console.WriteLine();

if (!connected)
{
    Console.WriteLine("❌ Could not connect. Possible reasons:");
    Console.WriteLine("  - File doesn't have Data Model enabled");
    Console.WriteLine("  - MSOLAP provider not installed");
    Console.WriteLine("  - File is locked by Excel");
    Console.WriteLine("  - Excel version doesn't support TOM API");
    return 1;
}

// Test 2: Create Measure (optional - requires table name)
if (args.Length >= 2)
{
    string tableName = args[1];
    Console.WriteLine($"Test 2: Can Create Measure in table '{tableName}'?");
    
    bool measureCreated = TomPrototype.CanCreateMeasure(
        excelFile,
        tableName,
        "TestMeasure_TomPrototype",
        "1" // Simple DAX formula
    );
    
    Console.WriteLine();
    
    if (!measureCreated)
    {
        Console.WriteLine($"❌ Could not create measure. Check table name: '{tableName}'");
    }
}

Console.WriteLine("=== TOM API Prototype Test Complete ===");
return 0;
