#!/usr/bin/env dotnet-script
//
// Data Model Template Generator
// ==============================
// Regenerates the test template file with pre-configured Data Model structure.
//
// Usage:
//   1. Close Excel: taskkill /F /IM EXCEL.EXE
//   2. Build tests:  dotnet build -c Debug
//   3. Run script:   dotnet script BuildDataModelTemplate.csx
//
// Only run this when you need to change the template structure.
// The template is stored in git and rarely needs regeneration.
//
#r "bin/Debug/net8.0/Sbroenne.ExcelMcp.Core.dll"
#r "bin/Debug/net8.0/Sbroenne.ExcelMcp.ComInterop.dll"
#r "bin/Debug/net8.0/Sbroenne.ExcelMcp.Core.Tests.dll"

using Sbroenne.ExcelMcp.Core.Tests.TestAssets;

var targetPath = "TestAssets/DataModelTemplate.xlsx";

Console.WriteLine("═══════════════════════════════════════════════════════");
Console.WriteLine("  Data Model Template Generator");
Console.WriteLine("═══════════════════════════════════════════════════════");
Console.WriteLine();
Console.WriteLine($"Target:  {targetPath}");
Console.WriteLine($"Version: {DataModelAssetBuilder.ASSET_VERSION}");
Console.WriteLine();
Console.WriteLine("This will take 60-120 seconds...");
Console.WriteLine();

try
{
    var result = await DataModelAssetBuilder.CreateDataModelAssetAsync(targetPath);

    Console.WriteLine();
    Console.WriteLine("═══════════════════════════════════════════════════════");
    Console.WriteLine("  ✅ SUCCESS!");
    Console.WriteLine("═══════════════════════════════════════════════════════");
    Console.WriteLine();
    Console.WriteLine("Next steps:");
    Console.WriteLine($"  git add {targetPath}");
    Console.WriteLine($"  git commit -m \"test: Update Data Model template\"");
    Console.WriteLine();

    return 0;
}
catch (Exception ex)
{
    Console.WriteLine();
    Console.WriteLine("═══════════════════════════════════════════════════════");
    Console.WriteLine("  ❌ ERROR");
    Console.WriteLine("═══════════════════════════════════════════════════════");
    Console.WriteLine();
    Console.WriteLine(ex.Message);
    Console.WriteLine();
    Console.WriteLine("Make sure:");
    Console.WriteLine("  - Excel is not running (taskkill /F /IM EXCEL.EXE)");
    Console.WriteLine("  - The file is not locked by another process");
    Console.WriteLine();

    return 1;
}
