using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Commands;
using Sbroenne.ExcelMcp.Core.Commands.PivotTable;
using Sbroenne.ExcelMcp.Core.Tests.Helpers;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Commands.PivotTable;

/// <summary>
/// Integration tests for PivotTable commands.
/// These tests require Excel installation and validate Core pivot table operations.
/// Tests use Core commands directly (not through CLI wrapper).
/// Each test uses a unique Excel file for complete test isolation.
/// </summary>
[Trait("Layer", "Core")]
[Trait("Category", "Integration")]
[Trait("RequiresExcel", "true")]
[Trait("Feature", "PivotTables")]
public partial class PivotTableCommandsTests : IClassFixture<TempDirectoryFixture>
{
    private readonly IPivotTableCommands _pivotCommands;
    private readonly string _tempDir;

    public PivotTableCommandsTests(TempDirectoryFixture fixture)
    {
        _pivotCommands = new PivotTableCommands();
        _tempDir = fixture.TempDir;
    }

    /// <summary>
    /// Helper to create test file with sample sales data for pivot tables
    /// </summary>
    private async Task<string> CreateTestFileWithDataAsync(string fileName)
    {
        var testFile = await CoreTestHelper.CreateUniqueTestFileAsync(
            nameof(PivotTableCommandsTests), fileName, _tempDir);

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
}
