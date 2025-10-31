using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Commands;
using Sbroenne.ExcelMcp.Core.Commands.Range;
using Sbroenne.ExcelMcp.Core.Commands.Table;
using Sbroenne.ExcelMcp.Core.Tests.Helpers;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Commands.Table;

/// <summary>
/// Integration tests for Table commands (Phase 1 & Phase 2).
/// These tests require Excel installation and validate Core table operations.
/// Tests use Core commands directly (not through CLI wrapper).
/// Each test uses a unique Excel file for complete test isolation.
///
/// Phase 1: Lifecycle, Structure, Filters, Columns, Data, DataModel
/// Phase 2: Structured References, Sorting
/// </summary>
[Trait("Layer", "Core")]
[Trait("Category", "Integration")]
[Trait("RequiresExcel", "true")]
[Trait("Feature", "Tables")]
public partial class TableCommandsTests : IDisposable
{
    private readonly ITableCommands _tableCommands;
    private readonly IRangeCommands _rangeCommands;
    private readonly string _tempDir;
    private bool _disposed;

    public TableCommandsTests()
    {
        _tableCommands = new TableCommands();
        _rangeCommands = new RangeCommands();
        _tempDir = Path.Combine(Path.GetTempPath(), $"ExcelCore_Table_Tests_{Guid.NewGuid():N}");
        Directory.CreateDirectory(_tempDir);
    }

    /// <summary>
    /// Helper to create test file with sample table data
    /// </summary>
    private async Task<string> CreateTestFileWithTableAsync(string fileName)
    {
        var testFile = await CoreTestHelper.CreateUniqueTestFileAsync(
            nameof(TableCommandsTests), fileName, _tempDir);

        await using var batch = await ExcelSession.BeginBatchAsync(testFile);

        // Create worksheet with sample data
        await batch.Execute<int>((ctx, ct) =>
        {
            dynamic sheet = ctx.Book.Worksheets.Item(1);
            sheet.Name = "Sales";

            // Add headers
            sheet.Range["A1"].Value2 = "Region";
            sheet.Range["B1"].Value2 = "Product";
            sheet.Range["C1"].Value2 = "Amount";
            sheet.Range["D1"].Value2 = "Date";

            // Add sample data
            sheet.Range["A2"].Value2 = "North";
            sheet.Range["B2"].Value2 = "Widget";
            sheet.Range["C2"].Value2 = 100;
            sheet.Range["D2"].Value2 = new DateTime(2025, 1, 15);

            sheet.Range["A3"].Value2 = "South";
            sheet.Range["B3"].Value2 = "Gadget";
            sheet.Range["C3"].Value2 = 250;
            sheet.Range["D3"].Value2 = new DateTime(2025, 2, 20);

            sheet.Range["A4"].Value2 = "East";
            sheet.Range["B4"].Value2 = "Widget";
            sheet.Range["C4"].Value2 = 150;
            sheet.Range["D4"].Value2 = new DateTime(2025, 3, 10);

            sheet.Range["A5"].Value2 = "West";
            sheet.Range["B5"].Value2 = "Gadget";
            sheet.Range["C5"].Value2 = 300;
            sheet.Range["D5"].Value2 = new DateTime(2025, 1, 25);

            return 0;
        });

        // Create table from range A1:D5
        var createResult = await _tableCommands.CreateAsync(batch, "Sales", "SalesTable", "A1:D5", true, "TableStyleMedium2");
        if (!createResult.Success)
        {
            throw new InvalidOperationException($"Failed to create test table: {createResult.ErrorMessage}");
        }

        await batch.SaveAsync();

        return testFile;
    }

    public void Dispose()
    {
        if (_disposed) return;

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

        _disposed = true;
        GC.SuppressFinalize(this);
    }
}
