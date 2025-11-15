using System.Diagnostics;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Xunit;
using Xunit.Abstractions;

namespace Sbroenne.ExcelMcp.ComInterop.Tests.Integration;

/// <summary>
/// Smoke tests to verify Excel COM automation is working.
/// These tests verify that Excel responds to basic COM calls.
///
/// Purpose: Catch environment issues (Excel hung, RPC errors, COM registration)
/// before running extensive test suites.
///
/// LAYER RESPONSIBILITY:
/// - ✅ Test basic Excel COM operations (create workbook, read/write cells)
/// - ✅ Verify Excel Application responds to COM calls
/// - ✅ Catch RPC errors (0x800706BE) early
/// </summary>
[Trait("Category", "Integration")]
[Trait("Speed", "Fast")]
[Trait("Layer", "ComInterop")]
[Trait("Feature", "ExcelCom")]
[Trait("RequiresExcel", "true")]
public class ExcelComSmokeTests : IAsyncLifetime
{
    private readonly ITestOutputHelper _output;
    private string _testFile = string.Empty;

    public ExcelComSmokeTests(ITestOutputHelper output)
    {
        _output = output;
    }

    public Task InitializeAsync()
    {
        // Kill any hung Excel processes
        try
        {
            var existingProcesses = Process.GetProcessesByName("EXCEL");
            if (existingProcesses.Length > 0)
            {
                _output.WriteLine($"Cleaning up {existingProcesses.Length} existing Excel processes...");
                foreach (var p in existingProcesses)
                {
                    try { p.Kill(); p.WaitForExit(2000); } catch { }
                }
                Thread.Sleep(2000);
            }
        }
        catch { }

        // Create test file
        _testFile = Path.Join(Path.GetTempPath(), $"excel-com-smoke-{Guid.NewGuid():N}.xlsx");

        return Task.CompletedTask;
    }

    public Task DisposeAsync()
    {
        if (File.Exists(_testFile))
        {
            try { File.Delete(_testFile); } catch { }
        }
        return Task.CompletedTask;
    }

    [Fact]
    public void Excel_CanCreateWorkbook_VerifiesComWorks()
    {
        // Act - Create new workbook via COM
        ExcelSession.CreateNew(_testFile, isMacroEnabled: false, (ctx, ct) =>
        {
            _output.WriteLine("✓ Excel.Application created successfully");
            _output.WriteLine($"✓ Workbook created: {ctx.WorkbookPath}");
            return 0;
        });

        // Assert
        Assert.True(File.Exists(_testFile), "Excel should create workbook file");
        _output.WriteLine("✓ Excel COM automation working");
    }

    [Fact]
    public void Excel_CanWriteAndReadCell_VerifiesDataAccess()
    {
        // Arrange - Create workbook
        ExcelSession.CreateNew(_testFile, isMacroEnabled: false, (ctx, ct) => 0);

        // Act - Write and read cell
        using var batch = ExcelSession.BeginBatch(_testFile);

        var writeResult = batch.Execute((ctx, ct) =>
        {
            dynamic sheet = ctx.Book.Worksheets.Item(1);
            dynamic range = sheet.Range["A1"];
            range.Value2 = "TestValue";

            _output.WriteLine("✓ Written value to A1");
            return 0;
        });

        var readResult = batch.Execute((ctx, ct) =>
        {
            dynamic sheet = ctx.Book.Worksheets.Item(1);
            dynamic range = sheet.Range["A1"];
            string value = range.Value2?.ToString() ?? string.Empty;

            _output.WriteLine($"✓ Read value from A1: {value}");
            return value;
        });

        // Assert
        Assert.Equal("TestValue", readResult);
        _output.WriteLine("✓ Excel data access working");
    }

    [Fact]
    public void Excel_CanAccessWorkbook_VerifiesBookContext()
    {
        // Arrange
        ExcelSession.CreateNew(_testFile, isMacroEnabled: false, (ctx, ct) => 0);

        // Act
        using var batch = ExcelSession.BeginBatch(_testFile);

        var result = batch.Execute((ctx, ct) =>
        {
            // Verify context.Book is accessible
            Assert.NotNull(ctx.Book);

            // Verify basic workbook properties
            dynamic sheets = ctx.Book.Worksheets;
            int sheetCount = sheets.Count;

            _output.WriteLine($"✓ Workbook has {sheetCount} worksheets");

            dynamic sheet1 = sheets.Item(1);
            string sheetName = sheet1.Name;

            _output.WriteLine($"✓ First sheet name: {sheetName}");

            return sheetCount;
        });

        // Assert
        Assert.True(result > 0, "Workbook should have at least one worksheet");
        _output.WriteLine("✓ Excel workbook context working");
    }

    [Fact]
    public void Excel_CanAccessQueries_VerifiesPowerQueryCom()
    {
        // Arrange
        ExcelSession.CreateNew(_testFile, isMacroEnabled: false, (ctx, ct) => 0);

        // Act - Try to access Queries collection (where RPC errors occur)
        using var batch = ExcelSession.BeginBatch(_testFile);

        var exception = Record.Exception(() =>
        {
            batch.Execute((ctx, ct) =>
            {
                dynamic queries = ctx.Book.Queries;
                int queryCount = queries.Count;

                _output.WriteLine($"✓ Power Query COM accessible, {queryCount} queries");
                return queryCount;
            });
        });

        // Assert
        if (exception != null)
        {
            _output.WriteLine($"⚠️ Power Query COM failed: {exception.Message}");
            Assert.Fail($"Power Query COM not working: {exception.Message}");
        }

        _output.WriteLine("✓ Power Query COM working");
    }

    [Fact]
    public void Excel_CanCreateQuery_VerifiesQueryCreation()
    {
        // Arrange
        ExcelSession.CreateNew(_testFile, isMacroEnabled: false, (ctx, ct) => 0);

        string mCode = "let\n    Source = {1, 2, 3},\n    Result = Source\nin\n    Result";

        // Act - Try to create a Power Query (where RPC errors occur)
        using var batch = ExcelSession.BeginBatch(_testFile);

        var exception = Record.Exception(() =>
        {
            batch.Execute((ctx, ct) =>
            {
                dynamic queries = ctx.Book.Queries;
                _output.WriteLine($"Queries collection accessed, Count: {queries.Count}");

                // This is the critical operation that fails with RPC error
                dynamic query = queries.Add("TestQuery", mCode);
                _output.WriteLine($"✓ Query created: {query.Name}");

                return 0;
            });
        });

        // Assert
        if (exception != null)
        {
            _output.WriteLine($"⚠️ Power Query creation failed: {exception.Message}");
            Assert.Fail($"Power Query creation not working: {exception.Message}");
        }

        _output.WriteLine("✓ Power Query creation working");
    }
}


