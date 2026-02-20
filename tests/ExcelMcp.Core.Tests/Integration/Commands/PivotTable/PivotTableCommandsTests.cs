using Microsoft.Extensions.Logging;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Commands.PivotTable;
using Sbroenne.ExcelMcp.Core.Tests.Helpers;
using Xunit;
using Xunit.Abstractions;

namespace Sbroenne.ExcelMcp.Core.Tests.Commands.PivotTable;

/// <summary>
/// Integration tests for PivotTable commands.
/// Uses PivotTableTestsFixture which creates ONE data file per test class (~5-10s setup).
/// Uses DataModelPivotTableFixture for OLAP tests (shared across ALL test classes via collection fixture).
/// Fixture initialization IS the test for data preparation.
/// Each test gets its own batch for isolation.
/// </summary>
[Collection("DataModel")]
[Trait("Layer", "Core")]
[Trait("Category", "Integration")]
[Trait("RequiresExcel", "true")]
[Trait("Feature", "PivotTables")]
public partial class PivotTableCommandsTests : IClassFixture<PivotTableTestsFixture>
{
    private readonly PivotTableCommands _pivotCommands;
    private readonly PivotTableTestsFixture _fixture;
    private readonly DataModelPivotTableFixture _olapFixture;
    private readonly string _pivotFile;
    private readonly PivotTableCreationResult _creationResult;
    private readonly ITestOutputHelper _output;
    private readonly ILoggerFactory _loggerFactory;

    /// <summary>
    /// Initializes a new instance of the <see cref="PivotTableCommandsTests"/> class.
    /// </summary>
    public PivotTableCommandsTests(PivotTableTestsFixture fixture, DataModelPivotTableFixture olapFixture, ITestOutputHelper output)
    {
        _pivotCommands = new PivotTableCommands();
        _fixture = fixture;
        _olapFixture = olapFixture;
        _pivotFile = fixture.TestFilePath;
        _creationResult = fixture.CreationResult;
        _output = output;
        _loggerFactory = LoggerFactory.Create(builder => builder
            .AddXUnit(output)
            .SetMinimumLevel(LogLevel.Trace));
    }

    /// <summary>
    /// Helper to create unique test file with sales data for pivot table tests.
    /// Used when tests need unique files for specific scenarios.
    /// </summary>
    private string CreateTestFileWithData(string testName)
    {
        var testFile = _fixture.CreateTestFile(testName);

        using var batch = ExcelSession.BeginBatch(testFile);

        batch.Execute((ctx, ct) =>
        {
            dynamic sheet = ctx.Book.Worksheets[1];
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

            // CRITICAL: Format Date column with date format so PivotTable recognizes dates
            // Without this, dates are stored as serial numbers (45672) and Excel won't group them
            sheet.Range["D2:D6"].NumberFormat = "m/d/yyyy";

            return 0;
        });

        batch.Save();

        return testFile;
    }

    /// <summary>
    /// Explicit test that validates the fixture creation results.
    /// This makes the data preparation test visible in test results and validates:
    /// - SessionManager.CreateSessionForNewFile()
    /// - Sales data creation
    /// - Batch.Save() persistence
    /// </summary>
    [Fact]
    [Trait("Speed", "Fast")]
    public void DataPreparation_ViaFixture_CreatesSalesData()
    {
        // Assert the fixture creation succeeded
        Assert.True(_creationResult.Success,
            $"Data preparation failed during fixture initialization: {_creationResult.ErrorMessage}");

        Assert.True(_creationResult.FileCreated, "File creation failed");
        Assert.Equal(5, _creationResult.DataRowsCreated);
        Assert.True(_creationResult.CreationTimeSeconds > 0);

        // This test appears in test results as proof that creation was tested
        Console.WriteLine($"? Data prepared successfully in {_creationResult.CreationTimeSeconds:F1}s");
    }

    /// <summary>
    /// Tests that sales data persists correctly after file close/reopen.
    /// Validates that SaveAsync() properly persisted the data.
    /// </summary>
    [Fact]
    [Trait("Speed", "Medium")]
    public void DataPreparation_Persists_AfterReopenFile()
    {
        // Close and reopen to verify persistence (new batch = new session)
        using var batch = ExcelSession.BeginBatch(_pivotFile);

        // Verify data persisted by reading range
        batch.Execute((ctx, ct) =>
        {
            dynamic sheet = ctx.Book.Worksheets["SalesData"];

            // Verify headers
            Assert.Equal("Region", sheet.Range["A1"].Value2?.ToString());
            Assert.Equal("Product", sheet.Range["B1"].Value2?.ToString());
            Assert.Equal("Sales", sheet.Range["C1"].Value2?.ToString());
            Assert.Equal("Date", sheet.Range["D1"].Value2?.ToString());

            // Verify first data row
            Assert.Equal("North", sheet.Range["A2"].Value2?.ToString());
            Assert.Equal("Widget", sheet.Range["B2"].Value2?.ToString());
            Assert.Equal(100.0, Convert.ToDouble(sheet.Range["C2"].Value2));

            return 0;
        });

        // This proves data creation + save worked correctly
    }
}

/// <summary>
/// Custom logger provider that writes to xUnit output
/// </summary>
internal sealed class TestLoggerProvider : ILoggerProvider
{
    private readonly ITestOutputHelper _output;

    public TestLoggerProvider(ITestOutputHelper output)
    {
        _output = output;
    }

    public ILogger CreateLogger(string categoryName)
    {
        return new TestLogger(_output, categoryName);
    }

    public void Dispose()
    {
    }
}

/// <summary>
/// Custom logger that writes to xUnit output
/// </summary>
internal sealed class TestLogger : ILogger
{
    private readonly ITestOutputHelper _output;
    private readonly string _categoryName;

    public TestLogger(ITestOutputHelper output, string categoryName)
    {
        _output = output;
        _categoryName = categoryName;
    }

    public IDisposable? BeginScope<TState>(TState state) where TState : notnull
    {
        return null;
    }

    public bool IsEnabled(LogLevel logLevel)
    {
        return true;
    }

    public void Log<TState>(
        LogLevel logLevel,
        EventId eventId,
        TState state,
        Exception? exception,
        Func<TState, Exception?, string> formatter)
    {
        var message = formatter(state, exception);
        _output.WriteLine($"[{logLevel}] {_categoryName}: {message}");
    }
}




