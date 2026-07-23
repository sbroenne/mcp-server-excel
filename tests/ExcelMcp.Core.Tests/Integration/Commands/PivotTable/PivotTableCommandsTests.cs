using Microsoft.Extensions.Logging;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Commands.PivotTable;
using Sbroenne.ExcelMcp.Core.Tests.Helpers;
using Xunit;
using Xunit.Abstractions;

namespace Sbroenne.ExcelMcp.Core.Tests.Commands.PivotTable;

/// <summary>
/// Integration tests for PivotTable commands.
/// Uses DataModelPivotTableFixture for all tests (shared across ALL test classes via collection fixture).
/// Fixture initialization IS the test for data preparation.
/// Each test gets its own batch for isolation.
/// </summary>
[Collection("DataModel")]
[Trait("Layer", "Core")]
[Trait("Category", "Integration")]
[Trait("RequiresExcel", "true")]
[Trait("Feature", "PivotTables")]
public partial class PivotTableCommandsTests
{
    private readonly PivotTableCommands _pivotCommands;
    private readonly DataModelPivotTableFixture _olapFixture;
    private readonly string _pivotFile;
    private readonly DataModelPivotTableCreationResult _creationResult;
    private readonly ITestOutputHelper _output;
    private readonly ILoggerFactory _loggerFactory;

    /// <summary>
    /// Initializes a new instance of the <see cref="PivotTableCommandsTests"/> class.
    /// </summary>
    public PivotTableCommandsTests(DataModelPivotTableFixture olapFixture, ITestOutputHelper output)
    {
        _pivotCommands = new PivotTableCommands();
        _olapFixture = olapFixture;
        _pivotFile = olapFixture.TestFilePath;
        _creationResult = olapFixture.CreationResult;
        _output = output;
        _loggerFactory = LoggerFactory.Create(builder => builder
            .AddXUnit(output)
            .SetMinimumLevel(LogLevel.Trace));
    }

    /// <summary>
    /// Helper to create a unique copy of the shared sales-data template for pivot table tests.
    /// Used when tests need unique files for specific scenarios.
    /// </summary>
    private string CreateTestFileWithData(string testName) => _olapFixture.CreateSalesDataTestFile(testName);

    /// <summary>
    /// Explicit test that validates the fixture creation results.
    /// This makes the data preparation test visible in test results and validates:
    /// - SessionManager.CreateSessionForNewFile()
    /// - Data Model tables, relationships, measures, and PivotTable creation
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
        Assert.True(_creationResult.TablesCreated > 0, "No tables were created");
        Assert.True(_creationResult.CreationTimeMs > 0);

        // This test appears in test results as proof that creation was tested
        Console.WriteLine($"? Data prepared successfully in {_creationResult.CreationTimeMs}ms");
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

        // Verify data persisted by reading range (SalesData sheet from DataModel fixture)
        batch.Execute((ctx, ct) =>
        {
            dynamic sheet = ctx.Book.Worksheets["SalesData"];

            // Verify headers match DataModel fixture's SalesTable columns
            Assert.Equal("SalesID", sheet.Range["A1"].Value2?.ToString());
            Assert.Equal("Date", sheet.Range["B1"].Value2?.ToString());
            Assert.Equal("CustomerID", sheet.Range["C1"].Value2?.ToString());
            Assert.Equal("ProductID", sheet.Range["D1"].Value2?.ToString());

            // Verify first data row
            Assert.Equal(1.0, Convert.ToDouble(sheet.Range["A2"].Value2));
            Assert.Equal(101.0, Convert.ToDouble(sheet.Range["C2"].Value2));
            Assert.Equal(1001.0, Convert.ToDouble(sheet.Range["D2"].Value2));

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




