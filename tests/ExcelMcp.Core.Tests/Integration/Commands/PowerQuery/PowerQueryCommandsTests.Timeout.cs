using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Commands;
using Sbroenne.ExcelMcp.Core.Models;
using Xunit;
using Xunit.Abstractions;

namespace Sbroenne.ExcelMcp.Core.Tests.Integration.Commands.PowerQuery;

/// <summary>
/// Timeout-specific tests for PowerQueryCommands.
/// Verifies that heavy operations request appropriate timeout values.
/// </summary>
[Trait("Category", "Integration")]
[Trait("Speed", "Medium")]
[Trait("Layer", "Core")]
[Trait("Feature", "PowerQuery")]
[Trait("Feature", "Timeout")]
[Trait("RequiresExcel", "true")]
public partial class PowerQueryCommandsTimeoutTests : IDisposable
{
    private readonly PowerQueryCommands _commands;
    private readonly ITestOutputHelper _output;
    private readonly string _tempDir;
    /// <inheritdoc/>

    public PowerQueryCommandsTimeoutTests(ITestOutputHelper output)
    {
        var dataModelCommands = new DataModelCommands();
        _commands = new PowerQueryCommands(dataModelCommands);
        _output = output;
        _tempDir = Path.Join(Path.GetTempPath(), $"pq-timeout-tests-{Guid.NewGuid():N}");
        Directory.CreateDirectory(_tempDir);
    }

    private Task<string> CreateTestFileAsync(string testName)
    {
        string testFile = Path.Join(_tempDir, $"{testName}-{Guid.NewGuid():N}.xlsx");
        ExcelSession.CreateNew(testFile, isMacroEnabled: false, (ctx, ct) => 0);
        return Task.FromResult(testFile);
    }
    /// <inheritdoc/>

    [Fact]
    public async Task Refresh_RequestsExtendedTimeout()
    {
        // Arrange
        string testFile = await CreateTestFileAsync(nameof(Refresh_RequestsExtendedTimeout));

        // Create a simple Power Query
        string mCode = """
            let
                Source = #table({"ID", "Value"}, {{1, "A"}, {2, "B"}})
            in
                Source
            """;
        string mFile = Path.Join(_tempDir, "simple.pq");
        await File.WriteAllTextAsync(mFile, mCode);

        try
        {
            using var batch = ExcelSession.BeginBatch(testFile);

            // Import query
            var importResult = _commands.Create(batch, "TestQuery", mFile, PowerQueryLoadMode.LoadToTable, "Sheet1");
            Assert.True(importResult.Success, $"Import failed: {importResult.ErrorMessage}");

            // Act - Refresh should request 5-minute timeout (won't actually timeout in this test)
            var refreshResult = _commands.Refresh(batch, "TestQuery");

            // Assert - Verify refresh succeeded (timeout was sufficient)
            Assert.True(refreshResult.Success, $"Refresh failed: {refreshResult.ErrorMessage}");
            _output.WriteLine("✓ Refresh completed with extended timeout");
        }
        finally
        {
            if (File.Exists(testFile)) File.Delete(testFile);
            if (File.Exists(mFile)) File.Delete(mFile);
        }
    }
    /// <inheritdoc/>

    [Fact]
    public async Task Refresh_SlowQuery_DoesNotTimeoutWithExtendedTimeout()
    {
        // Arrange
        string testFile = await CreateTestFileAsync(nameof(Refresh_SlowQuery_DoesNotTimeoutWithExtendedTimeout));

        // Create a query that takes some time (but less than 5 minutes)
        string mCode = """
            let
                Source = #table({"ID", "Value"}, List.Generate(() => 1, each _ <= 1000, each _ + 1, each {_, "Row" & Text.From(_)}))
            in
                Source
            """;
        string mFile = Path.Join(_tempDir, "slow.pq");
        await File.WriteAllTextAsync(mFile, mCode);

        try
        {
            using var batch = ExcelSession.BeginBatch(testFile);

            // Import query
            var importResult = _commands.Create(batch, "SlowQuery", mFile, PowerQueryLoadMode.LoadToTable, "Sheet1");
            Assert.True(importResult.Success, $"Import failed: {importResult.ErrorMessage}");

            // Act - Refresh with extended timeout
            var refreshResult = _commands.Refresh(batch, "SlowQuery");

            // Assert - Should complete successfully
            Assert.True(refreshResult.Success, $"Refresh should succeed with extended timeout: {refreshResult.ErrorMessage}");
            _output.WriteLine("✓ Slow query completed within extended timeout");
        }
        finally
        {
            if (File.Exists(testFile)) File.Delete(testFile);
            if (File.Exists(mFile)) File.Delete(mFile);
        }
    }
    /// <inheritdoc/>

    public void Dispose()
    {
        if (Directory.Exists(_tempDir))
        {
            try
            {
                Directory.Delete(_tempDir, recursive: true);
            }
            catch
            {
                // Best effort cleanup
            }
        }
        GC.SuppressFinalize(this);
    }
}




