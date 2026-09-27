using Sbroenne.ExcelMcp.ComInterop.Session;
using Xunit;
using Xunit.Abstractions;

namespace Sbroenne.ExcelMcp.McpServer.Tests.Integration.Tools;

[Collection("ProgramTransport")]
[Trait("Category", "Integration")]
[Trait("Layer", "McpServer")]
[Trait("Feature", "SessionManager")]
[Trait("RequiresExcel", "true")]
public sealed class McpProcessOwnershipTests(ITestOutputHelper output)
{
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task Teardown_IndependentExcelSurvives(bool failBeforeShutdown)
    {
        var directory = Path.Join(Path.GetTempPath(), $"McpOwnership_{Guid.NewGuid():N}");
        Directory.CreateDirectory(directory);
        var harness = new OwnershipHarness(output, directory, failBeforeShutdown);
        try
        {
            await harness.InitializeAsync();
            using var independent = ExcelBatch.CreateNewWorkbook(Path.Join(directory, "independent.xlsx"), isMacroEnabled: false);
            var shutdownError = await Record.ExceptionAsync(harness.DisposeAsync);

            Assert.True(independent.IsExcelProcessAlive(), "MCP cleanup must not terminate independently owned Excel.");
            independent.Execute((context, _) => Assert.Equal("independent.xlsx", context.Book.Name));
            if (failBeforeShutdown)
            {
                Assert.IsType<TimeoutException>(shutdownError);
            }
            else
            {
                Assert.Null(shutdownError);
            }
        }
        finally
        {
            await harness.DisposeAsync();
            Directory.Delete(directory, recursive: true);
        }
    }

    private sealed class OwnershipHarness(
        ITestOutputHelper output,
        string directory,
        bool failBeforeShutdown) : McpIntegrationTestBase(output, "OwnershipRegression")
    {
        protected override async Task InitializeTestAsync() =>
            await CreateWorkbookSessionAsync(Path.Join(directory, "owned.xlsx"));

        protected override Task BeforeServerShutdownAsync() =>
            failBeforeShutdown
                ? Task.FromException(new TimeoutException("Simulated test timeout before host cleanup."))
                : Task.CompletedTask;
    }
}
