using Sbroenne.ExcelMcp.CLI.Tests.Helpers;
using Xunit;

namespace Sbroenne.ExcelMcp.CLI.Tests.Integration;

[Collection("Service")]
[Trait("Category", "Integration")]
[Trait("Feature", "Range")]
[Trait("Layer", "CLI")]
[Trait("RequiresExcel", "true")]
public sealed class RangeScopedReadCliTests : IDisposable
{
    private readonly string _testFile =
        Path.Combine(Path.GetTempPath(), $"RangeScopedReadCli_{Guid.NewGuid():N}.xlsx");

    [Fact]
    public async Task GetValues_WithScope_ReturnsPageAndMetadataThroughCli()
    {
        Sbroenne.ExcelMcp.ComInterop.Session.ExcelSession.CreateNew(
            _testFile,
            isMacroEnabled: false,
            (ctx, ct) => 0,
            CancellationToken.None);

        var (openResult, openJson) = await CliProcessHelper.RunJsonAsync(
            ["session", "open", _testFile],
            timeoutMs: 60000);
        Assert.Equal(0, openResult.ExitCode);
        var sessionId = openJson.RootElement.GetProperty("sessionId").GetString();
        Assert.False(string.IsNullOrWhiteSpace(sessionId));

        try
        {
            var (setupResult, setupJson) = await CliProcessHelper.RunJsonAsync(
                [
                    "range", "set-values",
                    "--session", sessionId!,
                    "--sheet-name", "Sheet1",
                    "--range-address", "B2:E4",
                    "--values", """[["R1B","R1C","R1D","R1E"],["R2B","R2C","R2D","R2E"],["R3B","R3C","R3D","R3E"]]"""
                ],
                timeoutMs: 60000);
            using (setupJson)
            {
                Assert.Equal(0, setupResult.ExitCode);
                Assert.True(setupJson.RootElement.GetProperty("success").GetBoolean(), setupResult.Stdout);
            }

            var (readResult, readJson) = await CliProcessHelper.RunJsonAsync(
                [
                    "range", "get-values",
                    "--session", sessionId!,
                    "--sheet-name", "Sheet1",
                    "--range-address", "B2:E4",
                    "--row-offset", "1",
                    "--row-limit", "1",
                    "--columns", "E,B"
                ],
                timeoutMs: 60000);

            using (readJson)
            {
                Assert.Equal(0, readResult.ExitCode);
                var root = readJson.RootElement;
                Assert.True(root.GetProperty("success").GetBoolean(), readResult.Stdout);
                Assert.Equal(1, root.GetProperty("rowCount").GetInt32());
                Assert.Equal(2, root.GetProperty("columnCount").GetInt32());
                Assert.Equal(3, root.GetProperty("totalRowCount").GetInt32());
                Assert.Equal(4, root.GetProperty("totalColumnCount").GetInt32());
                Assert.Equal(1, root.GetProperty("rowOffset").GetInt32());
                Assert.True(root.GetProperty("hasMoreRows").GetBoolean());
                Assert.Equal(2, root.GetProperty("nextRowOffset").GetInt32());
                Assert.True(root.GetProperty("isTruncated").GetBoolean());
                Assert.Equal(["E", "B"], root.GetProperty("selectedColumns").EnumerateArray().Select(value => value.GetString()));
                Assert.Equal(["R2E", "R2B"], root.GetProperty("values")[0].EnumerateArray().Select(value => value.GetString()));
            }
        }
        finally
        {
            openJson.Dispose();
            await CliProcessHelper.RunAsync(
                ["session", "close", "--session", sessionId!, "--save", "false"],
                timeoutMs: 60000);
        }
    }

    public void Dispose()
    {
        if (File.Exists(_testFile))
        {
            try
            {
                File.Delete(_testFile);
            }
            catch (IOException)
            {
            }
            catch (UnauthorizedAccessException)
            {
            }
        }
    }
}
