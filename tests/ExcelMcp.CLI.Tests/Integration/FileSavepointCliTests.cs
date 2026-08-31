using System.Text.Json;
using Sbroenne.ExcelMcp.CLI.Tests.Helpers;
using Sbroenne.ExcelMcp.Core.Tests.Helpers;
using Xunit;

namespace Sbroenne.ExcelMcp.CLI.Tests.Integration;

[Trait("Category", "Integration")]
[Trait("Feature", "WorkbookSavepoints")]
[Trait("Layer", "CLI")]
[Trait("RequiresExcel", "true")]
[Trait("Speed", "Slow")]
public sealed class FileSavepointCliTests : IAsyncLifetime, IClassFixture<TempDirectoryFixture>
{
    private readonly string _pipeName = $"excelmcp-savepoint-cli-{Guid.NewGuid():N}";
    private readonly string _workbookPath;
    private string? _sessionId;

    public FileSavepointCliTests(TempDirectoryFixture fixture)
    {
        _workbookPath = Path.Combine(
            fixture.TempDir,
            $"FileSavepointCli_{Guid.NewGuid():N}.xlsx");
    }

    private Dictionary<string, string> EnvironmentVariables =>
        new() { ["EXCELMCP_CLI_PIPE"] = _pipeName };

    public async Task InitializeAsync()
    {
        await CliProcessHelper.RunAsync(
            ["service", "stop"],
            timeoutMs: 15_000,
            environmentVariables: EnvironmentVariables,
            diagnosticLabel: "savepoint-cli-initialize-stop");
    }

    public async Task DisposeAsync()
    {
        if (_sessionId != null)
        {
            await CliProcessHelper.RunAsync(
                ["session", "close", "--session", _sessionId],
                timeoutMs: 30_000,
                environmentVariables: EnvironmentVariables,
                diagnosticLabel: "savepoint-cli-cleanup-close");
        }

        await CliProcessHelper.RunAsync(
            ["service", "stop"],
            timeoutMs: 30_000,
            environmentVariables: EnvironmentVariables,
            diagnosticLabel: "savepoint-cli-cleanup-stop");
    }

    [Fact(Timeout = 180_000)]
    public async Task FileSavepointCommands_RestoreStateAndKeepSessionId()
    {
        using (var create = await RunSuccessAsync(
                   ["session", "create", _workbookPath],
                   "savepoint-cli-create"))
        {
            _sessionId = create.RootElement.GetProperty("sessionId").GetString();
            Assert.False(string.IsNullOrWhiteSpace(_sessionId));
        }

        await SetCellAsync("before", "savepoint-cli-set-before");

        using (var savepoint = await RunSuccessAsync(
                   [
                       "file", "create-savepoint",
                       "--session", _sessionId!,
                       "--name", "before-change"
                   ],
                   "savepoint-cli-create-savepoint"))
        {
            Assert.Equal(
                _sessionId,
                savepoint.RootElement.GetProperty("sessionId").GetString());
        }

        await SetCellAsync("after", "savepoint-cli-set-after");

        using (var rollback = await RunSuccessAsync(
                   [
                       "file", "rollback-savepoint",
                       "--session", _sessionId!,
                       "--name", "before-change"
                   ],
                   "savepoint-cli-rollback"))
        {
            Assert.Equal(
                _sessionId,
                rollback.RootElement.GetProperty("sessionId").GetString());
            Assert.True(rollback.RootElement.GetProperty("sessionReopened").GetBoolean());
        }

        using (var value = await RunSuccessAsync(
                   [
                       "range", "get-values",
                       "--session", _sessionId!,
                       "--sheet-name", "Sheet1",
                       "--range-address", "A1"
                   ],
                   "savepoint-cli-get-value"))
        {
            Assert.Equal(
                "before",
                value.RootElement.GetProperty("values")[0][0].GetString());
        }

        using (var release = await RunSuccessAsync(
                   [
                       "file", "release-savepoint",
                       "--session", _sessionId!,
                       "--name", "before-change"
                   ],
                   "savepoint-cli-release"))
        {
            Assert.True(release.RootElement.GetProperty("released").GetBoolean());
        }

        using (var list = await RunSuccessAsync(
                   ["file", "list-savepoints", "--session", _sessionId!],
                   "savepoint-cli-list"))
        {
            Assert.Equal(0, list.RootElement.GetProperty("count").GetInt32());
        }

        using var close = await RunSuccessAsync(
            ["session", "close", "--session", _sessionId!],
            "savepoint-cli-close");
        _sessionId = null;
    }

    private async Task SetCellAsync(string value, string label)
    {
        var values = JsonSerializer.Serialize(new[] { new[] { value } });
        using var _ = await RunSuccessAsync(
            [
                "range", "set-values",
                "--session", _sessionId!,
                "--sheet-name", "Sheet1",
                "--range-address", "A1",
                "--values", values
            ],
            label);
    }

    private async Task<JsonDocument> RunSuccessAsync(
        IReadOnlyList<string> arguments,
        string label)
    {
        var (result, json) = await CliProcessHelper.RunJsonAsync(
            arguments,
            timeoutMs: 60_000,
            environmentVariables: EnvironmentVariables,
            diagnosticLabel: label);
        Assert.True(
            result.ExitCode == 0,
            $"{label} exited with {result.ExitCode}: {result.Stdout}{Environment.NewLine}{result.Stderr}");
        Assert.True(
            json.RootElement.GetProperty("success").GetBoolean(),
            $"{label} failed: {result.Stdout}{Environment.NewLine}{result.Stderr}");
        return json;
    }
}
