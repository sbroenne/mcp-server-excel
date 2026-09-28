using Sbroenne.ExcelMcp.CLI.Tests.Helpers;
using Xunit;

namespace Sbroenne.ExcelMcp.CLI.Tests.Integration;

public sealed class CliWorkbookSessionFixture : IAsyncLifetime
{
    private string? _workbookPath;

    public string SessionId { get; private set; } = string.Empty;

    public async Task InitializeAsync()
    {
        _workbookPath = Path.Join(
            Path.GetTempPath(),
            $"CliClass_{Guid.NewGuid():N}.xlsx");
        var (result, json) = await CliProcessHelper.RunJsonAsync(
            ["session", "create", _workbookPath],
            timeoutMs: 60000);
        Assert.True(
            result.ExitCode == 0,
            $"CLI class session creation failed: {result.Stdout}{result.Stderr}");
        SessionId = json.RootElement.GetProperty("sessionId").GetString() ?? string.Empty;
        Assert.False(string.IsNullOrWhiteSpace(SessionId));
    }

    public async Task DisposeAsync()
    {
        List<Exception>? failures = null;

        if (!string.IsNullOrWhiteSpace(SessionId))
        {
            try
            {
                var result = await CliProcessHelper.RunAsync(
                    ["session", "close", "--session", SessionId, "--save", "false"],
                    timeoutMs: 60000);
                if (result.ExitCode != 0)
                {
                    throw new InvalidOperationException(
                        $"CLI class session close failed: {result.Stdout}{result.Stderr}");
                }
            }
            catch (Exception ex)
            {
                (failures ??= []).Add(ex);
            }
        }

        try
        {
            if (!string.IsNullOrWhiteSpace(_workbookPath)
                && File.Exists(_workbookPath))
            {
                File.Delete(_workbookPath);
            }
        }
        catch (Exception ex)
        {
            (failures ??= []).Add(ex);
        }

        SessionId = string.Empty;
        _workbookPath = null;

        if (failures is not null)
        {
            throw new AggregateException(
                "CLI class workbook fixture cleanup failed.",
                failures);
        }
    }
}
