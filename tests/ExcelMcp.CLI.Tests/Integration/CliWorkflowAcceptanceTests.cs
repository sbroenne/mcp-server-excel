using System.Text.Json;
using Sbroenne.ExcelMcp.CLI.Tests.Helpers;
using Xunit;
using Xunit.Abstractions;

namespace Sbroenne.ExcelMcp.CLI.Tests.Integration;

[Collection("Sequential")]
[Trait("Layer", "CLI")]
[Trait("Feature", "CLI")]
[Trait("RequiresExcel", "true")]
[Trait("Acceptance", "Required")]
public sealed class CliWorkflowAcceptanceTests(ITestOutputHelper output) : IAsyncLifetime
{
    private readonly string _workbook = Path.Combine(Path.GetTempPath(), $"CliAcceptance_{Guid.NewGuid():N}.xlsx");
    private readonly Dictionary<string, string> _environment = new()
    {
        ["EXCELMCP_CLI_PIPE"] = Environment.GetEnvironmentVariable("EXCELMCP_CLI_PIPE")
            ?? $"excelmcp-cli-acceptance-{Guid.NewGuid():N}"
    };
    private string? _session;

    public async Task InitializeAsync()
    {
        var created = await SendAsync("session", "create", _workbook);
        _session = created.GetProperty("data").GetProperty("sessionId").GetString();
        Assert.False(string.IsNullOrWhiteSpace(_session));
        Assert.True(File.Exists(_workbook));
    }

    [Fact]
    public async Task Lifecycle_SaveAndReopen_PreservesValue()
    {
        await SendAsync("sheet", "create", "--session", Session, "--sheet-name", "Data");
        await SendAsync("range", "set-values", "--session", Session, "--sheet-name", "Data",
            "--range-address", "A1:D1", "--values", "[[424242,false,0,null]]");
        await SendAsync("range", "set-formulas", "--session", Session, "--sheet-name", "Data",
            "--range-address", "A2", "--formulas", """[["=1/0"]]""");
        var read = await SendAsync("range", "get-values", "--session", Session, "--sheet-name", "Data",
            "--range-address", "A1:D2");
        var data = read.GetProperty("data");
        Assert.Equal(2, data.GetProperty("rowCount").GetInt32());
        Assert.Equal(4, data.GetProperty("columnCount").GetInt32());
        var values = data.GetProperty("values");
        Assert.Equal(2, values.GetArrayLength());
        Assert.Equal(4, values[0].GetArrayLength());
        Assert.Equal(4, values[1].GetArrayLength());
        Assert.Equal(424242, values[0][0].GetInt32());
        Assert.False(values[0][1].GetBoolean());
        Assert.Equal(0, values[0][2].GetInt32());
        Assert.Equal(JsonValueKind.Null, values[0][3].ValueKind);
        Assert.Equal("#DIV/0!", values[1][0].GetString());
        Assert.Single(data.GetProperty("cellErrors").EnumerateArray());
        await SendAsync("session", "close", "--session", Session, "--save");
        _session = null;
        var reopened = await SendAsync("session", "open", _workbook);
        _session = reopened.GetProperty("data").GetProperty("sessionId").GetString();
        Assert.False(string.IsNullOrWhiteSpace(_session));
        Assert.Equal(424242, (await ReadAsync()).GetProperty("data").GetProperty("values")[0][0].GetInt32());
        var persisted = await SendAsync("range", "get-values", "--session", Session, "--sheet-name", "Data",
            "--range-address", "A1:D2");
        Assert.True(JsonElement.DeepEquals(values, persisted.GetProperty("data").GetProperty("values")));
        var formulas = await SendAsync("range", "get-formulas", "--session", Session, "--sheet-name", "Data",
            "--range-address", "A2");
        Assert.Equal("=1/0", formulas.GetProperty("data").GetProperty("formulas")[0][0].GetString());
        var sheets = await SendAsync("sheet", "list", "--session", Session);
        Assert.Contains("Data", sheets.GetProperty("data").GetProperty("worksheets").GetRawText(), StringComparison.Ordinal);
        Assert.True(new FileInfo(_workbook).Length > 0);
    }

    public async Task DisposeAsync()
    {
        var failures = new List<Exception>();
        try
        {
            if (_session is not null)
            {
                await SendAsync("session", "close", "--session", Session, "--save", "false");
                _session = null;
            }
        }
        catch (Exception exception) { failures.Add(exception); }
        try { await SendAsync("service", "stop"); }
        catch (Exception exception) { failures.Add(exception); }
        try
        {
            if (Environment.GetEnvironmentVariable("EXCELMCP_CLI_WORKFLOW_KEEP_FILE") == "true")
            {
                output.WriteLine($"Kept workbook: {_workbook}");
            }
            else { File.Delete(_workbook); }
        }
        catch (Exception exception) { failures.Add(exception); }
        if (failures.Count > 0) { throw new AggregateException("CLI acceptance cleanup failed.", failures); }
    }

    private string Session => _session ?? throw new InvalidOperationException("No CLI acceptance session.");

    private Task<JsonElement> ReadAsync() =>
        SendAsync("range", "get-values", "--session", Session, "--sheet-name", "Data", "--range-address", "A1");

    private async Task<JsonElement> SendAsync(params string[] arguments)
    {
        var (result, json) = await CliProcessHelper.RunJsonAsync(arguments,
            timeoutMs: 60000, environmentVariables: _environment);
        using (json)
        {
            Assert.True(result.ExitCode == 0, $"{result.Stdout}\n{result.Stderr}");
            Assert.True(json.RootElement.GetProperty("success").GetBoolean(), result.Stdout);
            return json.RootElement.Clone();
        }
    }
}
