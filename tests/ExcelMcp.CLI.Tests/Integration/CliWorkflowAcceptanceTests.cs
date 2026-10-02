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
        _session = created.GetProperty("sessionId").GetString();
        Assert.False(string.IsNullOrWhiteSpace(_session));
        Assert.True(File.Exists(_workbook));
    }

    [Fact]
    public async Task Lifecycle_SaveAndReopen_PreservesValue()
    {
        await SendAsync("sheet", "create", "--session", Session, "--sheet-name", "Data");
        await WriteAsync("424242");
        await SendAsync("session", "close", "--session", Session, "--save");
        _session = null;
        var reopened = await SendAsync("session", "open", _workbook);
        _session = reopened.GetProperty("sessionId").GetString();
        Assert.False(string.IsNullOrWhiteSpace(_session));
        Assert.Equal(424242, (await ReadAsync()).GetProperty("values")[0][0].GetInt32());
        var sheets = await SendAsync("sheet", "list", "--session", Session);
        Assert.Contains("Data", sheets.GetProperty("worksheets").GetRawText(), StringComparison.Ordinal);
        Assert.True(new FileInfo(_workbook).Length > 0);
    }

    [Fact]
    public async Task Editing_ProtectsOccupiedCellsAndDeletesOnlyDisposableSheet()
    {
        await SendAsync("sheet", "create", "--session", Session, "--sheet-name", "Data");
        await WriteAsync("424242");
        var (result, json) = await CliProcessHelper.RunJsonAsync(
            ["range", "set-values", "--session", Session, "--sheet-name", "Data", "--range-address", "A1", "--values", "[[1]]"],
            environmentVariables: _environment);
        using (json)
        {
            Assert.Equal(1, result.ExitCode);
            Assert.False(json.RootElement.GetProperty("success").GetBoolean());
            Assert.Equal("Conflict", json.RootElement.GetProperty("errorCategory").GetString());
            Assert.Contains("$A$1", json.RootElement.GetProperty("errorMessage").GetString(), StringComparison.Ordinal);
        }
        Assert.Equal(424242, (await ReadAsync()).GetProperty("values")[0][0].GetInt32());
        await SendAsync("range", "set-values", "--session", Session, "--sheet-name", "Data",
            "--range-address", "A1", "--values", "[[424242]]", "--overwrite-policy", "allow");
        await SendAsync("sheet", "create", "--session", Session, "--sheet-name", "Disposable");
        await SendAsync("sheet", "delete", "--session", Session, "--sheet-name", "Disposable");
        var sheets = (await SendAsync("sheet", "list", "--session", Session)).GetProperty("worksheets").GetRawText();
        Assert.Contains("Data", sheets, StringComparison.Ordinal);
        Assert.DoesNotContain("Disposable", sheets, StringComparison.Ordinal);
        Assert.Equal(424242, (await ReadAsync()).GetProperty("values")[0][0].GetInt32());
    }

    [Fact]
    public async Task Formatting_TypedAndRepeatedArgumentsRoundTrip()
    {
        await SendAsync("sheet", "create", "--session", Session, "--sheet-name", "Data");
        await SendAsync("rangeformat", "format-ranges", "--session", Session, "--sheet-name", "Data",
            "--range-addresses", "A1:A2", "--range-addresses", "C1:C2", "--bold", "true",
            "--fill-color", "#FFFF00", "--number-format", "0.00");
        foreach (var address in new[] { "A1", "C1" })
        {
            var formats = await SendAsync("range", "get-number-formats", "--session", Session,
                "--sheet-name", "Data", "--range-address", address);
            Assert.Equal("0.00", formats.GetProperty("formats")[0][0].GetString());
        }
        await SendAsync("conditionalformat", "add-rule", "--session", Session, "--sheet-name", "Data",
            "--range-address", "B1:B10", "--rule-type", "top10", "--rank", "7", "--top10-percent", "true",
            "--font-bold", "true", "--font-italic", "false");
        var rules = await SendAsync("conditionalformat", "list-rules", "--session", Session,
            "--sheet-name", "Data", "--range-address", "B1:B10");
        var rule = Assert.Single(rules.GetProperty("rules").EnumerateArray());
        Assert.Equal(7, rule.GetProperty("top10").GetProperty("rank").GetInt32());
        Assert.True(rule.GetProperty("top10").GetProperty("percent").GetBoolean());
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

    private Task<JsonElement> WriteAsync(string value) =>
        SendAsync("range", "set-values", "--session", Session, "--sheet-name", "Data",
            "--range-address", "A1", "--values", $"[[{value}]]");

    private Task<JsonElement> ReadAsync() =>
        SendAsync("range", "get-values", "--session", Session, "--sheet-name", "Data", "--range-address", "A1");

    private async Task<JsonElement> SendAsync(params string[] arguments)
    {
        var (result, json) = await CliProcessHelper.RunJsonAsync(arguments,
            timeoutMs: 60000, environmentVariables: _environment);
        using (json)
        {
            Assert.True(result.ExitCode == 0, $"{result.Stdout}\n{result.Stderr}");
            if (json.RootElement.TryGetProperty("success", out var success)) { Assert.True(success.GetBoolean(), result.Stdout); }
            return json.RootElement.Clone();
        }
    }
}
