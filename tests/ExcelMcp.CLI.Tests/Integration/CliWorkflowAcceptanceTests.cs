using System.Diagnostics;
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
    public async Task DocumentInspector_ReportsAndRemovesOnlyHiddenWorksheet()
    {
        await SendAsync("sheet", "create", "--session", Session, "--sheet-name", "Data");
        await WriteAsync("424242");
        await SendAsync("sheet", "create", "--session", Session, "--sheet-name", "InspectorHidden");
        await SendAsync("range", "set-values", "--session", Session, "--sheet-name", "InspectorHidden",
            "--range-address", "A1", "--values", "[[123]]");
        await SendAsync("worksheetstyle", "hide", "--session", Session, "--sheet-name", "InspectorHidden");
        var inspection = await SendAsync("workbook", "inspect-document", "--session", Session);
        var selected = Assert.Single(inspection.GetProperty("inspectors").EnumerateArray(),
            item => item.GetProperty("statusCode").GetInt32() == 1);
        var before = await SendAsync("sheet", "list", "--session", Session);
        Assert.Contains(before.GetProperty("worksheets").EnumerateArray(),
            sheet => sheet.GetProperty("name").GetString() == "InspectorHidden");
        var fixedResult = await SendAsync("workbook", "fix-document-inspector", "--session", Session,
            "--inspector-index", selected.GetProperty("index").GetInt32().ToString(System.Globalization.CultureInfo.InvariantCulture),
            "--inspector-name", selected.GetProperty("name").GetString()!);
        Assert.Equal(0, fixedResult.GetProperty("fix").GetProperty("statusCode").GetInt32());
        Assert.Equal(0, fixedResult.GetProperty("inspection").GetProperty("statusCode").GetInt32());
        var after = await SendAsync("sheet", "list", "--session", Session);
        Assert.DoesNotContain(after.GetProperty("worksheets").EnumerateArray(),
            sheet => sheet.GetProperty("name").GetString() == "InspectorHidden");
        Assert.Equal(424242, (await ReadAsync()).GetProperty("values")[0][0].GetInt32());
        Assert.False((await SendAsync("workbook", "get-info", "--session", Session)).GetProperty("saved").GetBoolean());
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
    public async Task BatchStream_DependsOnPreviousResult_LeavesSessionOpenAndPreservesSavedValues()
    {
        await SendAsync("sheet", "create", "--session", Session, "--sheet-name", "Data");
        await SendAsync("range", "set-values", "--session", Session, "--sheet-name", "Data",
            "--range-address", "A1:B1", "--values", "[[10,777]]");

        var start = new ProcessStartInfo(CliProcessHelper.GetExePath())
        {
            UseShellExecute = false,
            RedirectStandardInput = true,
            RedirectStandardOutput = true,
            RedirectStandardError = true
        };
        foreach (var argument in new[] { "-q", "batch", "--stream", "--session", Session })
        {
            start.ArgumentList.Add(argument);
        }
        foreach (var (key, value) in _environment)
        {
            start.Environment[key] = value;
        }
        using var process = Process.Start(start);
        Assert.NotNull(process);
        var error = process.StandardError.ReadToEndAsync();
        try
        {
            var read = await SendEntryAsync(new
            {
                command = "range.get-values",
                args = new { sheetName = "Data", rangeAddress = "A1" }
            }, 0);
            var nextValue = read.GetProperty("values")[0][0].GetInt32() + 7;
            await SendEntryAsync(new
            {
                command = "range.set-values",
                args = new
                {
                    sheetName = "Data",
                    rangeAddress = "A1",
                    values = new[] { new[] { nextValue } },
                    overwritePolicy = "allow"
                }
            }, 1);
            Assert.False(process.HasExited);
            process.StandardInput.Close();
            await process.WaitForExitAsync().WaitAsync(TimeSpan.FromSeconds(15));
            Assert.Equal(0, process.ExitCode);
            Assert.Equal(string.Empty, await error);
        }
        finally
        {
            if (!process.HasExited)
            {
                process.StandardInput.Close();
                try { await process.WaitForExitAsync().WaitAsync(TimeSpan.FromSeconds(15)); }
                finally
                {
                    if (!process.HasExited)
                    {
                        process.Kill(entireProcessTree: true);
                        await process.WaitForExitAsync();
                    }
                }
            }
        }

        var listed = await SendAsync("session", "list");
        Assert.Contains(listed.GetProperty("sessions").EnumerateArray(),
            session => session.GetProperty("sessionId").GetString() == Session);
        await SendAsync("session", "close", "--session", Session, "--save");
        _session = null;
        var reopened = await SendAsync("session", "open", _workbook);
        _session = reopened.GetProperty("sessionId").GetString();
        Assert.False(string.IsNullOrWhiteSpace(_session));
        var values = await SendAsync("range", "get-values", "--session", Session,
            "--sheet-name", "Data", "--range-address", "A1:B1");
        Assert.Equal(17, values.GetProperty("values")[0][0].GetInt32());
        Assert.Equal(777, values.GetProperty("values")[0][1].GetInt32());

        async Task<JsonElement> SendEntryAsync(object command, int index)
        {
            await process.StandardInput.WriteLineAsync(JsonSerializer.Serialize(command));
            await process.StandardInput.FlushAsync();
            var line = await process.StandardOutput.ReadLineAsync().WaitAsync(TimeSpan.FromSeconds(60));
            Assert.False(string.IsNullOrWhiteSpace(line));
            using var result = JsonDocument.Parse(line);
            Assert.Equal(index, result.RootElement.GetProperty("index").GetInt32());
            Assert.True(result.RootElement.GetProperty("success").GetBoolean(), line);
            Assert.False(process.HasExited);
            return result.RootElement.GetProperty("result").Clone();
        }
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
            "--range-address", "A1", "--values", "[[123456]]", "--overwrite-policy", "allow");
        Assert.Equal(123456, (await ReadAsync()).GetProperty("values")[0][0].GetInt32());
        await SendAsync("sheet", "create", "--session", Session, "--sheet-name", "Disposable");
        await SendAsync("sheet", "delete", "--session", Session, "--sheet-name", "Disposable");
        var sheets = (await SendAsync("sheet", "list", "--session", Session)).GetProperty("worksheets").GetRawText();
        Assert.Contains("Data", sheets, StringComparison.Ordinal);
        Assert.DoesNotContain("Disposable", sheets, StringComparison.Ordinal);
        Assert.Equal(123456, (await ReadAsync()).GetProperty("values")[0][0].GetInt32());
    }

    [Fact]
    public async Task Formatting_TypedAndRepeatedArgumentsRoundTrip()
    {
        await SendAsync("sheet", "create", "--session", Session, "--sheet-name", "Data");
        var gapBefore = await SendAsync("rangeformat", "get-format", "--session", Session,
            "--sheet-name", "Data", "--range-address", "B1:B2");
        await SendAsync("rangeformat", "format", "--session", Session, "--sheet-name", "Data",
            "--range-addresses", "A1:A2", "--range-addresses", "C1:C2",
            "--format-options", """{"bold":true,"fillColor":"#FFFF00","numberFormat":"0.00"}""");
        foreach (var address in new[] { "A1:A2", "C1:C2" })
        {
            var formats = await SendAsync("rangeformat", "get-format", "--session", Session,
                "--sheet-name", "Data", "--range-address", address);
            Assert.Equal(2, formats.GetProperty("cellCount").GetInt64());
            var cells = formats.GetProperty("cells");
            Assert.Equal(2, cells.GetArrayLength());
            foreach (var cell in cells.EnumerateArray())
            {
                var stored = cell.GetProperty("stored");
                Assert.Equal("0.00", stored.GetProperty("numberFormat").GetString());
                Assert.True(stored.GetProperty("font").GetProperty("bold").GetBoolean());
                Assert.Equal("#FFFF00", stored.GetProperty("fill").GetProperty("color")
                    .GetProperty("rgb").GetString());
            }
        }
        var gapAfter = await SendAsync("rangeformat", "get-format", "--session", Session,
            "--sheet-name", "Data", "--range-address", "B1:B2");
        Assert.Equal(gapBefore.GetProperty("cells").GetRawText(),
            gapAfter.GetProperty("cells").GetRawText());
        await SendAsync("conditionalformat", "add-rule", "--session", Session, "--sheet-name", "Data",
            "--range-address", "B1:B10", "--rule-type", "top10", "--rank", "7", "--top10-percent", "true",
            "--font-bold", "true", "--font-italic", "false");
        var rules = await SendAsync("conditionalformat", "list-rules", "--session", Session,
            "--sheet-name", "Data", "--range-address", "B1:B10");
        var rule = Assert.Single(rules.GetProperty("rules").EnumerateArray());
        Assert.Equal(7, rule.GetProperty("top10").GetProperty("rank").GetInt32());
        Assert.True(rule.GetProperty("top10").GetProperty("percent").GetBoolean());
        Assert.True(rule.GetProperty("fontBold").GetBoolean());
        Assert.False(rule.GetProperty("fontItalic").GetBoolean());
        Assert.Equal("$B$1:$B$10", rule.GetProperty("appliesTo").GetString());
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
            Assert.True(json.RootElement.GetProperty("success").GetBoolean(), result.Stdout);
            return json.RootElement.Clone();
        }
    }
}
