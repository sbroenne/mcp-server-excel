using System.Runtime.ExceptionServices;
using System.Text.Json;
using Xunit;
using Xunit.Abstractions;

namespace Sbroenne.ExcelMcp.CLI.Tests.Helpers;

public abstract class CliNativeWorkbook(ITestOutputHelper output) : IAsyncLifetime
{
    private readonly string _directory = Path.Combine(Path.GetTempPath(), $"CliNative_{Guid.NewGuid():N}");
    private readonly Dictionary<string, string> _environment = new()
    {
        ["EXCELMCP_CLI_PIPE"] = Environment.GetEnvironmentVariable("EXCELMCP_CLI_PIPE")
            ?? $"excelmcp-cli-native-{Guid.NewGuid():N}"
    };
    private string? _session;

    protected string Workbook => Path.Combine(_directory, "native.xlsx");
    protected string ChartImage => Path.Combine(_directory, "chart.png");
    protected string Session => _session ?? throw new InvalidOperationException("No native CLI session.");

    public async Task InitializeAsync()
    {
        Directory.CreateDirectory(_directory);
        _session = Text(await RawAsync("session", "create", Workbook), "sessionId");
        Assert.False(string.IsNullOrWhiteSpace(_session));
        Assert.True(File.Exists(Workbook));
        await CommandAsync("sheet", "create", "--sheet-name", "Data");
        var sheets = await CommandAsync("sheet", "list");
        Assert.Contains(Items(sheets, "worksheets"), sheet => Text(sheet, "name") == "Data");
    }

    protected async Task<JsonElement> CommandAsync(string category, string action, params string[] arguments) =>
        await RawAsync([category, action, "--session", Session, .. arguments]);

    protected Task<JsonElement> RangeAsync(string category, string action, string address, params string[] arguments) =>
        CommandAsync(category, action, ["--sheet-name", "Data", "--range-address", address, .. arguments]);

    protected Task<JsonElement> ValuesAsync(string address, string values) =>
        RangeAsync("range", "set-values", address, "--values", values);

    protected async Task<JsonElement> RawAsync(params string[] arguments)
    {
        var (result, json) = await CliProcessHelper.RunJsonAsync(arguments, 60000, _environment);
        using (json)
        {
            output.WriteLine($"{string.Join(' ', arguments)}\n{result.Stdout}");
            Assert.True(result.ExitCode == 0, $"{result.Stdout}\n{result.Stderr}");
            VerifyOperationResult(json.RootElement);
            return json.RootElement.Clone();
        }
    }

    internal static void VerifyOperationResult(JsonElement response)
    {
        Assert.True(response.TryGetProperty("success", out var success) && success.ValueKind == JsonValueKind.True, response.GetRawText());
        if (response.TryGetProperty("errorMessage", out var error))
        {
            Assert.True(error.ValueKind == JsonValueKind.Null || string.IsNullOrEmpty(error.GetString()), response.GetRawText());
        }
    }

    protected async Task<JsonElement> RejectedAsync(params string[] arguments)
    {
        var (result, json) = await CliProcessHelper.RunJsonAsync(arguments, 60000, _environment);
        using (json)
        {
            Assert.Equal(1, result.ExitCode);
            Assert.False(Bool(json.RootElement, "success"));
            return json.RootElement.Clone();
        }
    }

    protected async Task ReopenAsync()
    {
        await RawAsync("session", "close", "--session", Session, "--save");
        _session = null;
        _session = Text(await RawAsync("session", "open", Workbook), "sessionId");
        Assert.False(string.IsNullOrWhiteSpace(_session));
        Assert.True(new FileInfo(Workbook).Length > 0);
    }

    protected async Task CreatePivotAsync()
    {
        await ValuesAsync("U1:V3", """[["Region","Sales"],["North",100],["South",300]]""");
        await CommandAsync("pivottable", "create-from-range", "--source-sheet", "Data", "--source-range", "U1:V3",
            "--destination-sheet", "Data", "--destination-cell", "X1", "--pivot-table-name", "CalculationPivot");
        await PivotAsync("pivottablefield", "add-row-field", "--field-name", "Region");
        await PivotAsync("pivottablefield", "add-value-field", "--field-name", "Sales", "--custom-name", "Total Sales");
    }

    protected Task<JsonElement> PivotAsync(string category, string action, params string[] arguments) =>
        CommandAsync(category, action, ["--pivot-table-name", "CalculationPivot", .. arguments]);

    protected static JsonElement At(JsonElement value, string path)
    {
        foreach (var part in path.Split('.'))
        {
            value = value.ValueKind == JsonValueKind.Array
                ? value[int.Parse(part, System.Globalization.CultureInfo.InvariantCulture)]
                : value.GetProperty(part);
        }
        return value;
    }

    protected static string Text(JsonElement value, string path) =>
        At(value, path).GetString() ?? throw new InvalidOperationException($"Missing string at {path}.");
    protected static double Number(JsonElement value, string path) => At(value, path).GetDouble();
    protected static bool Bool(JsonElement value, string path) => At(value, path).GetBoolean();
    protected static JsonElement[] Items(JsonElement value, string path) => [.. At(value, path).EnumerateArray()];

    protected static async Task WithCleanupAsync(Func<Task> action, Func<Task> cleanup)
    {
        Exception? primary = null;
        try { await action(); }
        catch (Exception exception) { primary = exception; }
        try { await cleanup(); }
        catch (Exception exception)
        {
            if (primary is not null) { throw new AggregateException("Native CLI assertion and restoration failed.", primary, exception); }
            throw;
        }
        if (primary is not null) { ExceptionDispatchInfo.Capture(primary).Throw(); }
    }

    public async Task DisposeAsync()
    {
        var failures = new List<Exception>();
        try
        {
            if (_session is not null) { await RawAsync("session", "close", "--session", Session, "--save", "false"); }
        }
        catch (Exception exception) { failures.Add(exception); }
        try { await RawAsync("service", "stop"); }
        catch (Exception exception) { failures.Add(exception); }
        try
        {
            if (Environment.GetEnvironmentVariable("EXCELMCP_CLI_WORKFLOW_KEEP_FILE") == "true")
            {
                output.WriteLine($"Kept native CLI files: {_directory}");
            }
            else { Directory.Delete(_directory, recursive: true); }
        }
        catch (Exception exception) { failures.Add(exception); }
        if (failures.Count > 0) { throw new AggregateException("Native CLI cleanup failed.", failures); }
    }
}
