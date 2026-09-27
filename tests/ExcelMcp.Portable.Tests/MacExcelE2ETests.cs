using System.Diagnostics;
using System.Globalization;
using System.IO.Compression;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.Json;
using ModelContextProtocol.Client;
using ModelContextProtocol.Protocol;
using Sbroenne.ExcelMcp.Service.Mac;
using Xunit;
using Xunit.Abstractions;

namespace Sbroenne.ExcelMcp.Portable.Tests;

public sealed class MacExcelTheoryAttribute : TheoryAttribute
{
    public MacExcelTheoryAttribute()
    {
        if (!OperatingSystem.IsMacOS() || Environment.GetEnvironmentVariable("EXCELMCP_MAC_E2E") != "1")
        {
            Skip = "Explicit macOS desktop Excel run required: scripts/Test-MacE2E.ps1.";
        }
    }
}

[CollectionDefinition("Mac Excel E2E", DisableParallelization = true)]
public sealed class MacExcelCollectionDefinition;

[Collection("Mac Excel E2E")]
public sealed class MacExcelE2ETests(ITestOutputHelper output)
{
    private static readonly int[][] SentinelValues = [[9876]];
    private static readonly int[][] UnsavedValues = [[999]];
    private static readonly string[][] SumFormula = [["=SUM(B2:B3)"]];
    private static readonly string[] InitialSheetNames = ["Data", "Spare"];

    [MacExcelTheory]
    [InlineData("cli")]
    [InlineData("mcp")]
    [Trait("Category", "Integration")]
    [Trait("RequiresExcel", "true")]
    [Trait("Feature", "MacHandoff")]
    public async Task ExistingWorkbook_RealEntryPointRoundTrip(string entryPoint)
    {
        Assert.Equal(0, MacAutomationAccess.Check());
        var root = FindRepository();
        var directory = Directory.CreateTempSubdirectory("excelmcp-mac-e2e-");
        var workbookName = $"main with spaces-{Guid.NewGuid():N}.xlsx";
        var main = Path.Combine(directory.FullName, workbookName);
        var created = Path.Combine(directory.FullName, $"created-{Guid.NewGuid():N}.XLSX");
        var createdMacro = Path.Combine(directory.FullName, $"created-{Guid.NewGuid():N}.xlsm");
        var duplicateDirectory = Directory.CreateTempSubdirectory("excelmcp-mac-e2e-duplicate-");
        var duplicate = Path.Combine(duplicateDirectory.FullName, workbookName);
        var sentinel = Path.Combine(directory.FullName, $"sentinel-{Guid.NewGuid():N}.xlsx");
        var dataFile = Path.Combine(directory.FullName, "values.json");
        CreateBlankWorkbook(main);
        File.Copy(main, duplicate);
        File.Copy(main, sentinel);
        await File.WriteAllTextAsync(dataFile, """[["Item","Amount"],["Alpha",10],["Beta",20]]""");
        using var deadline = new CancellationTokenSource(TimeSpan.FromMinutes(4));
        await using var client = await EntryPointClient.CreateAsync(root, entryPoint, output, deadline.Token);
        var completed = false;
        try
        {
            var mainSession = SessionId(await client.CallAsync("file", "open", null,
                new() { ["path"] = main }, deadline.Token));
            var duplicateName = await client.CallAsync("file", "open", null,
                new() { ["path"] = duplicate }, deadline.Token);
            Assert.False(duplicateName.GetProperty("success").GetBoolean());
            Assert.Contains("same name", duplicateName.GetProperty("errorMessage").GetString(),
                StringComparison.OrdinalIgnoreCase);
            Assert.Equal("ComInterop", duplicateName.GetProperty("errorCategory").GetString());

            var sheets = Success(await client.CallAsync("sheet", "list", mainSession, new(), deadline.Token));
            Assert.Equal(InitialSheetNames, sheets.GetProperty("worksheets").EnumerateArray()
                .Select(sheet => sheet.GetProperty("name").GetString()!).ToArray());
            Success(await client.CallAsync("sheet", "rename", mainSession,
                new() { ["old_name"] = "Spare", ["new_name"] = "DeleteMe" }, deadline.Token));
            Success(await client.CallAsync("sheet", "delete", mainSession,
                new() { ["sheet_name"] = "DeleteMe" }, deadline.Token));
            sheets = Success(await client.CallAsync("sheet", "list", mainSession, new(), deadline.Token));
            Assert.Equal("Data", Assert.Single(sheets.GetProperty("worksheets").EnumerateArray()).GetProperty("name").GetString());
            Success(await client.CallAsync("sheet", "create", mainSession,
                new() { ["sheet_name"] = "Created" }, deadline.Token));
            sheets = Success(await client.CallAsync("sheet", "list", mainSession, new(), deadline.Token));
            Assert.Contains(
                sheets.GetProperty("worksheets").EnumerateArray(),
                sheet => sheet.GetProperty("name").GetString() == "Created");
            Success(await client.CallAsync("sheet", "delete", mainSession,
                new() { ["sheet_name"] = "Created" }, deadline.Token));
            var dirtyPowerQueryRead = await client.CallAsync(
                "powerquery", "list", mainSession, new(), deadline.Token);
            Assert.False(dirtyPowerQueryRead.GetProperty("success").GetBoolean());
            Assert.Contains(
                "saved workbook",
                dirtyPowerQueryRead.GetProperty("errorMessage").GetString(),
                StringComparison.OrdinalIgnoreCase);

            var createdSession = SessionId(await client.CallAsync("file", "create", null,
                new() { ["path"] = created }, deadline.Token));
            Success(await client.CallAsync("range", "set-values", createdSession,
                RangeArgsOnSheet("Sheet1", "A1:B1", ("values", new object?[][] { ["keep", "clear"] })), deadline.Token));
            Success(await client.CallAsync("range", "clear-formats", createdSession,
                RangeArgsOnSheet("Sheet1", "A1"), deadline.Token));
            var preserved = Success(await client.CallAsync("range", "get-values", createdSession,
                RangeArgsOnSheet("Sheet1", "A1"), deadline.Token));
            Assert.Equal("keep", preserved.GetProperty("values")[0][0].GetString());
            Success(await client.CallAsync("range", "clear-contents", createdSession,
                RangeArgsOnSheet("Sheet1", "B1"), deadline.Token));
            Success(await client.CallAsync("range", "clear-all", createdSession,
                RangeArgsOnSheet("Sheet1", "A1"), deadline.Token));
            var cleared = Success(await client.CallAsync("range", "get-values", createdSession,
                RangeArgsOnSheet("Sheet1", "A1:B1"), deadline.Token));
            Assert.All(cleared.GetProperty("values").EnumerateArray().SelectMany(row => row.EnumerateArray()),
                value => Assert.True(value.ValueKind is JsonValueKind.Null
                    || value.ValueKind == JsonValueKind.String && string.IsNullOrEmpty(value.GetString())));
            Success(await client.CallAsync("file", "close", createdSession, new(), deadline.Token));
            Success(await client.CallAsync("file", "close", createdSession, new(), deadline.Token));

            var macroSession = SessionId(await client.CallAsync("file", "create", null,
                new() { ["path"] = createdMacro }, deadline.Token));
            var unsupportedVba = await client.CallAsync("vba", "run", macroSession,
                new() { ["procedure_name"] = "Module1.NotAvailable", ["timeout"] = 1 }, deadline.Token);
            Assert.False(unsupportedVba.GetProperty("success").GetBoolean());
            Assert.Equal("PlatformNotSupported", unsupportedVba.GetProperty("errorCategory").GetString());
            Assert.Contains(
                "preflight reports",
                unsupportedVba.GetProperty("errorMessage").GetString(),
                StringComparison.OrdinalIgnoreCase);
            Assert.Contains(
                "repository-owned synthetic fixture",
                unsupportedVba.GetProperty("errorMessage").GetString(),
                StringComparison.OrdinalIgnoreCase);
            var unsupportedVbaList = await client.CallAsync(
                "vba", "list", macroSession, new(), deadline.Token);
            Assert.False(unsupportedVbaList.GetProperty("success").GetBoolean());
            Assert.Equal(
                "PlatformNotSupported",
                unsupportedVbaList.GetProperty("errorCategory").GetString());
            Assert.Contains(
                "project object model",
                unsupportedVbaList.GetProperty("errorMessage").GetString(),
                StringComparison.OrdinalIgnoreCase);
            Assert.Contains(
                "scripting dictionary",
                unsupportedVbaList.GetProperty("errorMessage").GetString(),
                StringComparison.OrdinalIgnoreCase);
            Success(await client.CallAsync("file", "close", macroSession, new(), deadline.Token));

            var sentinelSession = SessionId(await client.CallAsync("file", "open", null,
                new() { ["path"] = sentinel }, deadline.Token));
            Success(await client.CallAsync("range", "set-values", sentinelSession,
                RangeArgs("A1", ("values", SentinelValues)), deadline.Token));
            Success(await client.CallAsync("range", "set-values", mainSession,
                RangeArgs("A1:B3", ("values_file", dataFile)), deadline.Token));
            Success(await client.CallAsync("range", "set-number-format", mainSession,
                RangeArgs("B2:B3", ("format_code", "0.00")), deadline.Token));
            var numberFormats = Success(await client.CallAsync("range", "get-number-formats", mainSession,
                RangeArgs("B2:B3"), deadline.Token));
            Assert.Equal(2, numberFormats.GetProperty("rowCount").GetInt32());
            Assert.Equal(1, numberFormats.GetProperty("columnCount").GetInt32());
            Assert.All(
                numberFormats.GetProperty("formats").EnumerateArray(),
                row => Assert.Equal("0.00", row[0].GetString()));
            Success(await client.CallAsync("rangeformat", "set-column-width", mainSession,
                RangeArgs("A:B", ("column_width", 14)), deadline.Token));
            Success(await client.CallAsync("rangeformat", "set-row-height", mainSession,
                RangeArgs("1:3", ("row_height", 20)), deadline.Token));
            Success(await client.CallAsync("range", "set-formulas", mainSession,
                RangeArgs("C1", ("formulas", SumFormula)), deadline.Token));
            Success(await client.CallAsync("calculation_mode", "calculate", mainSession,
                RangeArgs("C1", ("scope", "range")), deadline.Token));

            var values = Success(await client.CallAsync("range", "get-values", mainSession,
                RangeArgs("A1:B3"), deadline.Token));
            Assert.Equal(3, values.GetProperty("rowCount").GetInt32());
            Assert.Equal(2, values.GetProperty("columnCount").GetInt32());
            Assert.Equal("Alpha", values.GetProperty("values")[1][0].GetString());
            Assert.Equal(20, values.GetProperty("values")[2][1].GetDouble());
            var formula = Success(await client.CallAsync("range", "get-formulas", mainSession,
                RangeArgs("C1"), deadline.Token));
            Assert.Equal("=SUM(B2:B3)", formula.GetProperty("formulas")[0][0].GetString());
            Assert.Equal(30, formula.GetProperty("values")[0][0].GetDouble());
            Assert.Equal(1, formula.GetProperty("rowCount").GetInt32());
            Assert.Equal(1, formula.GetProperty("columnCount").GetInt32());

            var missing = RangeArgs("A1");
            missing["sheet_name"] = "Missing";
            var error = await client.CallAsync("range", "get-values", mainSession, missing, deadline.Token);
            Assert.False(error.GetProperty("success").GetBoolean());
            Assert.False(string.IsNullOrWhiteSpace(error.GetProperty("errorMessage").GetString()));
            Success(await client.CallAsync("range", "get-values", mainSession, RangeArgs("A1"), deadline.Token));

            Success(await client.CallAsync("file", "close", mainSession, new() { ["save"] = true }, deadline.Token));
            mainSession = SessionId(await client.CallAsync("file", "open", null, new() { ["path"] = main }, deadline.Token));
            var powerQueries = Success(await client.CallAsync("powerquery", "list", mainSession, new(), deadline.Token));
            Assert.Empty(powerQueries.GetProperty("queries").EnumerateArray());
            var saved = Success(await client.CallAsync("range", "get-values", mainSession, RangeArgs("C1"), deadline.Token));
            Assert.Equal(30, saved.GetProperty("values")[0][0].GetDouble());
            Success(await client.CallAsync("range", "set-values", mainSession,
                RangeArgs("B2", ("values", UnsavedValues)), deadline.Token));
            var dirtyClose = await InvokeAutomationHostAsync(
                root,
                "session.close-if-saved",
                new { filePath = main },
                deadline.Token);
            Assert.False(dirtyClose.GetProperty("success").GetBoolean());
            Assert.Equal("InvalidOperation", dirtyClose.GetProperty("errorCategory").GetString());
            Assert.Contains(
                "saved workbook",
                dirtyClose.GetProperty("errorMessage").GetString(),
                StringComparison.OrdinalIgnoreCase);
            var stillOpen = Success(await client.CallAsync(
                "range", "get-values", mainSession, RangeArgs("B2"), deadline.Token));
            Assert.Equal(999, stillOpen.GetProperty("values")[0][0].GetDouble());
            Success(await client.CallAsync("file", "close", mainSession, new(), deadline.Token));
            mainSession = SessionId(await client.CallAsync("file", "open", null, new() { ["path"] = main }, deadline.Token));
            var discarded = Success(await client.CallAsync("range", "get-values", mainSession, RangeArgs("B2"), deadline.Token));
            Assert.Equal(10, discarded.GetProperty("values")[0][0].GetDouble());
            var untouched = Success(await client.CallAsync("range", "get-values", sentinelSession, RangeArgs("A1"), deadline.Token));
            Assert.Equal(9876, untouched.GetProperty("values")[0][0].GetDouble());
            Success(await client.CallAsync("file", "close", mainSession, new(), deadline.Token));
            Success(await client.CallAsync("file", "close", sentinelSession, new(), deadline.Token));
            completed = true;
            output.WriteLine($"{entryPoint}: real open/write/formula/calculate/errors/save/reopen/discard/sentinel/close passed.");
        }
        finally
        {
            if (completed)
            {
                File.Delete(main);
                File.Delete(created);
                File.Delete(createdMacro);
                File.Delete(duplicate);
                File.Delete(sentinel);
                File.Delete(dataFile);
                directory.Delete();
                duplicateDirectory.Delete();
            }
            else
            {
                output.WriteLine($"Failed run retained synthetic fixtures at {directory.FullName}.");
            }
        }
    }

    [MacExcelTheory]
    [InlineData("cli")]
    [InlineData("mcp")]
    [Trait("Category", "Integration")]
    [Trait("RequiresExcel", "true")]
    [Trait("Feature", "MacAnalysis")]
    public async Task SpecializedAnalysis_RealEntryPointRoundTrip(string entryPoint)
    {
        Assert.Equal(0, MacAutomationAccess.Check());
        var root = FindRepository();
        var directory = Directory.CreateTempSubdirectory("excelmcp-mac-analysis-");
        var workbookPath = Path.Combine(directory.FullName, $"analysis-{Guid.NewGuid():N}.xlsx");
        CreateBlankWorkbook(workbookPath);
        using var deadline = new CancellationTokenSource(TimeSpan.FromMinutes(3));
        await using var client = await EntryPointClient.CreateAsync(root, entryPoint, output, deadline.Token);
        var completed = false;
        try
        {
            var sessionId = SessionId(await client.CallAsync("file", "open", null,
                new() { ["path"] = workbookPath }, deadline.Token));
            Success(await client.CallAsync("range", "set-values", sessionId,
                RangeArgs("F1", ("values", new object?[][] { [2] })), deadline.Token));
            Success(await client.CallAsync("range", "set-formulas", sessionId,
                RangeArgs("G1", ("formulas", new object?[][] { ["=F1*F1"] })), deadline.Token));
            var goalSeek = Success(await client.CallAsync("analysis", "goal-seek", sessionId,
                new()
                {
                    ["sheet_name"] = "Data",
                    ["formula_cell"] = "G1",
                    ["goal"] = 25,
                    ["changing_cell"] = "F1"
                }, deadline.Token));
            Assert.True(goalSeek.GetProperty("converged").GetBoolean());
            Assert.InRange(goalSeek.GetProperty("formulaValue").GetDouble(), 24.999, 25.001);
            Assert.InRange(goalSeek.GetProperty("changingValue").GetDouble(), 4.999, 5.001);

            Success(await client.CallAsync("range", "set-values", sessionId,
                RangeArgs("I1:J4", ("values", new object?[][]
                {
                    [null, null],
                    [1, null],
                    [2, null],
                    [3, null]
                })), deadline.Token));
            Success(await client.CallAsync("range", "set-formulas", sessionId,
                RangeArgs("J1", ("formulas", new object?[][] { ["=$G$1"] })), deadline.Token));
            Success(await client.CallAsync("analysis", "create-data-table", sessionId,
                new()
                {
                    ["sheet_name"] = "Data",
                    ["table_range"] = "I1:J4",
                    ["column_input_cell"] = "F1"
                }, deadline.Token));
            var dataTable = Success(await client.CallAsync("range", "get-values", sessionId,
                RangeArgs("J2:J4"), deadline.Token));
            Assert.Equal([1d, 4d, 9d], dataTable.GetProperty("values").EnumerateArray()
                .Select(row => row[0].GetDouble()).ToArray());

            Success(await client.CallAsync("file", "close", sessionId, new(), deadline.Token));
            completed = true;
            output.WriteLine($"{entryPoint}: real Goal Seek and data-table round trip passed.");
        }
        finally
        {
            if (completed)
            {
                File.Delete(workbookPath);
                directory.Delete();
            }
            else
            {
                output.WriteLine($"Failed run retained synthetic fixture at {workbookPath}.");
            }
        }
    }

    private static Dictionary<string, object?> RangeArgs(string address, params (string Key, object? Value)[] extras)
        => RangeArgsOnSheet("Data", address, extras);

    private static async Task<JsonElement> InvokeAutomationHostAsync(
        string repositoryRoot,
        string command,
        object arguments,
        CancellationToken cancellationToken)
    {
        var executable = Path.Combine(
            repositoryRoot,
            "src",
            "ExcelMcp.CLI",
            "bin",
            "Release",
            "net10.0",
            "excelcli");
        var start = new ProcessStartInfo(executable)
        {
            UseShellExecute = false,
            RedirectStandardInput = true,
            RedirectStandardOutput = true,
            RedirectStandardError = true
        };
        start.ArgumentList.Add(MacAutomationHost.Marker);
        start.ArgumentList.Add(command);
        using var process = new Process { StartInfo = start };
        Assert.True(process.Start());
        await process.StandardInput.WriteAsync(
            JsonSerializer.Serialize(arguments).AsMemory(),
            cancellationToken);
        process.StandardInput.Close();
        var stdout = process.StandardOutput.ReadToEndAsync(cancellationToken);
        var stderr = process.StandardError.ReadToEndAsync(cancellationToken);
        try
        {
            await process.WaitForExitAsync(cancellationToken);
        }
        catch (OperationCanceledException)
        {
            if (!process.HasExited)
            {
                process.Kill(entireProcessTree: true);
            }
            await process.WaitForExitAsync();
            throw;
        }
        Assert.Equal("", await stderr);
        Assert.Equal(0, process.ExitCode);
        using var result = JsonDocument.Parse(await stdout);
        return result.RootElement.Clone();
    }

    private static Dictionary<string, object?> RangeArgsOnSheet(
        string sheetName,
        string address,
        params (string Key, object? Value)[] extras)
    {
        var result = new Dictionary<string, object?> { ["sheet_name"] = sheetName, ["range_address"] = address };
        foreach (var (key, value) in extras) { result.Add(key, value); }
        return result;
    }

    private static JsonElement Success(JsonElement result)
    {
        Assert.True(result.GetProperty("success").GetBoolean(), result.GetRawText());
        if (result.TryGetProperty("errorMessage", out var error))
        {
            Assert.True(error.ValueKind == JsonValueKind.Null || string.IsNullOrEmpty(error.GetString()), result.GetRawText());
        }
        return result;
    }

    private static string SessionId(JsonElement result)
    {
        Success(result);
        var property = result.TryGetProperty("session_id", out var id) ? id : result.GetProperty("sessionId");
        Assert.False(string.IsNullOrWhiteSpace(property.GetString()));
        return property.GetString()!;
    }

    private static string FindRepository()
    {
        var directory = new DirectoryInfo(AppContext.BaseDirectory);
        while (directory is not null && !File.Exists(Path.Combine(directory.FullName, "Sbroenne.ExcelMcp.sln")))
        {
            directory = directory.Parent;
        }
        return directory?.FullName ?? throw new InvalidOperationException("Repository root not found.");
    }

    private static void CreateBlankWorkbook(string path)
    {
        var parts = new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = """<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>""",
            ["_rels/.rels"] = """<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>""",
            ["xl/workbook.xml"] = """<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Data" sheetId="1" r:id="rId1"/><sheet name="Spare" sheetId="2" r:id="rId2"/></sheets></workbook>""",
            ["xl/_rels/workbook.xml.rels"] = """<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/></Relationships>""",
            ["xl/worksheets/sheet1.xml"] = """<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData/></worksheet>""",
            ["xl/worksheets/sheet2.xml"] = """<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData/></worksheet>"""
        };
        using var archive = ZipFile.Open(path, ZipArchiveMode.Create);
        foreach (var (name, content) in parts)
        {
            using var writer = new StreamWriter(archive.CreateEntry(name).Open(), new UTF8Encoding(false));
            writer.Write(content);
        }
    }

    private sealed class EntryPointClient : IAsyncDisposable
    {
        private readonly string _cli;
        private readonly string _pipe = Environment.GetEnvironmentVariable("EXCELMCP_MAC_E2E_PIPE") ?? $"em-{Guid.NewGuid():N}";
        private readonly ITestOutputHelper _output;
        private Process? _server;
        private McpClient? _mcp;
        private Task<string>? _serverError;

        private EntryPointClient(string root, ITestOutputHelper output)
        {
            _cli = Path.Combine(root, "src", "ExcelMcp.CLI", "bin", "Release", "net10.0", "excelcli");
            _output = output;
        }

        public static async Task<EntryPointClient> CreateAsync(
            string root, string entryPoint, ITestOutputHelper output, CancellationToken cancellationToken)
        {
            var client = new EntryPointClient(root, output);
            if (entryPoint == "mcp")
            {
                var dll = Path.Combine(root, "src", "ExcelMcp.McpServer", "bin", "Release", "net10.0", "Sbroenne.ExcelMcp.McpServer.dll");
                var start = CreateStart(dll);
                start.RedirectStandardInput = true;
                client._server = Process.Start(start) ?? throw new InvalidOperationException("MCP process did not start.");
                client._serverError = client._server.StandardError.ReadToEndAsync();
                try
                {
                    client._mcp = await McpClient.CreateAsync(new StreamClientTransport(
                        client._server.StandardInput.BaseStream, client._server.StandardOutput.BaseStream),
                        cancellationToken: cancellationToken);
                }
                catch
                {
                    await client.DisposeAsync();
                    throw;
                }
            }
            return client;
        }

        public async Task<JsonElement> CallAsync(
            string tool, string action, string? sessionId, Dictionary<string, object?> args, CancellationToken cancellationToken)
        {
            string json;
            if (_mcp is not null)
            {
                args["action"] = action;
                if (sessionId is not null) { args["session_id"] = sessionId; }
                if (tool == "file" && action == "open") { args["timeout_seconds"] = 15; }
                var mcpTool = tool switch
                {
                    "sheet" => "worksheet",
                    "rangeformat" => "range_format",
                    _ => tool
                };
                var result = await _mcp.CallToolAsync(mcpTool, args, cancellationToken: cancellationToken);
                json = Assert.Single(result.Content.OfType<TextContentBlock>()).Text;
            }
            else
            {
                var command = new List<string> { "-q", tool == "file" ? "session" : tool == "calculation_mode" ? "calculationmode" : tool, action };
                if (sessionId is not null) { command.AddRange(["--session", sessionId]); }
                foreach (var (key, value) in args)
                {
                    if (key == "path") { command.Add((string)value!); continue; }
                    if (key == "save")
                    {
                        if (value is true) { command.Add("--save"); }
                        continue;
                    }
                    command.Add(key switch
                    {
                        "sheet_name" => "--sheet",
                        "range_address" => "--range",
                        _ => "--" + key.Replace('_', '-')
                    });
                    command.Add(value is string text ? text : JsonSerializer.Serialize(value));
                }
                if (tool == "file" && action == "open") { command.AddRange(["--timeout", "15"]); }
                json = await RunCliAsync(command, cancellationToken);
            }
            using var document = JsonDocument.Parse(json);
            return document.RootElement.Clone();
        }

        private static ProcessStartInfo CreateStart(string executable)
        {
            var managedAssembly = executable.EndsWith(".dll", StringComparison.OrdinalIgnoreCase);
            var start = new ProcessStartInfo(managedAssembly ? "dotnet" : executable)
            {
                UseShellExecute = false,
                RedirectStandardOutput = true,
                RedirectStandardError = true
            };
            if (managedAssembly) { start.ArgumentList.Add(executable); }
            if (!start.Environment.TryGetValue("DOTNET_ROOT", out var root) || string.IsNullOrWhiteSpace(root))
            {
                start.Environment["DOTNET_ROOT"] = Path.GetFullPath(Path.Combine(RuntimeEnvironment.GetRuntimeDirectory(), "../../.."));
            }
            return start;
        }

        private async Task<string> RunCliAsync(IEnumerable<string> arguments, CancellationToken cancellationToken)
        {
            using var commandDeadline = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken);
            commandDeadline.CancelAfter(TimeSpan.FromSeconds(45));
            cancellationToken = commandDeadline.Token;
            cancellationToken.ThrowIfCancellationRequested();
            var start = CreateStart(_cli);
            start.Environment["EXCELMCP_CLI_PIPE"] = _pipe;
            foreach (var argument in arguments) { start.ArgumentList.Add(argument); }
            using var process = Process.Start(start) ?? throw new InvalidOperationException("CLI did not start.");
            var stdout = process.StandardOutput.ReadToEndAsync(cancellationToken);
            var stderr = process.StandardError.ReadToEndAsync(cancellationToken);
            try
            {
                await process.WaitForExitAsync(cancellationToken);
                await Task.WhenAll(stdout, stderr).WaitAsync(cancellationToken);
            }
            catch (OperationCanceledException)
            {
                if (!process.HasExited) { process.Kill(entireProcessTree: true); }
                await process.WaitForExitAsync();
                throw;
            }
            var error = await stderr;
            if (!string.IsNullOrWhiteSpace(error)) { _output.WriteLine(error); }
            var result = await stdout;
            Assert.False(string.IsNullOrWhiteSpace(result), $"CLI exited {process.ExitCode.ToString(CultureInfo.InvariantCulture)} with no JSON.");
            return result;
        }

        public async ValueTask DisposeAsync()
        {
            if (_server is not null)
            {
                try
                {
                    if (_mcp is not null) { await _mcp.DisposeAsync().AsTask().WaitAsync(TimeSpan.FromSeconds(5)); }
                }
                finally
                {
                    if (!_server.HasExited) { _server.Kill(entireProcessTree: true); }
                    await _server.WaitForExitAsync();
                    if (_serverError is not null) { _output.WriteLine(await _serverError); }
                    _server.Dispose();
                }
            }
            else
            {
                using var timeout = new CancellationTokenSource(TimeSpan.FromSeconds(20));
                using var result = JsonDocument.Parse(await RunCliAsync(["-q", "service", "stop"], timeout.Token));
                Success(result.RootElement);
            }
        }
    }
}
