using System.Text.Json;
using Sbroenne.ExcelMcp.CLI.Tests.Helpers;
using Xunit;
using Xunit.Abstractions;

namespace Sbroenne.ExcelMcp.CLI.Tests.Integration;

[Collection("Service")]
[Trait("Category", "Integration")]
[Trait("Feature", "VBA")]
[Trait("Layer", "CLI")]
[Trait("RequiresExcel", "true")]
public sealed class VbaRunCliTransportProofTests : IDisposable
{
    private const string ModuleName = "TransportProof";
    private const string ProcedureName = "TransportProof.WriteTransportProof";
    private const string MarkerValue = "cli-vba-run-ok";

    private readonly ITestOutputHelper _output;
    private readonly string _tempDir;
    private readonly string _macroWorkbook;
    private readonly string _moduleFile;

    public VbaRunCliTransportProofTests(ITestOutputHelper output)
    {
        _output = output;
        _tempDir = Path.Combine(Path.GetTempPath(), $"CliVbaRunProof_{Guid.NewGuid():N}");
        Directory.CreateDirectory(_tempDir);

        _macroWorkbook = Path.Combine(_tempDir, "VbaRunProof.xlsm");
        _moduleFile = Path.Combine(_tempDir, "TransportProof.bas");
    }

    [Fact]
    public async Task VbaRun_OnMacroWorkbook_ViaCli_ExecutesAndPersistsWorkbookSideEffect()
    {
        await File.WriteAllTextAsync(
            _moduleFile,
            """
            Sub WriteTransportProof()
                ThisWorkbook.Sheets(1).Range("A1").Value = "cli-vba-run-ok"
            End Sub
            """);

        string? sessionId = await CreateSessionAsync(_macroWorkbook, "cli-vba-create-session");
        var saveOnClose = false;

        try
        {
            await AssertCommandSuccessAsync(
                ["vba", "import", "--session", sessionId, "--module-name", ModuleName, "--vba-code-file", _moduleFile],
                timeoutMs: 60000,
                diagnosticLabel: "cli-vba-import-module");

            await AssertCommandSuccessAsync(
                ["vba", "run", "--session", sessionId, "--procedure-name", ProcedureName],
                timeoutMs: 60000,
                diagnosticLabel: "cli-vba-run-procedure");

            var inSessionValue = await ReadCellValueAsync(sessionId, "cli-vba-read-in-session");
            Assert.Equal(MarkerValue, inSessionValue);

            saveOnClose = true;
        }
        finally
        {
            if (!string.IsNullOrWhiteSpace(sessionId))
            {
                await CloseSessionAsync(sessionId, saveOnClose, "cli-vba-close-created-session");
            }
        }

        var reopenedSessionId = await OpenSessionAsync(_macroWorkbook, "cli-vba-reopen-session");
        try
        {
            var persistedValue = await ReadCellValueAsync(reopenedSessionId, "cli-vba-read-after-reopen");
            Assert.Equal(MarkerValue, persistedValue);
        }
        finally
        {
            await CloseSessionAsync(reopenedSessionId, save: false, "cli-vba-close-reopened-session");
        }
    }

    private async Task<string> CreateSessionAsync(string workbookPath, string diagnosticLabel)
    {
        var (result, json) = await CliProcessHelper.RunJsonAsync(
            ["session", "create", workbookPath],
            timeoutMs: 60000,
            diagnosticLabel: diagnosticLabel);

        _output.WriteLine($"[{diagnosticLabel}] Stdout: {result.Stdout}");
        _output.WriteLine($"[{diagnosticLabel}] Stderr: {result.Stderr}");

        Assert.True(result.ExitCode == 0, $"{diagnosticLabel} failed. Stdout: {result.Stdout} Stderr: {result.Stderr}");
        Assert.True(json.RootElement.GetProperty("success").GetBoolean(), $"{diagnosticLabel} returned success=false. Stdout: {result.Stdout}");

        return json.RootElement.GetProperty("sessionId").GetString()
            ?? throw new InvalidOperationException($"{diagnosticLabel} did not return a sessionId.");
    }

    private async Task<string> OpenSessionAsync(string workbookPath, string diagnosticLabel)
    {
        var result = await CliProcessHelper.RunAsync(
            ["session", "open", workbookPath],
            timeoutMs: 60000,
            diagnosticLabel: diagnosticLabel);

        _output.WriteLine($"[{diagnosticLabel}] Stdout: {result.Stdout}");
        _output.WriteLine($"[{diagnosticLabel}] Stderr: {result.Stderr}");

        Assert.True(result.ExitCode == 0, $"{diagnosticLabel} failed. Stdout: {result.Stdout} Stderr: {result.Stderr}");

        using var json = JsonDocument.Parse(result.Stdout);
        Assert.True(json.RootElement.GetProperty("success").GetBoolean(), $"{diagnosticLabel} returned success=false. Stdout: {result.Stdout}");

        return json.RootElement.GetProperty("sessionId").GetString()
            ?? throw new InvalidOperationException($"{diagnosticLabel} did not return a sessionId.");
    }

    private async Task CloseSessionAsync(string sessionId, bool save, string diagnosticLabel)
    {
        var (result, json) = await CliProcessHelper.RunJsonAsync(
            ["session", "close", "--session", sessionId, "--save", save ? "true" : "false"],
            timeoutMs: 60000,
            diagnosticLabel: diagnosticLabel);

        _output.WriteLine($"[{diagnosticLabel}] Stdout: {result.Stdout}");
        _output.WriteLine($"[{diagnosticLabel}] Stderr: {result.Stderr}");

        Assert.True(result.ExitCode == 0, $"{diagnosticLabel} failed. Stdout: {result.Stdout} Stderr: {result.Stderr}");
        Assert.True(json.RootElement.GetProperty("success").GetBoolean(), $"{diagnosticLabel} returned success=false. Stdout: {result.Stdout}");
    }

    private async Task AssertCommandSuccessAsync(IReadOnlyList<string> args, int timeoutMs, string diagnosticLabel)
    {
        var (result, json) = await CliProcessHelper.RunJsonAsync(
            args,
            timeoutMs: timeoutMs,
            diagnosticLabel: diagnosticLabel);

        _output.WriteLine($"[{diagnosticLabel}] Stdout: {result.Stdout}");
        _output.WriteLine($"[{diagnosticLabel}] Stderr: {result.Stderr}");

        Assert.True(result.ExitCode == 0, $"{diagnosticLabel} failed. Stdout: {result.Stdout} Stderr: {result.Stderr}");
        Assert.True(json.RootElement.GetProperty("success").GetBoolean(), $"{diagnosticLabel} returned success=false. Stdout: {result.Stdout}");
    }

    private async Task<string?> ReadCellValueAsync(string sessionId, string diagnosticLabel)
    {
        var (result, json) = await CliProcessHelper.RunJsonAsync(
            ["range", "get-values", "--session", sessionId, "--sheet-name", "Sheet1", "--range-address", "A1"],
            timeoutMs: 60000,
            diagnosticLabel: diagnosticLabel);

        _output.WriteLine($"[{diagnosticLabel}] Stdout: {result.Stdout}");
        _output.WriteLine($"[{diagnosticLabel}] Stderr: {result.Stderr}");

        Assert.True(result.ExitCode == 0, $"{diagnosticLabel} failed. Stdout: {result.Stdout} Stderr: {result.Stderr}");
        Assert.True(json.RootElement.GetProperty("success").GetBoolean(), $"{diagnosticLabel} returned success=false. Stdout: {result.Stdout}");

        return json.RootElement.GetProperty("values")[0][0].GetString();
    }

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
            }
        }

        GC.SuppressFinalize(this);
    }
}
