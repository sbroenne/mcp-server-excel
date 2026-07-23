using System.Text.Json;
using Sbroenne.ExcelMcp.CLI.Tests.Helpers;
using Xunit;

namespace Sbroenne.ExcelMcp.CLI.Tests.Integration;

[Collection("Service")]
[Trait("Category", "Integration")]
[Trait("Feature", "VBA")]
[Trait("Layer", "CLI")]
[Trait("RequiresExcel", "true")]
public sealed class VbaRunValidationCliTests : IDisposable
{
    private readonly string _testFile;

    public VbaRunValidationCliTests()
    {
        _testFile = Path.Combine(Path.GetTempPath(), $"VbaRunValidation_{Guid.NewGuid():N}.xlsm");
    }

    [Fact]
    public async Task VbaRun_WhitespaceProcedureName_IsRejectedBeforeServiceCall()
    {
        var (createResult, createJson) = await CliProcessHelper.RunJsonAsync(
            ["session", "create", _testFile],
            timeoutMs: 60000,
            diagnosticLabel: "session create for vba whitespace validation");

        Assert.Equal(0, createResult.ExitCode);
        var sessionId = createJson.RootElement.GetProperty("sessionId").GetString();
        Assert.False(string.IsNullOrWhiteSpace(sessionId));

        try
        {
            var result = await CliProcessHelper.RunAsync(
                ["vba", "run", "--session", sessionId!, "--procedure-name", "   "],
                timeoutMs: 60000,
                diagnosticLabel: "vba run with whitespace procedure name");

            Assert.Equal(1, result.ExitCode);

            using var json = JsonDocument.Parse(result.Stdout);
            Assert.False(json.RootElement.GetProperty("success").GetBoolean());
            Assert.Contains("procedureName is required for run action", json.RootElement.GetProperty("error").GetString());
        }
        finally
        {
            if (!string.IsNullOrWhiteSpace(sessionId))
            {
                try
                {
                    await CliProcessHelper.RunAsync(
                        ["session", "close", "--session", sessionId, "--save", "false"],
                        timeoutMs: 60000,
                        diagnosticLabel: "session close after vba whitespace validation");
                }
                catch
                {
                }
            }

            if (File.Exists(_testFile))
            {
                try
                {
                    File.Delete(_testFile);
                }
                catch
                {
                }
            }
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
            catch
            {
            }
        }

        GC.SuppressFinalize(this);
    }
}
