using System.Text.Json;
using Sbroenne.ExcelMcp.CLI.Tests.Helpers;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Xunit;

namespace Sbroenne.ExcelMcp.CLI.Tests.Integration;

[Collection("Service")]
[Trait("Category", "Integration")]
[Trait("Feature", "Range")]
[Trait("Layer", "CLI")]
[Trait("RequiresExcel", "true")]
public sealed class RangeFormulaErrorCliParityTests : IDisposable
{
    private static readonly string[][] ReferenceErrorFormula = [["=INDIRECT(\"A0\")"]];

    private readonly string _testFile =
        Path.Join(Path.GetTempPath(), $"RangeFormulaErrorCli_{Guid.NewGuid():N}.xlsx");

    [Fact]
    public async Task RangeReads_ReturnCanonicalFormulaErrorThroughCli()
    {
        ExcelSession.CreateNew(
            _testFile,
            isMacroEnabled: false,
            (ctx, ct) => 0,
            CancellationToken.None);

        var open = await CliProcessHelper.RunAsync(["session", "open", _testFile]);
        Assert.Equal(0, open.ExitCode);
        using var openDocument = JsonDocument.Parse(open.Stdout);
        string? sessionId = openDocument.RootElement.GetProperty("sessionId").GetString();
        Assert.False(string.IsNullOrWhiteSpace(sessionId));

        try
        {
            string formulas = JsonSerializer.Serialize(ReferenceErrorFormula);
            var set = await CliProcessHelper.RunAsync(
                [
                    "range", "set-formulas",
                    "--session", sessionId!,
                    "--sheet-name", "Sheet1",
                    "--range-address", "A1",
                    "--formulas", formulas
                ]);
            Assert.Equal(0, set.ExitCode);

            var values = await CliProcessHelper.RunAsync(
                [
                    "range", "get-values",
                    "--session", sessionId!,
                    "--sheet-name", "Sheet1",
                    "--range-address", "A1"
                ]);
            var formulasResult = await CliProcessHelper.RunAsync(
                [
                    "range", "get-formulas",
                    "--session", sessionId!,
                    "--sheet-name", "Sheet1",
                    "--range-address", "A1"
                ]);

            Assert.Equal(0, values.ExitCode);
            Assert.Equal(0, formulasResult.ExitCode);
            AssertCanonicalReferenceError(values.Stdout);
            AssertCanonicalReferenceError(formulasResult.Stdout);
        }
        finally
        {
            await CliProcessHelper.RunAsync(
                ["session", "close", "--session", sessionId!, "--save", "false"]);
        }
    }

    private static void AssertCanonicalReferenceError(string json)
    {
        using var document = JsonDocument.Parse(json);
        var root = document.RootElement;
        Assert.Equal("#REF!", root.GetProperty("values")[0][0].GetString());

        var error = root.GetProperty("cellErrors")[0];
        Assert.Equal("A1", error.GetProperty("cellAddress").GetString());
        Assert.Equal("#REF!", error.GetProperty("errorName").GetString());
        Assert.Equal("=INDIRECT(\"A0\")", error.GetProperty("formula").GetString());
        Assert.Equal(-2146826265, error.GetProperty("errorCode").GetInt32());
        Assert.Equal(-2146826265, error.GetProperty("currentValue").GetInt32());
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
                // The session close path reports cleanup failures; deletion is only fixture cleanup.
            }
        }
    }
}
