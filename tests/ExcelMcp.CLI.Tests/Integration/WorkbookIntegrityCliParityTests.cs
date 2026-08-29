using System.Text.Json;
using Sbroenne.ExcelMcp.CLI.Tests.Helpers;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Xunit;

namespace Sbroenne.ExcelMcp.CLI.Tests.Integration;

[Collection("Service")]
[Trait("Category", "Integration")]
[Trait("Feature", "Workbook")]
[Trait("Layer", "CLI")]
[Trait("RequiresExcel", "true")]
public sealed class WorkbookIntegrityCliParityTests : IDisposable
{
    private static readonly string[][] ReferenceErrorFormula = [["=INDIRECT(\"A0\")"]];

    private readonly string _testFile =
        Path.Join(Path.GetTempPath(), $"WorkbookIntegrityCli_{Guid.NewGuid():N}.xlsx");

    [Fact]
    public async Task ValidateIntegrity_ReturnsTypedCanonicalFindingThroughCli()
    {
        ExcelSession.CreateNew(
            _testFile,
            isMacroEnabled: false,
            (context, cancellationToken) => 0,
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

            var validate = await CliProcessHelper.RunAsync(
                [
                    "workbook", "validate-integrity",
                    "--session", sessionId!,
                    "--checks", """["formula-errors"]""",
                    "--worksheet-names", """["Sheet1"]""",
                    "--max-findings", "10"
                ]);

            Assert.Equal(0, validate.ExitCode);
            AssertIntegrityReferenceError(validate.Stdout);

            var setValue = await CliProcessHelper.RunAsync(
                [
                    "range", "set-values",
                    "--session", sessionId!,
                    "--sheet-name", "Sheet1",
                    "--range-address", "B1",
                    "--values", "[[100]]"
                ]);
            Assert.Equal(0, setValue.ExitCode);

            var controlTotal = await CliProcessHelper.RunAsync(
                [
                    "workbook", "validate-integrity",
                    "--session", sessionId!,
                    "--checks", """["control-totals"]""",
                    "--control-totals", """[{"sheetName":"Sheet1","cellAddress":"B1","expectedValue":100,"tolerance":0}]"""
                ]);
            Assert.Equal(0, controlTotal.ExitCode);
            using var controlDocument = JsonDocument.Parse(controlTotal.Stdout);
            Assert.Equal("passed", controlDocument.RootElement.GetProperty("overallStatus").GetString());
            Assert.Equal(
                "control-totals",
                controlDocument.RootElement.GetProperty("checkedChecks")[0].GetString());
        }
        finally
        {
            await CliProcessHelper.RunAsync(
                ["session", "close", "--session", sessionId!, "--save", "false"]);
        }
    }

    private static void AssertIntegrityReferenceError(string json)
    {
        using var document = JsonDocument.Parse(json);
        var root = document.RootElement;
        Assert.Equal("failed", root.GetProperty("overallStatus").GetString());
        Assert.Equal(1, root.GetProperty("findingCount").GetInt32());
        Assert.False(root.GetProperty("findingsTruncated").GetBoolean());
        Assert.Equal("formula-errors", root.GetProperty("checkedChecks")[0].GetString());

        var group = Assert.Single(root.GetProperty("groups").EnumerateArray());
        Assert.Equal("error", group.GetProperty("severity").GetString());
        Assert.Equal("broken-reference", group.GetProperty("category").GetString());

        var finding = Assert.Single(group.GetProperty("findings").EnumerateArray());
        Assert.Equal("broken-formula-reference", finding.GetProperty("code").GetString());
        Assert.Equal("deterministic", finding.GetProperty("reliability").GetString());
        Assert.Equal("Sheet1", finding.GetProperty("sheetName").GetString());
        Assert.Equal("A1", finding.GetProperty("cellAddress").GetString());
        Assert.Equal("=INDIRECT(\"A0\")", finding.GetProperty("formula").GetString());
        Assert.Equal("#REF!", finding.GetProperty("errorName").GetString());
        Assert.Equal(-2146826265, finding.GetProperty("errorCode").GetInt32());
    }

    public void Dispose()
    {
        if (!File.Exists(_testFile))
        {
            return;
        }

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
