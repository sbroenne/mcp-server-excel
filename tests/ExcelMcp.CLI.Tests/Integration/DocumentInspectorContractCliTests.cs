using System.Text.Json;
using Sbroenne.ExcelMcp.CLI.Tests.Helpers;
using Sbroenne.ExcelMcp.Service;
using Xunit;

namespace Sbroenne.ExcelMcp.CLI.Tests.Integration;

[Collection("Service")]
[Trait("Category", "Integration")]
[Trait("Feature", "DocumentInspector")]
[Trait("Layer", "CLI")]
[Trait("RequiresExcel", "false")]
[Trait("Speed", "Fast")]
public sealed class DocumentInspectorContractCliTests
{
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task Actions_ForwardExactInspectorSelectionAndNativeResults(bool fix)
    {
        ServiceRequest? captured = null;
        const string payload = """{"success":true,"inspectors":[{"index":3,"name":"Hidden &Worksheets","statusCode":1,"status":"issue-found","results":"One finding"}]}""";
        var arguments = new List<string>
        {
            "-q", "workbook", fix ? "fix-document-inspector" : "inspect-document", "--session", "session-1"
        };
        if (fix)
            arguments.AddRange(["--inspector-index", "3", "--inspector-name", "Hidden &Worksheets"]);
        var result = await InProcessCliHelper.RunAsync([.. arguments], request =>
        {
            captured = request;
            return new ServiceResponse { Success = true, Result = payload };
        });
        Assert.Equal(0, result.ExitCode);
        Assert.NotNull(captured);
        Assert.Equal(fix ? "workbook.fix-document-inspector" : "workbook.inspect-document", captured.Command);
        Assert.Equal("session-1", captured.SessionId);
        if (fix)
        {
            using var args = JsonDocument.Parse(captured.Args!);
            Assert.Equal(3, args.RootElement.GetProperty("inspectorIndex").GetInt32());
            Assert.Equal("Hidden &Worksheets", args.RootElement.GetProperty("inspectorName").GetString());
        }
        else
            Assert.Null(captured.Args);
        using var output = JsonDocument.Parse(result.Stdout);
        Assert.Equal("One finding", output.RootElement.GetProperty("data").GetProperty("inspectors")[0].GetProperty("results").GetString());
    }

    [Theory]
    [InlineData("--inspector-index", "3")]
    [InlineData("--inspector-name", "Hidden &Worksheets")]
    public async Task Fix_RequiresBothSelectionParametersBeforeServiceCall(string option, string value)
    {
        bool called = false;
        var result = await InProcessCliHelper.RunAsync(
            ["-q", "workbook", "fix-document-inspector", "--session", "session-1", option, value],
            _ =>
            {
                called = true;
                return new ServiceResponse { Success = true };
            });
        Assert.NotEqual(0, result.ExitCode);
        Assert.False(called);
        Assert.Contains("inspector", result.Stdout + result.Stderr, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public async Task NativeFailure_ReturnsNonzeroExitCodeAndRetainsPartialResults()
    {
        var result = await InProcessCliHelper.RunAsync(
            ["-q", "workbook", "inspect-document", "--session", "session-1"],
            _ => new ServiceResponse
            {
                Success = true,
                Result = """{"success":false,"errorMessage":"Native inspector failed","inspectors":[{"index":1,"statusCode":2,"results":"Native error"}]}"""
            });
        Assert.NotEqual(0, result.ExitCode);
        using var output = JsonDocument.Parse(result.Stdout);
        Assert.False(output.RootElement.GetProperty("success").GetBoolean());
        Assert.Equal("Native error", output.RootElement.GetProperty("data").GetProperty("inspectors")[0].GetProperty("results").GetString());
    }
}
