using System.Text.Json;
using Sbroenne.ExcelMcp.CLI.Tests.Helpers;
using Sbroenne.ExcelMcp.Service;
using Xunit;

namespace Sbroenne.ExcelMcp.CLI.Tests.Integration;

/// <summary>
/// CLI output for formula error cells. The Excel error values themselves are
/// covered by PersistentServiceRangeFormulaValidationTests.
/// </summary>
[Collection("Service")]
[Trait("Category", "Integration")]
[Trait("Feature", "Range")]
[Trait("Layer", "CLI")]
[Trait("RequiresExcel", "false")]
[Trait("Speed", "Fast")]
public sealed class RangeFormulaErrorContractCliTests
{
    private const string ResponseJson = """
        {
          "success": true,
          "values": [["#REF!"]],
          "cellErrors": [{
            "cellAddress": "A1",
            "errorName": "#REF!",
            "formula": "=INDIRECT(\"A0\")",
            "errorCode": -2146826265,
            "currentValue": -2146826265
          }]
        }
        """;

    [Theory]
    [InlineData("get-values")]
    [InlineData("get-formulas")]
    public async Task RangeRead_PreservesFormulaErrorDetails(string action)
    {
        ServiceRequest? captured = null;
        var result = await InProcessCliHelper.RunAsync(
            ["range", action, "--session", "session-1", "--sheet-name", "Sheet1", "--range-address", "A1"],
            request =>
            {
                captured = request;
                return new ServiceResponse { Success = true, Result = ResponseJson };
            });

        Assert.True(result.ExitCode == 0, result.Stdout + result.Stderr);
        Assert.NotNull(captured);
        Assert.Equal($"range.{action}", captured.Command);
        Assert.Equal("session-1", captured.SessionId);
        using (var args = JsonDocument.Parse(captured.Args!))
        {
            Assert.Equal("Sheet1", args.RootElement.GetProperty("sheetName").GetString());
            Assert.Equal("A1", args.RootElement.GetProperty("rangeAddress").GetString());
        }

        using var document = JsonDocument.Parse(result.Stdout);
        var root = document.RootElement;
        Assert.True(root.GetProperty("success").GetBoolean());
        Assert.Equal("#REF!", root.GetProperty("data").GetProperty("values")[0][0].GetString());
        var error = Assert.Single(root.GetProperty("data").GetProperty("cellErrors").EnumerateArray());
        Assert.Equal("A1", error.GetProperty("cellAddress").GetString());
        Assert.Equal("#REF!", error.GetProperty("errorName").GetString());
        Assert.Equal("=INDIRECT(\"A0\")", error.GetProperty("formula").GetString());
        Assert.Equal(-2146826265, error.GetProperty("errorCode").GetInt32());
        Assert.Equal(-2146826265, error.GetProperty("currentValue").GetInt32());
    }
}
