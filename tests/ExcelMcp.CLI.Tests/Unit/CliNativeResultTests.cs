using System.Text.Json;
using Sbroenne.ExcelMcp.CLI.Tests.Helpers;
using Xunit;

namespace Sbroenne.ExcelMcp.CLI.Tests.Unit;

[Trait("RequiresExcel", "false")]
[Trait("Layer", "CLI")]
[Trait("Category", "Unit")]
[Trait("Feature", "ErrorHandling")]
[Trait("Speed", "Fast")]
public sealed class CliNativeResultTests
{
    [Theory]
    [InlineData("""{"success":true,"errorMessage":null}""", true)]
    [InlineData("""{"success":true,"errorMessage":""}""", true)]
    [InlineData("""{"errorMessage":null}""", false)]
    [InlineData("""{"success":false,"errorMessage":null}""", false)]
    [InlineData("""{"success":true,"errorMessage":"failure"}""", false)]
    public void NativeAcceptance_RequiresExplicitSuccessAndNoError(string json, bool accepts)
    {
        using var response = JsonDocument.Parse(json);
        var failure = Record.Exception(() => CliNativeWorkbook.VerifyOperationResult(response.RootElement));
        Assert.Equal(accepts, failure is null);
    }
}
