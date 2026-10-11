using System.Reflection;
using Sbroenne.ExcelMcp.Core.Commands;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Unit;

[Trait("Category", "Unit")]
[Trait("Speed", "Fast")]
[Trait("Layer", "Core")]
[Trait("Feature", "Connection")]
[Trait("RequiresExcel", "false")]
public sealed class ExternalOlapSchemaErrorTests
{
    private const string Sentinel = "sentinel-provider-secret";

    [Theory]
    [InlineData(unchecked((int)0x80070005), "denied access")]
    [InlineData(unchecked((int)0x800A0CB3), "does not support")]
    [InlineData(unchecked((int)0x80004005), "could not return OLAP metadata")]
    public void ProviderError_KeepsClassificationAndHresultButOmitsProviderText(int hresult, string classification)
    {
        var method = typeof(ConnectionCommands).GetMethod(
            "CreateOlapQueryError", BindingFlags.NonPublic | BindingFlags.Static);
        Assert.NotNull(method);

        var error = Assert.IsType<InvalidOperationException>(method.Invoke(
            null, ["Cube", hresult, $"Provider failure: Password={Sentinel};Data Source=private-host"]));

        Assert.Contains(classification, error.Message, StringComparison.Ordinal);
        Assert.Contains($"0x{hresult:X8}", error.Message, StringComparison.Ordinal);
        Assert.DoesNotContain(Sentinel, error.Message, StringComparison.Ordinal);
        Assert.DoesNotContain("private-host", error.Message, StringComparison.Ordinal);
    }
}
