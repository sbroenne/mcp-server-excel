using System.Reflection;
using Sbroenne.ExcelMcp.McpServer.Tests.Integration.Tools;
using Xunit;

namespace Sbroenne.ExcelMcp.McpServer.Tests.Unit;

[Trait("Category", "Unit")]
[Trait("Speed", "Fast")]
[Trait("Layer", "McpServer")]
[Trait("Feature", "GeneratedContracts")]
[Trait("RequiresExcel", "false")]
public sealed class GeneratedMcpParameterTests
{
    [Fact]
    public void TimeSpanAndFileInputs_AreExposedAsOptionalSecondsAndFileAliases()
    {
        var timeout = GeneratedToolContract.GetParameter("powerquery", "timeout_seconds");
        Assert.Equal(typeof(int?), timeout.ParameterType);
        Assert.True(timeout.IsOptional);
        Assert.Null(timeout.DefaultValue);
        var file = GeneratedToolContract.GetParameter("powerquery", "m_code_file");
        Assert.Equal(typeof(string), file.ParameterType);
        Assert.True(file.IsOptional);
        Assert.Null(file.DefaultValue);
        Assert.Equal(NullabilityState.Nullable, new NullabilityInfoContext().Create(file).ReadState);
    }

    [Fact]
    public void CalculationParameters_KeepActionSpecificEnumsAsOptionalStrings()
    {
        foreach (var name in new[] { "mode", "scope" })
        {
            var parameter = GeneratedToolContract.GetParameter("calculation_mode", name);
            Assert.Equal(typeof(string), parameter.ParameterType);
            Assert.True(parameter.IsOptional);
            Assert.Null(parameter.DefaultValue);
        }
    }

    [Theory]
    [InlineData("range", "values")]
    [InlineData("table", "rows")]
    public void NativeCellParameters_PreserveMixedValueCollections(string tool, string parameterName)
    {
        var parameter = GeneratedToolContract.GetParameter(tool, parameterName);
        Assert.Equal(typeof(List<List<object?>>), parameter.ParameterType);
        var nullability = new NullabilityInfoContext().Create(parameter);
        Assert.Equal(NullabilityState.Nullable, nullability.GenericTypeArguments[0].GenericTypeArguments[0].ReadState);
        Assert.True(parameter.IsOptional);
        Assert.Null(parameter.DefaultValue);
    }
}
