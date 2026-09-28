using Sbroenne.ExcelMcp.Core.Models;
using Sbroenne.ExcelMcp.Core.Commands;
using Xunit;

namespace Sbroenne.ExcelMcp.Service.Tests;

[Trait("Category", "Unit")]
[Trait("RequiresExcel", "false")]
public sealed class ServiceCommandProxyTests
{
    [Fact]
    public void DeserializeResult_ObjectIntegerWithinInt32Range_PreservesInt32()
    {
        const string json =
            """{"success":true,"values":[[-2146826288]],"rowCount":1,"columnCount":1}""";

        var result = ServiceCommandProxy.DeserializeResult<RangeValueResult>(json);

        Assert.IsType<int>(result.Values[0][0]);
        Assert.Equal(-2146826288, result.Values[0][0]);
    }

    [Fact]
    public void DeserializeResult_ObjectIsoDateString_PreservesString()
    {
        const string json =
            """{"success":true,"values":[["2025-01-15"]],"rowCount":1,"columnCount":1}""";

        var result = ServiceCommandProxy.DeserializeResult<RangeValueResult>(json);

        Assert.Equal("2025-01-15", Assert.IsType<string>(result.Values[0][0]));
    }

    [Fact]
    public void GetActionName_UnannotatedContractMethod_UsesGeneratorConvention()
    {
        var method = typeof(IPowerQueryCommands).GetMethod(
            nameof(IPowerQueryCommands.Create))!;

        Assert.Equal("create", ServiceCommandProxy.GetActionName(method));
    }

    [Fact]
    public void GetParameterName_FromStringOverride_UsesExposedContractName()
    {
        var parameter = typeof(IPowerQueryCommands).GetMethod(
            nameof(IPowerQueryCommands.Create))!
            .GetParameters()
            .Single(candidate => candidate.Name == "loadMode");

        Assert.Equal(
            "loadDestination",
            ServiceCommandProxy.GetParameterName(parameter, 3));
    }

    [Fact]
    public void IsAmbientProgressParameter_ExcludesGeneratedProgressContext()
    {
        var parameters = typeof(IPowerQueryCommands)
            .GetMethod(nameof(IPowerQueryCommands.Refresh))!
            .GetParameters();

        Assert.False(
            ServiceCommandProxy.IsAmbientProgressParameter(parameters[1]));
        Assert.True(
            ServiceCommandProxy.IsAmbientProgressParameter(parameters[^1]));
    }

    [Fact]
    public void NormalizeRequestValue_ConvertsTimeoutToPublicWholeSeconds()
    {
        var timeout = typeof(IPowerQueryCommands)
            .GetMethod(nameof(IPowerQueryCommands.Refresh))!
            .GetParameters()
            .Single(parameter => parameter.Name == "timeout");

        var result = ServiceCommandProxy.NormalizeRequestValue(
            timeout,
            TimeSpan.FromMinutes(5));

        Assert.Equal(300, Assert.IsType<int>(result));
    }

}
