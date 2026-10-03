using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Commands.Range;
using Sbroenne.ExcelMcp.Core.Models;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Unit;

[Trait("Category", "Unit")]
[Trait("Layer", "Core")]
[Trait("Feature", "Range")]
[Trait("RequiresExcel", "false")]
public sealed class RangeFormatRangesContractTests
{
    [Fact]
    public void Format_PublicInterfaceAndImplementation_UseOneTypedRequestWithoutObsoleteActions()
    {
        var parameterTypes = new[]
        {
            typeof(IExcelBatch),
            typeof(string),
            typeof(string[]),
            typeof(CellFormatOptions)
        };
        var expectedNames = new[]
        {
            "batch", "sheetName", "rangeAddresses", "formatOptions"
        };

        var interfaceMethod = typeof(IRangeFormatCommands)
            .GetMethod("Format", parameterTypes);
        Assert.NotNull(interfaceMethod);
        Assert.Equal(
            expectedNames,
            interfaceMethod.GetParameters().Select(parameter => parameter.Name));

        var implementationMethod = typeof(RangeCommands)
            .GetMethod("Format", parameterTypes);
        Assert.NotNull(implementationMethod);
        Assert.Equal(typeof(OperationResult), implementationMethod.ReturnType);
        Assert.Null(typeof(IRangeFormatCommands).GetMethod("FormatRange"));
        Assert.Null(typeof(IRangeFormatCommands).GetMethod("FormatRanges"));
        Assert.Null(typeof(RangeCommands).GetMethod("FormatRange"));
        Assert.Null(typeof(RangeCommands).GetMethod("FormatRanges"));
    }
}
