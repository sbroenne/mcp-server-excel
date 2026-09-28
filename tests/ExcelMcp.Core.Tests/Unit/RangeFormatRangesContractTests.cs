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
    public void FormatRanges_PublicInterfaceAndImplementation_KeepExpectedSignature()
    {
        var parameterTypes = new[]
        {
            typeof(IExcelBatch),
            typeof(string),
            typeof(string[]),
            typeof(string),
            typeof(double?),
            typeof(bool?),
            typeof(bool?),
            typeof(bool?),
            typeof(string),
            typeof(string),
            typeof(string),
            typeof(string),
            typeof(string),
            typeof(string),
            typeof(string),
            typeof(bool?),
            typeof(int?),
            typeof(string)
        };
        var expectedNames = new[]
        {
            "batch", "sheetName", "rangeAddresses", "fontName", "fontSize",
            "bold", "italic", "underline", "fontColor", "fillColor",
            "borderStyle", "borderColor", "borderWeight",
            "horizontalAlignment", "verticalAlignment", "wrapText",
            "orientation", "numberFormat"
        };

        var interfaceMethod = typeof(IRangeFormatCommands)
            .GetMethod("FormatRanges", parameterTypes);
        Assert.NotNull(interfaceMethod);
        Assert.Equal(
            expectedNames,
            interfaceMethod.GetParameters().Select(parameter => parameter.Name));

        var implementationMethod = typeof(RangeCommands)
            .GetMethod("FormatRanges", parameterTypes);
        Assert.NotNull(implementationMethod);
        Assert.Equal(typeof(OperationResult), implementationMethod.ReturnType);
    }
}
