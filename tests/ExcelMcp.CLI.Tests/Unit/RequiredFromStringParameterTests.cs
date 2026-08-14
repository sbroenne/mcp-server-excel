using Sbroenne.ExcelMcp.Generated;
using Xunit;

namespace Sbroenne.ExcelMcp.CLI.Tests.Unit;

/// <summary>
/// Verifies generated CLI routing rejects missing required string-backed enum parameters.
/// </summary>
[Trait("Layer", "CLI")]
[Trait("Category", "Unit")]
[Trait("Feature", "ActionValidation")]
[Trait("Speed", "Fast")]
public sealed class RequiredFromStringParameterTests
{
    [Theory]
    [InlineData(null)]
    [InlineData("")]
    [InlineData(" ")]
    public void SheetStyleGroup_MissingAxis_ThrowsBeforeDispatch(string? axis)
    {
        var exception = Assert.Throws<ArgumentException>(() =>
            ServiceRegistry.SheetStyle.RouteCliArgs(
                "group",
                sheetName: "Sheet1",
                rangeAddress: "2:5",
                axis: axis));

        Assert.Contains("axis", exception.Message, StringComparison.OrdinalIgnoreCase);
    }
}
