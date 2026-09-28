using Sbroenne.ExcelMcp.Core.Commands.Drawing;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Integration.Commands.Drawing;

[Trait("Layer", "Core")]
[Trait("Category", "Unit")]
[Trait("Feature", "Drawing")]
[Trait("RequiresExcel", "false")]
public sealed class DrawingCommandsTests
{
    [Fact]
    public void AddFormControl_ExposesOnlyFormsControls_NotActiveX()
    {
        var supportedControls = Enum.GetNames<DrawingFormControlType>();

        Assert.Contains(nameof(DrawingFormControlType.Button), supportedControls);
        Assert.Contains(
            nameof(DrawingFormControlType.CheckBox), supportedControls);
        Assert.Contains(
            nameof(DrawingFormControlType.DropDown), supportedControls);
        Assert.DoesNotContain(
            supportedControls,
            name => name.Contains("ActiveX", StringComparison.OrdinalIgnoreCase));
        Assert.DoesNotContain(
            supportedControls,
            name => name.Contains("Ole", StringComparison.OrdinalIgnoreCase));
    }
}
