using System.Reflection;
using Sbroenne.ExcelMcp.Core.Commands.Chart;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Unit;

[Trait("Category", "Unit")]
[Trait("Layer", "Core")]
[Trait("Feature", "ChartDepth")]
[Trait("RequiresExcel", "false")]
public sealed class ChartAxisMappingTests
{
    [Theory]
    [InlineData(ChartAxisType.Category, 1, 1)]
    [InlineData(ChartAxisType.Value, 2, 1)]
    [InlineData(ChartAxisType.Primary, 1, 1)]
    [InlineData(ChartAxisType.Secondary, 2, 1)]
    [InlineData(ChartAxisType.CategorySecondary, 1, 2)]
    [InlineData(ChartAxisType.ValueSecondary, 2, 2)]
    public void AxisSelector_MapsToDocumentedNativeAxis(ChartAxisType axis, int type, int group)
    {
        var method = typeof(ChartCommands).GetMethod("MapAxisType", BindingFlags.NonPublic | BindingFlags.Static);
        Assert.NotNull(method);
        var actual = Assert.IsType<(int, int)>(method.Invoke(null, [axis]));
        Assert.Equal((type, group), actual);
    }
}
