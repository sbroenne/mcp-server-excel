using System.Reflection;
using System.Runtime.CompilerServices;
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
        var actual = Assert.IsAssignableFrom<ITuple>(method.Invoke(null, [axis]));
        Assert.Equal(2, actual.Length);
        Assert.Equal(type, Convert.ToInt32(actual[0], System.Globalization.CultureInfo.InvariantCulture));
        Assert.Equal(group, Convert.ToInt32(actual[1], System.Globalization.CultureInfo.InvariantCulture));
    }
}
