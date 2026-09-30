using Sbroenne.ExcelMcp.Core.Commands.Chart;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Unit;

[Trait("Category", "Unit")]
[Trait("Feature", "Chart")]
[Trait("Layer", "Core")]
[Trait("RequiresExcel", "false")]
[Trait("Speed", "Fast")]
public sealed class ChartPositionGuidanceTests
{
    [Theory]
    [InlineData(1)]
    [InlineData(3)]
    public void ClearLayout_DoesNotRequireScreenshot(int chartCount)
    {
        var message = ChartPositionHelpers.FormatCollisionWarnings([], chartCount);

        Assert.DoesNotContain("MUST", message, StringComparison.Ordinal);
        Assert.Contains("interactive desktop", message, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("OVERLAP WARNING", message, StringComparison.Ordinal);
    }

    [Fact]
    public void Overlap_PreservesWarningAndOffersNonVisualVerification()
    {
        var message = ChartPositionHelpers.FormatCollisionWarnings(["Chart overlaps data area A1:D20"]);

        Assert.Contains("OVERLAP WARNING: Chart overlaps data area A1:D20", message);
        Assert.Contains("fit-to-range", message);
        Assert.Contains("chart read", message);
        Assert.Contains("interactive desktop", message);
    }

    [Fact]
    public void PivotChartAddSeries_IdentifiesFieldTool()
    {
        var error = Assert.Throws<NotSupportedException>(() =>
            new PivotChartStrategy().AddSeries(new object(), "Sales", "A1:B2", null));

        Assert.Contains("pivottable_field", error.Message);
        Assert.Contains("add-value-field", error.Message);
    }

    [Fact]
    public void PivotChartRemoveSeries_IdentifiesFieldTool()
    {
        var error = Assert.Throws<NotSupportedException>(() =>
            new PivotChartStrategy().RemoveSeries(new object(), 1));

        Assert.Contains("pivottable_field", error.Message);
        Assert.Contains("remove-field", error.Message);
    }
}
