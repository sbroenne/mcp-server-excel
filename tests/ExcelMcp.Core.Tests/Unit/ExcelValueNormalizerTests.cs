using Sbroenne.ExcelMcp.Core.Utilities;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Unit;

[Trait("Category", "Unit")]
[Trait("Speed", "Fast")]
[Trait("Layer", "Core")]
[Trait("Feature", "Range")]
[Trait("RequiresExcel", "false")]
public sealed class ExcelValueNormalizerTests
{
    [Fact]
    public void Normalize_Scalar_ReturnsOneByOneGrid()
    {
        var result = ExcelValueNormalizer.Normalize("Header");

        Assert.Equal(1, result.RowCount);
        Assert.Equal(1, result.ColumnCount);
        Assert.Equal([["Header"]], result.Values);
    }

    [Fact]
    public void Normalize_OneBasedArray_PreservesEveryValue()
    {
        var values = Array.CreateInstance(typeof(object), [2, 2], [1, 1]);
        values.SetValue("A", 1, 1);
        values.SetValue("B", 1, 2);
        values.SetValue(1d, 2, 1);
        values.SetValue(2d, 2, 2);

        var result = ExcelValueNormalizer.Normalize(values);

        Assert.Equal(2, result.RowCount);
        Assert.Equal(2, result.ColumnCount);
        Assert.Equal("A", result.Values[0][0]);
        Assert.Equal("B", result.Values[0][1]);
        Assert.Equal(1d, result.Values[1][0]);
        Assert.Equal(2d, result.Values[1][1]);
    }
}
