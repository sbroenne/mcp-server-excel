using Sbroenne.ExcelMcp.Core.Utilities;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Unit;

[Trait("Layer", "Core")]
[Trait("Category", "Unit")]
[Trait("Feature", "Range")]
[Trait("Speed", "Fast")]
public sealed class ExcelErrorMapperTests
{
    [Theory]
    [InlineData(-2146826288, "#NULL!", true)]
    [InlineData(-2146826281, "#DIV/0!", true)]
    [InlineData(-2146826273, "#VALUE!", true)]
    [InlineData(-2146826265, "#REF!", true)]
    [InlineData(-2146826259, "#NAME?", true)]
    [InlineData(-2146826252, "#NUM!", true)]
    [InlineData(-2146826246, "#N/A", true)]
    [InlineData(-2146826245, "#GETTING_DATA", false)]
    [InlineData(-2146826243, "#SPILL!", true)]
    [InlineData(-2146826242, "#CONNECT!", false)]
    [InlineData(-2146826241, "#BLOCKED!", false)]
    [InlineData(-2146826240, "#UNKNOWN!", true)]
    [InlineData(-2146826239, "#FIELD!", true)]
    [InlineData(-2146826238, "#CALC!", true)]
    [InlineData(ExcelErrorMapper.BusyErrorCode, "#BUSY!", false)]
    [InlineData(ExcelErrorMapper.PythonErrorCode, "#PYTHON!", false)]
    public void TryGet_KnownComError_ReturnsCanonicalMapping(
        int errorCode,
        string expectedName,
        bool expectedFormulaError)
    {
        bool found = ExcelErrorMapper.TryGet(errorCode, out var error);

        Assert.True(found);
        Assert.Equal(expectedName, error.Name);
        Assert.Equal(expectedFormulaError, error.IsExcelFormulaError);
        Assert.False(string.IsNullOrWhiteSpace(error.Description));
        Assert.False(string.IsNullOrWhiteSpace(error.Suggestion));
    }

    [Fact]
    public void TryGet_UnknownNegativeValue_DoesNotInventFormulaError()
    {
        Assert.False(ExcelErrorMapper.TryGet(-1, out _));
        Assert.False(ExcelErrorMapper.IsExcelFormulaError(-1));
    }

    [Fact]
    public void TryGet_DoubleComVariant_ReturnsCanonicalMapping()
    {
        bool found = ExcelErrorMapper.TryGet(
            (object)(double)-2146826265,
            out int errorCode,
            out var error);

        Assert.True(found);
        Assert.Equal(-2146826265, errorCode);
        Assert.Equal("#REF!", error.Name);
    }

    [Theory]
    [InlineData(-2146826265.5)]
    [InlineData(double.PositiveInfinity)]
    public void TryGetErrorCode_InvalidDouble_DoesNotConvert(double value)
    {
        Assert.False(ExcelErrorMapper.TryGetErrorCode(value, out _));
    }
}
