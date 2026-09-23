using Sbroenne.ExcelMcp.Core.Commands.PythonInExcel;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Unit;

/// <summary>
/// Regression tests for issue #753: Excel reports an unavailable Python in Excel feature as
/// a worksheet #NAME? error instead of throwing a COM exception.
/// </summary>
[Trait("Layer", "Core")]
[Trait("Category", "Unit")]
[Trait("Feature", "PythonInExcel")]
[Trait("Speed", "Fast")]
[Trait("RequiresExcel", "false")]
public sealed class PythonInExcelAvailabilityClassificationTests
{
    [Theory]
    [InlineData("=PY(\"1 + 1\",0)")]
    [InlineData("=_xlfn.PY(\"1 + 1\",0)")]
    [InlineData("  =py(\"1 + 1\",0)  ")]
    public void IsPythonInExcelUnavailable_NameErrorOnTopLevelPythonFormula_ReturnsTrue(string formula)
    {
        Assert.True(PythonInExcelCommands.IsPythonInExcelUnavailable(formula, -2146826259, "#NAME?"));
    }

    [Fact]
    public void IsPythonInExcelUnavailable_NameErrorCodeWithoutRenderedText_ReturnsTrue()
    {
        Assert.True(PythonInExcelCommands.IsPythonInExcelUnavailable(
            "=PY(\"1 + 1\",0)",
            -2146826259,
            string.Empty));
    }

    [Fact]
    public void IsPythonInExcelUnavailable_DoubleNameErrorCode_ReturnsTrue()
    {
        Assert.True(PythonInExcelCommands.IsPythonInExcelUnavailable(
            "=PY(\"1 + 1\",0)",
            (double)-2146826259,
            string.Empty));
    }

    [Theory]
    [InlineData("#BUSY!")]
    [InlineData("#CONNECT!")]
    [InlineData("#BLOCKED!")]
    [InlineData("#PYTHON!")]
    public void IsPythonInExcelUnavailable_NonNameMarker_ReturnsFalse(string displayedText)
    {
        Assert.False(PythonInExcelCommands.IsPythonInExcelUnavailable(
            "=PY(\"1 + 1\",0)",
            null,
            displayedText));
    }

    [Theory]
    [InlineData("=SUM(A1:A2)")]
    [InlineData("=PYTHON(\"1 + 1\")")]
    [InlineData("=SUM(PY(\"1 + 1\",0))")]
    [InlineData("")]
    public void IsPythonInExcelUnavailable_NameErrorOnNonTopLevelPythonFormula_ReturnsFalse(string formula)
    {
        Assert.False(PythonInExcelCommands.IsPythonInExcelUnavailable(formula, -2146826259, "#NAME?"));
    }

    [Theory]
    [InlineData("#NAME")]
    [InlineData("#NAME? ")]
    [InlineData("")]
    public void IsPythonInExcelUnavailable_NonCanonicalNameText_ReturnsFalse(string displayedText)
    {
        Assert.False(PythonInExcelCommands.IsPythonInExcelUnavailable(
            "=PY(\"1 + 1\",0)",
            null,
            displayedText));
    }
}
