using Sbroenne.ExcelMcp.Core.Commands.PythonInExcel;
using Sbroenne.ExcelMcp.Core.Models;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Unit;

[Trait("Layer", "Core")]
[Trait("Category", "Unit")]
[Trait("Feature", "PythonInExcel")]
[Trait("Speed", "Fast")]
[Trait("RequiresExcel", "false")]
public sealed class PythonInExcelErrorClassificationTests
{
    [Fact]
    public void PopulateErrorResult_PythonObjectModeWithDisplayedFormulaError_ReturnsFailure()
    {
        var result = new PythonInExcelResult();

        PythonInExcelCommands.PopulateErrorResult(
            result,
            errorCode: -2146826265,
            formulaReturnType: 1,
            displayedText: "#REF!");

        Assert.False(result.Success);
        Assert.False(result.IsPythonObject);
        Assert.False(result.IsPythonError);
        Assert.StartsWith("#REF!", result.ErrorMessage, StringComparison.Ordinal);
    }

    [Fact]
    public void PopulateErrorResult_PythonObjectModeWithSharedErrorCode_PreservesObject()
    {
        var result = new PythonInExcelResult();

        PythonInExcelCommands.PopulateErrorResult(
            result,
            errorCode: -2146826273,
            formulaReturnType: 1,
            displayedText: "DataFrame");

        Assert.True(result.Success);
        Assert.True(result.IsPythonObject);
        Assert.False(result.IsPythonError);
        Assert.Equal("DataFrame", result.TypeName);
        Assert.True(string.IsNullOrEmpty(result.ErrorMessage));
    }
}
