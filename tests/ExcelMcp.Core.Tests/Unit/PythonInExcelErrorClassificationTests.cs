using Sbroenne.ExcelMcp.Core.Commands.PythonInExcel;
using Sbroenne.ExcelMcp.Core.Models;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Unit;

/// <summary>
/// Pure classification tests for Excel error values returned by an already-settled PY cell.
/// </summary>
[Trait("Layer", "Core")]
[Trait("Category", "Unit")]
[Trait("Feature", "PythonInExcel")]
[Trait("RequiresExcel", "false")]
[Trait("Speed", "Fast")]
public sealed class PythonInExcelErrorClassificationTests
{
    [Fact]
    public void TryApplyPythonUnavailable_ExactNameError_ReturnsStableDiagnostic()
    {
        var result = new PythonInExcelResult
        {
            Success = true,
            IsPythonError = false,
            IsPythonObject = false
        };

        bool classified = PythonInExcelCommands.TryApplyPythonUnavailable(result, -2146826252);

        Assert.True(classified);
        Assert.False(result.Success);
        Assert.True(result.IsPythonUnavailable);
        Assert.False(result.IsPythonError);
        Assert.False(result.IsPythonObject);
        Assert.Equal("Python in Excel unavailable", result.ErrorMessage);
    }

    [Theory]
    [InlineData(-2146826237)] // #BUSY!
    [InlineData(-2146826288)] // #NULL!
    [InlineData(-2146826273)] // #VALUE!
    [InlineData(null)]
    public void TryApplyPythonUnavailable_OtherValues_RemainUnclassified(int? value)
    {
        var result = new PythonInExcelResult { Success = true };

        bool classified = PythonInExcelCommands.TryApplyPythonUnavailable(result, value);

        Assert.False(classified);
        Assert.True(result.Success);
        Assert.False(result.IsPythonUnavailable);
        Assert.Null(result.ErrorMessage);
    }
}
