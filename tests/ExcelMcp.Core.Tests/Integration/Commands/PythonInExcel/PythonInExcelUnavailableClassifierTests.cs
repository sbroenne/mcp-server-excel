using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Commands.PythonInExcel;
using Sbroenne.ExcelMcp.Core.Commands.Range;
using Sbroenne.ExcelMcp.Core.Tests.Helpers;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Commands.PythonInExcel;

[Trait("Layer", "Core")]
[Trait("Category", "Integration")]
[Collection("Sequential")]
[Trait("RequiresExcel", "true")]
[Trait("Feature", "PythonInExcel")]
[Trait("RunType", "OnDemand")]
public sealed class PythonInExcelUnavailableClassifierTests :
    IClassFixture<PythonInExcelTestsFixture>
{
    private readonly PythonInExcelTestsFixture _fixture;

    public PythonInExcelUnavailableClassifierTests(
        PythonInExcelTestsFixture fixture)
    {
        _fixture = fixture;
    }

    [Fact]
    public void UnavailableClassifier_UsesExcelNameErrorValue()
    {
        using var batch = ExcelSession.BeginBatch(_fixture.TestFilePath);
        var startRow = _fixture.GetUniqueRowBlockStart();
        var targetCell = $"D{startRow}";
        var rangeCommands = new RangeCommands();

        var setResult = rangeCommands.SetFormulas(
            batch,
            "Sheet1",
            targetCell,
            [["=UNDEFINEDFUNCTION()"]]);
        Assert.True(setResult.Success, setResult.ErrorMessage);

        var formulaResult = rangeCommands.GetFormulas(
            batch,
            "Sheet1",
            targetCell);
        Assert.True(formulaResult.Success, formulaResult.ErrorMessage);
        object? nameErrorValue =
            Assert.Single(formulaResult.CellErrors).CurrentValue;

        Assert.True(PythonInExcelCommands.IsPythonInExcelUnavailable(
            "=PY(\"1 + 1\",0)",
            nameErrorValue,
            string.Empty));
    }
}
