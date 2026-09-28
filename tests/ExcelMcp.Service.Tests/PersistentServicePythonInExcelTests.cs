using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Commands.PythonInExcel;
using Sbroenne.ExcelMcp.Core.Commands.Range;
using Sbroenne.ExcelMcp.Core.Models;
using Xunit;

namespace Sbroenne.ExcelMcp.Service.Tests;

internal interface IPersistentPythonInExcelCommands :
    IPythonInExcelCommands,
    IRangeCommands;

/// <summary>
/// Integration tests for Python in Excel (=PY()) Core operations.
/// These tests require Excel installation, a Microsoft 365 subscription with the
/// Python in Excel feature enabled, and network access to the Microsoft-hosted
/// Python cloud sandbox. They are intentionally excluded from default CI runs via
/// [Trait("RunType", "OnDemand")] since results depend on a real cloud round-trip
/// (typically a few seconds, occasionally longer under cold-start conditions).
/// Tests use Core commands directly (not through CLI wrapper).
/// </summary>
[Trait("Layer", "Service")]
[Trait("Category", "Integration")]
[Collection("ServiceWorkflow")]
[Trait("RequiresExcel", "true")]
[Trait("Feature", "PythonInExcel")]
[Trait("RunType", "OnDemand")]
public class PersistentServicePythonInExcelTests :
    PersistentServiceWorkbookTestBase,
    IClassFixture<PersistentServiceWorkbookFixture>
{
    // Every test opens its own Excel process via ExcelSession.BeginBatch(), so every test hits a
    // fresh cold start of the Microsoft-hosted Python cloud sandbox (no session/process reuse
    // across test methods). A generous wait accounts for this cold-start latency.
    private const int DefaultMaxWaitSeconds = 90;

    private readonly IPersistentPythonInExcelCommands _pythonCommands;
    private readonly IRangeCommands _rangeCommands;

    /// <summary>
    /// Initializes a new instance of the <see cref="PersistentServicePythonInExcelTests"/> class.
    /// </summary>
    public PersistentServicePythonInExcelTests(
        PersistentServiceWorkbookFixture fixture) :
        base(fixture)
    {
        _pythonCommands =
            fixture.CreateCommands<IPersistentPythonInExcelCommands>();
        _rangeCommands = _pythonCommands;
    }

    /// <summary>
    /// Calls GetResult repeatedly until <paramref name="isAcceptable"/> is satisfied or the retry
    /// budget is exhausted. Completion detection is now deterministic (CalculationState + a per-cell
    /// #BUSY! guard - see PythonInExcelCommands.GetResult), so a single call normally converges; this
    /// wrapper only adds defensive headroom for extreme cold-start delays that exceed the deadline.
    /// The dedicated <see cref="SetFormula_ThenSingleGetResult_ConvergesWithoutRetry"/> test asserts
    /// the no-retry path directly.
    /// </summary>
    private PythonInExcelResult GetResultWithRetry(
        IExcelBatch batch,
        string sheetName,
        string cell,
        Func<PythonInExcelResult, bool> isAcceptable,
        int attempts = 3)
    {
        PythonInExcelResult result = _pythonCommands.GetResult(batch, sheetName, cell, DefaultMaxWaitSeconds);
        for (int i = 1; i < attempts && !isAcceptable(result); i++)
        {
            Thread.Sleep(3000);
            result = _pythonCommands.GetResult(batch, sheetName, cell, DefaultMaxWaitSeconds);
        }

        return result;
    }

    /// <summary>
    /// Writes a simple numeric column (5, 15, 25, 35, 45) starting at the given row, and
    /// returns the sheet-qualified source range address (e.g. "A10:A14").
    /// </summary>
    private string WriteSourceData(
        IExcelBatch batch,
        string sheetName,
        int startRow)
    {
        var values = new List<List<object?>>
        {
            new() { 5 },
            new() { 15 },
            new() { 25 },
            new() { 35 },
            new() { 45 },
        };
        string sourceRange = $"A{startRow}:A{startRow + 4}";
        var setResult = _rangeCommands.SetValues(
            batch,
            sheetName,
            sourceRange,
            values);
        Assert.True(setResult.Success, setResult.ErrorMessage);
        return sourceRange;
    }

    /// <summary>
    /// Regression test for the completion-detection flakiness fix: a single GetResult call (no retry
    /// loop) must reliably return the converged cloud value rather than a stale/#BUSY! placeholder.
    /// Before the fix, GetResult used a value-stability heuristic that could return before the cloud
    /// result arrived; it now waits deterministically on Application.CalculationState + a per-cell
    /// #BUSY! guard, so one call is enough.
    /// </summary>
    [Fact]
    public void SetFormula_ThenSingleGetResult_ConvergesWithoutRetry()
    {
        // Arrange
        var batch = _fixture.BatchToken;
        var sheetName = _fixture.CreateTestSheet(batch);
        const int startRow = 1;
        string sourceRange = WriteSourceData(batch, sheetName, startRow); // 5,15,25,35,45 => sum 125
        string targetCell = $"D{startRow}";
        string code = $"xl('{sourceRange}').sum()";

        var setResult = _pythonCommands.SetFormula(batch, sheetName, targetCell, code, returnType: 0);
        Assert.True(setResult.Success, setResult.ErrorMessage);

        // Act - a SINGLE GetResult call, deliberately NOT wrapped in GetResultWithRetry.
        var getResult = _pythonCommands.GetResult(batch, sheetName, targetCell, DefaultMaxWaitSeconds);

        // Assert - must be the real converged result, never a #BUSY! placeholder or premature read.
        Assert.True(getResult.Success, getResult.ErrorMessage);
        Assert.False(getResult.IsPythonError);
        Assert.False(getResult.IsPythonObject);
        Assert.NotNull(getResult.Value);
        Assert.Equal(125d, Convert.ToDouble(getResult.Value, System.Globalization.CultureInfo.InvariantCulture));
    }

    /// <summary>
    /// Preserves support for writing one Python formula across a multi-cell target range while the
    /// immediate availability check samples only the top-left cell.
    /// </summary>
    [Fact]
    public void SetFormula_OnMultiCellRange_WritesEveryCell()
    {
        var batch = _fixture.BatchToken;
        var sheetName = _fixture.CreateTestSheet(batch);
        const int startRow = 1;
        string targetRange = $"D{startRow}:D{startRow + 1}";

        var setResult = _pythonCommands.SetFormula(
            batch,
            sheetName,
            targetRange,
            "1 + 1",
            returnType: 0);

        Assert.True(setResult.Success, setResult.ErrorMessage);

        var formulaResult = _rangeCommands.GetFormulas(batch, sheetName, targetRange);
        Assert.True(formulaResult.Success, formulaResult.ErrorMessage);
        Assert.All(formulaResult.Formulas, row => Assert.StartsWith("=PY(", row[0], StringComparison.Ordinal));
    }

    /// <inheritdoc/>
    [Fact]
    public void SetFormula_ThenGetResult_ComputesMeanOfWorksheetRange()
    {
        // Arrange
        var batch = _fixture.BatchToken;
        var sheetName = _fixture.CreateTestSheet(batch);
        const int startRow = 1;
        string sourceRange = WriteSourceData(batch, sheetName, startRow);
        string targetCell = $"D{startRow}";
        string code = $"xl('{sourceRange}').mean()";

        // Act
        var setResult = _pythonCommands.SetFormula(batch, sheetName, targetCell, code, returnType: 0);
        Assert.True(setResult.Success, setResult.ErrorMessage);

        var getResult = GetResultWithRetry(
            batch, sheetName, targetCell,
            r => r.Success && !r.IsPythonError && Convert.ToDouble(r.Value, System.Globalization.CultureInfo.InvariantCulture) == 25d);

        // Assert
        Assert.True(getResult.Success, getResult.ErrorMessage);
        Assert.False(getResult.IsPythonError);
        Assert.False(getResult.IsPythonObject);
        Assert.NotNull(getResult.Value);
        Assert.Equal(25d, Convert.ToDouble(getResult.Value, System.Globalization.CultureInfo.InvariantCulture));
    }

    /// <inheritdoc/>
    [Fact]
    public void SetFormula_ThenGetResult_ComputesMaxOfWorksheetRange()
    {
        // Arrange
        var batch = _fixture.BatchToken;
        var sheetName = _fixture.CreateTestSheet(batch);
        const int startRow = 1;
        string sourceRange = WriteSourceData(batch, sheetName, startRow);
        string targetCell = $"D{startRow}";
        string code = $"xl('{sourceRange}').max()";

        // Act
        _pythonCommands.SetFormula(batch, sheetName, targetCell, code, returnType: 0);
        var getResult = GetResultWithRetry(
            batch, sheetName, targetCell,
            r => r.Success && !r.IsPythonError && Convert.ToDouble(r.Value, System.Globalization.CultureInfo.InvariantCulture) == 45d);

        // Assert
        Assert.True(getResult.Success, getResult.ErrorMessage);
        Assert.Equal(45d, Convert.ToDouble(getResult.Value, System.Globalization.CultureInfo.InvariantCulture));
    }

    /// <inheritdoc/>
    [Fact]
    public void SetFormula_WithSyntaxError_GetResultReturnsPythonError()
    {
        // Arrange
        var batch = _fixture.BatchToken;
        var sheetName = _fixture.CreateTestSheet(batch);
        const int startRow = 1;
        string targetCell = $"D{startRow}";

        // Act - deliberately invalid Python syntax
        _pythonCommands.SetFormula(batch, sheetName, targetCell, "this is not valid python (((", returnType: 0);
        var getResult = GetResultWithRetry(batch, sheetName, targetCell, r => r.IsPythonError);

        // Assert
        Assert.False(getResult.Success);
        Assert.True(getResult.IsPythonError);
        Assert.NotNull(getResult.ErrorMessage);
    }

    /// <inheritdoc/>
    [Fact]
    public void SetFormula_WithReturnTypePythonObject_GetResultReportsObjectType()
    {
        // Arrange
        var batch = _fixture.BatchToken;
        var sheetName = _fixture.CreateTestSheet(batch);
        const int startRow = 1;
        string sourceRange = WriteSourceData(batch, sheetName, startRow);
        string targetCell = $"D{startRow}";
        // pandas.Series is a Python Object when returnType=1 (Python Object mode)
        string code = $"xl('{sourceRange}').squeeze()";

        // Act
        _pythonCommands.SetFormula(batch, sheetName, targetCell, code, returnType: 1);
        var getResult = GetResultWithRetry(batch, sheetName, targetCell, r => r.Success && r.IsPythonObject);

        // Assert
        Assert.True(getResult.Success, getResult.ErrorMessage);
        Assert.True(getResult.IsPythonObject);
        Assert.False(string.IsNullOrEmpty(getResult.TypeName));
    }

    /// <inheritdoc/>
    [Fact]
    public void GetResult_OnCellWithoutPyFormula_ReturnsFailure()
    {
        // Arrange
        var batch = _fixture.BatchToken;
        var sheetName = _fixture.CreateTestSheet(batch);
        const int startRow = 1;
        string targetCell = $"D{startRow}";

        // Act - target cell has no PY() formula at all
        var getResult = _pythonCommands.GetResult(batch, sheetName, targetCell, DefaultMaxWaitSeconds);

        // Assert
        Assert.False(getResult.Success);
        Assert.Contains("does not contain", getResult.ErrorMessage, StringComparison.OrdinalIgnoreCase);
    }

    /// <inheritdoc/>
    [Fact]
    public void SetFormula_WithEmptyCode_ThrowsArgumentException()
    {
        // Arrange
        var batch = _fixture.BatchToken;
        var sheetName = _fixture.CreateTestSheet(batch);
        const int startRow = 1;
        string targetCell = $"D{startRow}";

        // Act & Assert
        Assert.Throws<ArgumentException>(
            () => _pythonCommands.SetFormula(batch, sheetName, targetCell, string.Empty));
    }
}
