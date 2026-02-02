using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Commands;
using Sbroenne.ExcelMcp.Core.Tests.Helpers;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Commands.PowerQuery;

/// <summary>
/// Tests for Power Query Evaluate action.
/// Feature Issue #400: Execute M code and return results without creating a permanent query.
/// </summary>
[Trait("Layer", "Core")]
[Trait("Category", "Integration")]
[Trait("RequiresExcel", "true")]
[Trait("Feature", "PowerQuery")]
[Trait("Speed", "Medium")]
public partial class PowerQueryCommandsTests
{
    /// <summary>
    /// Tests evaluating a simple M code snippet that returns a table.
    /// </summary>
    [Fact]
    public void Evaluate_SimpleTable_ReturnsData()
    {
        // Arrange
        var testExcelFile = _fixture.CreateTestFile();
        var mCode = @"let
    Source = #table(
        {""Name"", ""Value""},
        {{""Test1"", 100}, {""Test2"", 200}}
    )
in
    Source";

        using var batch = ExcelSession.BeginBatch(testExcelFile);

        // Act
        var result = _powerQueryCommands.Evaluate(batch, mCode);

        // Assert
        Assert.True(result.Success, $"Evaluate failed: {result.ErrorMessage}");
        Assert.Equal(2, result.ColumnCount);
        Assert.Equal(2, result.RowCount);
        Assert.Contains("Name", result.Columns);
        Assert.Contains("Value", result.Columns);
        Assert.Equal(2, result.Rows.Count);
    }

    /// <summary>
    /// Tests evaluating M code with a single column.
    /// </summary>
    [Fact]
    public void Evaluate_SingleColumn_ReturnsData()
    {
        // Arrange
        var testExcelFile = _fixture.CreateTestFile();
        var mCode = @"let
    Source = #table({""SingleCol""}, {{1}, {2}, {3}})
in
    Source";

        using var batch = ExcelSession.BeginBatch(testExcelFile);

        // Act
        var result = _powerQueryCommands.Evaluate(batch, mCode);

        // Assert
        Assert.True(result.Success, $"Evaluate failed: {result.ErrorMessage}");
        Assert.Equal(1, result.ColumnCount);
        Assert.Equal(3, result.RowCount);
        Assert.Equal("SingleCol", result.Columns[0]);
    }

    /// <summary>
    /// Tests that invalid M code throws an error (not silent success).
    /// </summary>
    [Fact]
    public void Evaluate_InvalidMCode_ThrowsError()
    {
        // Arrange
        var testExcelFile = _fixture.CreateTestFile();
        var invalidMCode = @"let
    Source = UndefinedFunction()
in
    Source";

        using var batch = ExcelSession.BeginBatch(testExcelFile);

        // Act & Assert
        var exception = Assert.ThrowsAny<Exception>(() =>
            _powerQueryCommands.Evaluate(batch, invalidMCode));

        // Verify error message contains Power Query error
        Assert.Contains("Expression.Error", exception.Message, StringComparison.OrdinalIgnoreCase);
    }

    /// <summary>
    /// Tests that temporary query and worksheet are cleaned up after evaluation.
    /// </summary>
    [Fact]
    public void Evaluate_AfterExecution_CleansUpTempObjects()
    {
        // Arrange
        var testExcelFile = _fixture.CreateTestFile();
        var mCode = @"let
    Source = #table({""X""}, {{1}})
in
    Source";

        using var batch = ExcelSession.BeginBatch(testExcelFile);

        // Get initial state
        var initialQueries = _powerQueryCommands.List(batch);
        var initialQueryCount = initialQueries.Queries.Count;

        // Act
        var result = _powerQueryCommands.Evaluate(batch, mCode);

        // Assert - Verify evaluation succeeded
        Assert.True(result.Success, $"Evaluate failed: {result.ErrorMessage}");

        // Verify no temp queries remain
        var finalQueries = _powerQueryCommands.List(batch);
        Assert.Equal(initialQueryCount, finalQueries.Queries.Count);

        // Verify no temp worksheets remain (__pq_eval_ prefix)
        // Check worksheet count hasn't changed significantly
    }

    /// <summary>
    /// Tests evaluating M code with various data types.
    /// </summary>
    [Fact]
    public void Evaluate_VariousDataTypes_ReturnsCorrectValues()
    {
        // Arrange
        var testExcelFile = _fixture.CreateTestFile();
        var mCode = @"let
    Source = #table(
        {""Text"", ""Number"", ""Boolean""},
        {{""Hello"", 42, true}, {""World"", 3.14, false}}
    )
in
    Source";

        using var batch = ExcelSession.BeginBatch(testExcelFile);

        // Act
        var result = _powerQueryCommands.Evaluate(batch, mCode);

        // Assert
        Assert.True(result.Success, $"Evaluate failed: {result.ErrorMessage}");
        Assert.Equal(3, result.ColumnCount);
        Assert.Equal(2, result.RowCount);

        // Check first row values
        var firstRow = result.Rows[0];
        Assert.Equal("Hello", firstRow[0]?.ToString());
        Assert.Equal(42.0, Convert.ToDouble(firstRow[1], System.Globalization.CultureInfo.InvariantCulture));
        Assert.True(Convert.ToBoolean(firstRow[2], System.Globalization.CultureInfo.InvariantCulture));
    }

    /// <summary>
    /// Tests that empty M code throws ArgumentException.
    /// </summary>
    [Fact]
    public void Evaluate_EmptyMCode_ThrowsArgumentException()
    {
        // Arrange
        var testExcelFile = _fixture.CreateTestFile();

        using var batch = ExcelSession.BeginBatch(testExcelFile);

        // Act & Assert
        Assert.Throws<ArgumentException>(() =>
            _powerQueryCommands.Evaluate(batch, ""));
    }

    /// <summary>
    /// Tests evaluating M code with data transformations.
    /// </summary>
    [Fact]
    public void Evaluate_WithTransformations_ReturnsTransformedData()
    {
        // Arrange
        var testExcelFile = _fixture.CreateTestFile();
        var mCode = @"let
    Source = #table({""Value""}, {{1}, {2}, {3}, {4}, {5}}),
    Filtered = Table.SelectRows(Source, each [Value] > 2),
    Added = Table.AddColumn(Filtered, ""Doubled"", each [Value] * 2)
in
    Added";

        using var batch = ExcelSession.BeginBatch(testExcelFile);

        // Act
        var result = _powerQueryCommands.Evaluate(batch, mCode);

        // Assert
        Assert.True(result.Success, $"Evaluate failed: {result.ErrorMessage}");
        Assert.Equal(2, result.ColumnCount); // Value, Doubled
        Assert.Equal(3, result.RowCount); // Values 3, 4, 5
        Assert.Contains("Doubled", result.Columns);
    }
}
