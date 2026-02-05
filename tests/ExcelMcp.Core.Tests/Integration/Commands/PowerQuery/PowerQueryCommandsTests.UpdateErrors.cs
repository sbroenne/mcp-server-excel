using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Commands;
using Sbroenne.ExcelMcp.Core.Models;
using Sbroenne.ExcelMcp.Core.Tests.Helpers;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Commands.PowerQuery;

/// <summary>
/// Tests for Power Query Update error handling.
/// Validates that COM errors are caught and translated to meaningful messages.
///
/// DISCOVERY: Excel accepts invalid M code during formula assignment (query.Formula = mCode).
/// The error only occurs during REFRESH when the engine actually parses the M code.
/// This means the error handling in Update catches errors during the post-update refresh,
/// not during the formula assignment itself.
/// </summary>
[Trait("Layer", "Core")]
[Trait("Category", "Integration")]
[Trait("RequiresExcel", "true")]
[Trait("Feature", "PowerQuery")]
[Trait("Speed", "Medium")]
public partial class PowerQueryCommandsTests
{
    /// <summary>
    /// Verifies that Excel accepts invalid M code during formula assignment.
    /// The M code validation happens lazily during refresh, not during assignment.
    /// This is expected Excel behavior.
    /// </summary>
    [Fact]
    public void Update_InvalidMCodeSyntax_AcceptedByExcel()
    {
        // Arrange - Create a query first
        var testExcelFile = _fixture.CreateTestFile();
        var queryName = "TestQuery";
        var validMCode = @"let
    Source = #table({""A""}, {{1}})
in
    Source";

        // This M code has syntax errors but Excel will accept it during assignment
        var invalidMCode = @"let
    Source = this is not valid M code syntax!!!
in
    Source";

        using var batch = ExcelSession.BeginBatch(testExcelFile);

        // Create valid query first
        _powerQueryCommands.Create(batch, queryName, validMCode, PowerQueryLoadMode.ConnectionOnly);
        batch.Save();

        // Act - Excel accepts the invalid M code (validation is lazy)
        // No exception should be thrown - ConnectionOnly queries don't refresh
        _powerQueryCommands.Update(batch, queryName, invalidMCode);

        // Assert - The formula was updated (even though invalid)
        var result = _powerQueryCommands.View(batch, queryName);
        Assert.True(result.Success, result.ErrorMessage);
        Assert.Contains("not valid", result.MCode);
    }

    /// <summary>
    /// Verifies that Update with valid M code succeeds (regression test).
    /// </summary>
    [Fact]
    public void Update_ValidMCode_Succeeds()
    {
        // Arrange
        var testExcelFile = _fixture.CreateTestFile();
        var queryName = "TestQuery3";
        var initialMCode = @"let
    Source = #table({""A""}, {{1}})
in
    Source";

        var updatedMCode = @"let
    Source = #table({""A"", ""B""}, {{1, 2}})
in
    Source";

        using var batch = ExcelSession.BeginBatch(testExcelFile);

        // Create query
        _powerQueryCommands.Create(batch, queryName, initialMCode, PowerQueryLoadMode.ConnectionOnly);
        batch.Save();

        // Act - Update with valid M code should succeed
        _powerQueryCommands.Update(batch, queryName, updatedMCode);

        // Assert - View should show updated M code
        var result = _powerQueryCommands.View(batch, queryName);
        Assert.True(result.Success, result.ErrorMessage);
        Assert.Contains("B", result.MCode); // New column added
    }

    /// <summary>
    /// Verifies that Update on non-existent query throws InvalidOperationException.
    /// </summary>
    [Fact]
    public void Update_NonExistentQuery_ThrowsWithMeaningfulMessage()
    {
        // Arrange
        var testExcelFile = _fixture.CreateTestFile();
        var queryName = "NonExistentQuery";
        var mCode = @"let Source = 1 in Source";

        using var batch = ExcelSession.BeginBatch(testExcelFile);

        // Act & Assert
        var exception = Assert.Throws<InvalidOperationException>(() =>
            _powerQueryCommands.Update(batch, queryName, mCode));

        // Verify error message is meaningful
        Assert.Contains(queryName, exception.Message);
    }
}




