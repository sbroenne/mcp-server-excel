using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Commands;
using Sbroenne.ExcelMcp.Core.Models;
using Sbroenne.ExcelMcp.Core.Tests.Helpers;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Commands.PowerQuery;

/// <summary>
/// Integration tests for Power Query rename operations.
/// - success, content unchanged
/// - conflict detection (case-insensitive + trim)
/// - missing query failure
/// - invalid name failure (empty/whitespace)
/// - no-op (normalized names equal)
/// - case-only rename (Excel decides outcome)
/// </summary>
[Trait("Layer", "Core")]
[Trait("Category", "Integration")]
[Trait("RequiresExcel", "true")]
[Trait("Feature", "PowerQuery")]
[Trait("Speed", "Medium")]
public class PowerQueryCommandsRenameTests : IClassFixture<PowerQueryTestsFixture>
{
    private readonly PowerQueryCommands _commands;
    private readonly PowerQueryTestsFixture _fixture;

    public PowerQueryCommandsRenameTests(PowerQueryTestsFixture fixture)
    {
        _fixture = fixture;
        var dataModelCommands = new DataModelCommands();
        _commands = new PowerQueryCommands(dataModelCommands);
    }

    #region Success scenarios

    /// <summary>
    /// Rename an existing query to a new unique name.
    /// LLM use case: "rename query OldName to NewName"
    /// </summary>
    [Fact]
    public void Rename_UniqueNewName_ReturnsSuccess()
    {
        // Arrange
        var testFile = _fixture.CreateTestFile();
        var queryName = $"PQ_Rename_{Guid.NewGuid():N}"[..20];
        var newName = $"PQ_Renamed_{Guid.NewGuid():N}"[..20];
        var mCode = "let Source = 1 in Source";

        using var batch = ExcelSession.BeginBatch(testFile);
        _commands.Create(batch, queryName, mCode, PowerQueryLoadMode.ConnectionOnly);

        // Act
        var result = _commands.Rename(batch, queryName, newName);

        // Assert
        Assert.True(result.Success, $"Rename failed: {result.ErrorMessage}");
        Assert.Equal("power-query", result.ObjectType);
        Assert.Equal(queryName, result.OldName);
        Assert.Equal(newName, result.NewName);

        // Verify query exists under new name
        var list = _commands.List(batch);
        Assert.Contains(list.Queries, q => q.Name == newName);
        Assert.DoesNotContain(list.Queries, q => q.Name == queryName);
    }

    /// <summary>
    /// Verify M code content is unchanged after rename.
    /// </summary>
    [Fact]
    public void Rename_ContentUnchanged_AfterRename()
    {
        // Arrange
        var testFile = _fixture.CreateTestFile();
        var queryName = $"PQ_Content_{Guid.NewGuid():N}"[..20];
        var newName = $"PQ_NewContent_{Guid.NewGuid():N}"[..20];
        var mCode = "let Source = \"OriginalContent\" in Source";

        using var batch = ExcelSession.BeginBatch(testFile);
        _commands.Create(batch, queryName, mCode, PowerQueryLoadMode.ConnectionOnly);

        // Act
        var result = _commands.Rename(batch, queryName, newName);

        // Assert
        Assert.True(result.Success, $"Rename failed: {result.ErrorMessage}");

        var view = _commands.View(batch, newName);
        Assert.True(view.Success);
        Assert.Contains("OriginalContent", view.MCode);
    }

    #endregion

    #region No-op scenarios

    /// <summary>
    /// Rename where normalized names are equal returns success (no-op).
    /// </summary>
    [Fact]
    public void Rename_TrimEqual_ReturnsNoOpSuccess()
    {
        // Arrange
        var testFile = _fixture.CreateTestFile();
        var queryName = "TrimTest";
        var mCode = "let Source = 1 in Source";

        using var batch = ExcelSession.BeginBatch(testFile);
        _commands.Create(batch, queryName, mCode, PowerQueryLoadMode.ConnectionOnly);

        // Act: rename with leading/trailing spaces (should trim to same name)
        var result = _commands.Rename(batch, queryName, "  TrimTest  ");

        // Assert
        Assert.True(result.Success, $"No-op rename should succeed: {result.ErrorMessage}");
        Assert.Equal("TrimTest", result.NormalizedOldName);
        Assert.Equal("TrimTest", result.NormalizedNewName);

        // Query still exists
        var list = _commands.List(batch);
        Assert.Contains(list.Queries, q => q.Name == "TrimTest");
    }

    /// <summary>
    /// Rename with identical name is no-op success.
    /// </summary>
    [Fact]
    public void Rename_IdenticalName_ReturnsNoOpSuccess()
    {
        // Arrange
        var testFile = _fixture.CreateTestFile();
        var queryName = "IdenticalTest";
        var mCode = "let Source = 1 in Source";

        using var batch = ExcelSession.BeginBatch(testFile);
        _commands.Create(batch, queryName, mCode, PowerQueryLoadMode.ConnectionOnly);

        // Act
        var result = _commands.Rename(batch, queryName, queryName);

        // Assert
        Assert.True(result.Success, $"Identical name should be no-op success: {result.ErrorMessage}");
    }

    #endregion

    #region Case-only rename

    /// <summary>
    /// Case-only rename attempts COM rename (Excel decides outcome).
    /// </summary>
    [Fact]
    public void Rename_CaseOnlyChange_AttemptsRename()
    {
        // Arrange
        var testFile = _fixture.CreateTestFile();
        var queryName = "casetest";
        var newName = "CaseTest";
        var mCode = "let Source = 1 in Source";

        using var batch = ExcelSession.BeginBatch(testFile);
        _commands.Create(batch, queryName, mCode, PowerQueryLoadMode.ConnectionOnly);

        // Act
        var result = _commands.Rename(batch, queryName, newName);

        // Assert: Excel may accept or reject case-only rename; either way, we have a result
        // The important thing is that we attempted it (not treated as no-op)
        Assert.NotNull(result);

        if (result.Success)
        {
            // Verify new casing appears in list
            var list = _commands.List(batch);
            Assert.Contains(list.Queries, q => q.Name == newName);
        }
    }

    #endregion

    #region Error scenarios

    /// <summary>
    /// Rename to a name that conflicts (case-insensitive) with another query.
    /// </summary>
    [Fact]
    public void Rename_ConflictingName_ReturnsError()
    {
        // Arrange
        var testFile = _fixture.CreateTestFile();
        var query1 = "QueryOne";
        var query2 = "QueryTwo";
        var mCode = "let Source = 1 in Source";

        using var batch = ExcelSession.BeginBatch(testFile);
        _commands.Create(batch, query1, mCode, PowerQueryLoadMode.ConnectionOnly);
        _commands.Create(batch, query2, mCode, PowerQueryLoadMode.ConnectionOnly);

        // Act: try to rename query1 to query2 (conflict)
        var result = _commands.Rename(batch, query1, query2);

        // Assert
        Assert.False(result.Success);
        Assert.Contains("already exists", result.ErrorMessage, StringComparison.OrdinalIgnoreCase);
    }

    /// <summary>
    /// Rename to a name that conflicts case-insensitively.
    /// </summary>
    [Fact]
    public void Rename_CaseInsensitiveConflict_ReturnsError()
    {
        // Arrange
        var testFile = _fixture.CreateTestFile();
        var query1 = "QueryAlpha";
        var query2 = "QueryBeta";
        var mCode = "let Source = 1 in Source";

        using var batch = ExcelSession.BeginBatch(testFile);
        _commands.Create(batch, query1, mCode, PowerQueryLoadMode.ConnectionOnly);
        _commands.Create(batch, query2, mCode, PowerQueryLoadMode.ConnectionOnly);

        // Act: try to rename query1 to "querybeta" (case-insensitive conflict)
        var result = _commands.Rename(batch, query1, "querybeta");

        // Assert
        Assert.False(result.Success);
        Assert.Contains("already exists", result.ErrorMessage, StringComparison.OrdinalIgnoreCase);
    }

    /// <summary>
    /// Rename a query that does not exist.
    /// </summary>
    [Fact]
    public void Rename_MissingQuery_ReturnsError()
    {
        // Arrange
        var testFile = _fixture.CreateTestFile();

        using var batch = ExcelSession.BeginBatch(testFile);

        // Act
        var result = _commands.Rename(batch, "NonExistent", "NewName");

        // Assert
        Assert.False(result.Success);
        Assert.Contains("not found", result.ErrorMessage, StringComparison.OrdinalIgnoreCase);
    }

    /// <summary>
    /// Rename with empty new name.
    /// </summary>
    [Fact]
    public void Rename_EmptyNewName_ReturnsError()
    {
        // Arrange
        var testFile = _fixture.CreateTestFile();
        var queryName = "EmptyTest";
        var mCode = "let Source = 1 in Source";

        using var batch = ExcelSession.BeginBatch(testFile);
        _commands.Create(batch, queryName, mCode, PowerQueryLoadMode.ConnectionOnly);

        // Act
        var result = _commands.Rename(batch, queryName, "");

        // Assert
        Assert.False(result.Success);
        Assert.Contains("empty", result.ErrorMessage, StringComparison.OrdinalIgnoreCase);
    }

    /// <summary>
    /// Rename with whitespace-only new name.
    /// </summary>
    [Fact]
    public void Rename_WhitespaceNewName_ReturnsError()
    {
        // Arrange
        var testFile = _fixture.CreateTestFile();
        var queryName = "WhitespaceTest";
        var mCode = "let Source = 1 in Source";

        using var batch = ExcelSession.BeginBatch(testFile);
        _commands.Create(batch, queryName, mCode, PowerQueryLoadMode.ConnectionOnly);

        // Act
        var result = _commands.Rename(batch, queryName, "   ");

        // Assert
        Assert.False(result.Success);
        Assert.Contains("empty", result.ErrorMessage, StringComparison.OrdinalIgnoreCase);
    }

    #endregion
}
