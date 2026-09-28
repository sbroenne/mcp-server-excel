using Sbroenne.ExcelMcp.Core.Commands;
using Sbroenne.ExcelMcp.Core.Models;
using Xunit;

namespace Sbroenne.ExcelMcp.Service.Tests;

[Collection("ServiceWorkflow")]
[Trait("Category", "Integration")]
[Trait("Layer", "Service")]
[Trait("Feature", "PowerQuery")]
[Trait("RequiresExcel", "true")]
public sealed class PersistentServicePowerQueryRenameTests(
    PersistentServiceWorkbookFixture fixture) :
    PersistentServiceWorkbookTestBase(fixture),
    IClassFixture<PersistentServiceWorkbookFixture>
{
    private const string MCode = "let Source = 1 in Source";
    private readonly IPowerQueryCommands _queries =
        ServiceCommandProxy.Create<IPowerQueryCommands>(fixture);

    [Fact]
    public void Rename_UniqueNewName_ReturnsSuccess()
    {
        var queryName = $"PQ_Rename_{Guid.NewGuid():N}"[..20];
        var newName = $"PQ_Renamed_{Guid.NewGuid():N}"[..20];
        var batch = _fixture.BatchToken;
        _queries.Create(batch, queryName, MCode, PowerQueryLoadMode.ConnectionOnly);

        var result = _queries.Rename(batch, queryName, newName);

        Assert.True(result.Success, $"Rename failed: {result.ErrorMessage}");
        Assert.Equal("power-query", result.ObjectType);
        Assert.Equal(queryName, result.OldName);
        Assert.Equal(newName, result.NewName);
        var list = _queries.List(batch);
        Assert.Contains(list.Queries, query => query.Name == newName);
        Assert.DoesNotContain(list.Queries, query => query.Name == queryName);
    }

    [Fact]
    public void Rename_ContentUnchanged_AfterRename()
    {
        var queryName = $"PQ_Content_{Guid.NewGuid():N}"[..20];
        var newName = $"PQ_NewContent_{Guid.NewGuid():N}"[..20];
        const string mCode = "let Source = \"OriginalContent\" in Source";
        var batch = _fixture.BatchToken;
        _queries.Create(batch, queryName, mCode, PowerQueryLoadMode.ConnectionOnly);

        var result = _queries.Rename(batch, queryName, newName);

        Assert.True(result.Success, $"Rename failed: {result.ErrorMessage}");
        var view = _queries.View(batch, newName);
        Assert.True(view.Success);
        Assert.Contains("OriginalContent", view.MCode);
    }

    [Fact]
    public void Rename_TrimEqual_ReturnsNoOpSuccess()
    {
        var queryName = $"TrimTest_{Guid.NewGuid():N}"[..20];
        var batch = _fixture.BatchToken;
        _queries.Create(batch, queryName, MCode, PowerQueryLoadMode.ConnectionOnly);

        var result = _queries.Rename(batch, queryName, $"  {queryName}  ");

        Assert.True(result.Success, $"No-op rename should succeed: {result.ErrorMessage}");
        Assert.Equal(queryName, result.NormalizedOldName);
        Assert.Equal(queryName, result.NormalizedNewName);
        var list = _queries.List(batch);
        Assert.Contains(list.Queries, query => query.Name == queryName);
    }

    [Fact]
    public void Rename_IdenticalName_ReturnsNoOpSuccess()
    {
        var queryName = $"Identical_{Guid.NewGuid():N}"[..20];
        var batch = _fixture.BatchToken;
        _queries.Create(batch, queryName, MCode, PowerQueryLoadMode.ConnectionOnly);

        var result = _queries.Rename(batch, queryName, queryName);

        Assert.True(result.Success, $"Identical name should be no-op success: {result.ErrorMessage}");
    }

    [Fact]
    public void Rename_CaseOnlyChange_AttemptsRename()
    {
        var suffix = Guid.NewGuid().ToString("N")[..8];
        var queryName = $"case{suffix}";
        var newName = $"Case{suffix}";
        var batch = _fixture.BatchToken;
        _queries.Create(batch, queryName, MCode, PowerQueryLoadMode.ConnectionOnly);

        var result = _queries.Rename(batch, queryName, newName);

        Assert.NotNull(result);
        if (result.Success)
        {
            var list = _queries.List(batch);
            Assert.Contains(list.Queries, query => query.Name == newName);
        }
    }

    [Fact]
    public void Rename_ConflictingName_ReturnsError()
    {
        var suffix = Guid.NewGuid().ToString("N")[..8];
        var query1 = $"QueryOne{suffix}";
        var query2 = $"QueryTwo{suffix}";
        var batch = _fixture.BatchToken;
        _queries.Create(batch, query1, MCode, PowerQueryLoadMode.ConnectionOnly);
        _queries.Create(batch, query2, MCode, PowerQueryLoadMode.ConnectionOnly);

        var result = _queries.Rename(batch, query1, query2);

        Assert.False(result.Success);
        Assert.Contains("already exists", result.ErrorMessage, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Rename_CaseInsensitiveConflict_ReturnsError()
    {
        var suffix = Guid.NewGuid().ToString("N")[..8];
        var query1 = $"QueryAlpha{suffix}";
        var query2 = $"QueryBeta{suffix}";
        var batch = _fixture.BatchToken;
        _queries.Create(batch, query1, MCode, PowerQueryLoadMode.ConnectionOnly);
        _queries.Create(batch, query2, MCode, PowerQueryLoadMode.ConnectionOnly);

        var result = _queries.Rename(batch, query1, query2.ToLowerInvariant());

        Assert.False(result.Success);
        Assert.Contains("already exists", result.ErrorMessage, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Rename_MissingQuery_ReturnsError()
    {
        var result = _queries.Rename(
            _fixture.BatchToken,
            $"Missing{Guid.NewGuid():N}",
            $"New{Guid.NewGuid():N}");

        Assert.False(result.Success);
        Assert.Contains("not found", result.ErrorMessage, StringComparison.OrdinalIgnoreCase);
    }

    [Theory]
    [InlineData("")]
    [InlineData("   ")]
    public void Rename_EmptyNewName_ThrowsHelpfulPublicError(string newName)
    {
        var queryName = $"PQ_Empty_{Guid.NewGuid():N}"[..20];
        var batch = _fixture.BatchToken;
        _queries.Create(batch, queryName, MCode, PowerQueryLoadMode.ConnectionOnly);
        _fixture.RegisterPowerQueryForCleanup(queryName);

        var exception = Assert.Throws<ArgumentException>(() =>
            _queries.Rename(batch, queryName, newName));

        Assert.Contains(
            "newName is required",
            exception.Message,
            StringComparison.OrdinalIgnoreCase);
    }

}
