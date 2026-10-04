using System.Text.Json;
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
    private const string MCode =
        "let Source = #table(type table [Value = Int64.Type], {{17}, {29}}) in Source";
    private const string GuardMCode =
        "let Source = #table(type table [Value = Int64.Type], {{47}, {83}}) in Source";
    private readonly IPowerQueryCommands _queries =
        fixture.CreateCommands<IPowerQueryCommands>();

    [Fact]
    public void Rename_UniqueNewName_ReturnsSuccess()
    {
        var name = CreateTracked();
        var guard = CreateGuard();
        var newName = UniqueName("Renamed");
        var result = RequireSuccess(_queries.Rename(_fixture.BatchToken, name, newName));
        _fixture.ForgetPowerQuery(name);
        _fixture.RegisterPowerQueryForCleanup(newName);
        AssertRenameMetadata(result, name, newName);
        PowerQueryStateAssertions.AssertRemoved(_fixture, name);
        AssertStored(newName);
        AssertGuard(guard);
    }

    [Fact]
    public void Rename_ContentUnchanged_AfterRename()
    {
        const string code =
            "let Source = #table({\"Text\"}, {{\"OriginalContent\"}, {\"SecondRecord\"}}) in Source";
        var name = CreateTracked(code);
        var newName = UniqueName("Content");
        var result = RequireSuccess(_queries.Rename(_fixture.BatchToken, name, newName));
        _fixture.ForgetPowerQuery(name);
        _fixture.RegisterPowerQueryForCleanup(newName);
        AssertRenameMetadata(result, name, newName);
        PowerQueryStateAssertions.AssertRemoved(_fixture, name);
        PowerQueryStateAssertions.AssertStored(_fixture, newName, code,
            PowerQueryLoadMode.ConnectionOnly, null, ["Text"], [["OriginalContent"], ["SecondRecord"]]);
        var evaluated = RequireSuccess(_queries.Evaluate(_fixture.BatchToken, code));
        Assert.Equal(["Text"], evaluated.Columns);
        Assert.Equal(2, evaluated.RowCount);
        Assert.Equal(1, evaluated.ColumnCount);
        PowerQueryStateAssertions.AssertRows([["OriginalContent"], ["SecondRecord"]], evaluated.Rows);
        Assert.Equal(code, RequireSuccess(_queries.View(_fixture.BatchToken, newName)).MCode);
    }

    [Fact]
    public void Rename_TrimEqual_ReturnsNoOpSuccess() => AssertNoOp(trim: true);

    [Fact]
    public void Rename_IdenticalName_ReturnsNoOpSuccess() => AssertNoOp(trim: false);

    [Fact]
    public void Rename_CaseOnlyChange_AttemptsRename()
    {
        var name = CreateTracked(name: "case" + Guid.NewGuid().ToString("N")[..8]);
        var newName = "Case" + name[4..];
        var beforeCount = RequireSuccess(_queries.List(_fixture.BatchToken)).Queries.Count;
        var result = RequireSuccess(_queries.Rename(_fixture.BatchToken, name, newName));
        _fixture.ForgetPowerQuery(name);
        _fixture.RegisterPowerQueryForCleanup(newName);
        AssertRenameMetadata(result, name, newName);
        var listed = RequireSuccess(_queries.List(_fixture.BatchToken)).Queries;
        Assert.Equal(beforeCount, listed.Count);
        Assert.DoesNotContain(listed, query => query.Name == name);
        AssertStored(newName);
    }

    [Fact]
    public void Rename_ConflictingName_ReturnsError() => AssertConflict(ignoreCase: false);

    [Fact]
    public void Rename_CaseInsensitiveConflict_ReturnsError() => AssertConflict(ignoreCase: true);

    [Fact]
    public void Rename_MissingQuery_ReturnsError()
    {
        var name = CreateTracked();
        var guard = CreateGuard();
        var missing = UniqueName("Missing");
        var newName = UniqueName("New");
        var before = SnapshotQueries();
        var result = _queries.Rename(_fixture.BatchToken, missing, newName);
        Assert.False(result.Success);
        AssertRenameMetadata(result, missing, newName);
        Assert.Contains("not found", result.ErrorMessage, StringComparison.OrdinalIgnoreCase);
        Assert.Contains(missing, result.ErrorMessage);
        Assert.Equal(before, SnapshotQueries());
        AssertStored(name);
        AssertGuard(guard);
    }

    [Theory]
    [InlineData("")]
    [InlineData("   ")]
    public void Rename_EmptyNewName_ThrowsHelpfulPublicError(string newName)
    {
        var name = CreateTracked();
        var guard = CreateGuard();
        var before = SnapshotQueries();
        var exception = Assert.Throws<ArgumentException>(() =>
            _queries.Rename(_fixture.BatchToken, name, newName));
        Assert.Contains("newName is required", exception.Message, StringComparison.OrdinalIgnoreCase);
        Assert.Equal(before, SnapshotQueries());
        AssertStored(name);
        AssertGuard(guard);
    }

    private void AssertNoOp(bool trim)
    {
        var name = CreateTracked();
        var guard = CreateGuard();
        var newName = trim ? $"  {name}  " : name;
        var before = SnapshotQueries();
        var result = RequireSuccess(_queries.Rename(_fixture.BatchToken, name, newName));
        AssertRenameMetadata(result, name, newName);
        Assert.Equal(before, SnapshotQueries());
        AssertStored(name);
        AssertGuard(guard);
    }

    private void AssertConflict(bool ignoreCase)
    {
        var name = CreateTracked();
        var guard = CreateGuard();
        var newName = ignoreCase ? guard.ToLowerInvariant() : guard;
        var before = SnapshotQueries();
        var result = _queries.Rename(_fixture.BatchToken, name, newName);
        Assert.False(result.Success);
        AssertRenameMetadata(result, name, newName);
        Assert.Contains("already exists", result.ErrorMessage, StringComparison.OrdinalIgnoreCase);
        Assert.Equal(before, SnapshotQueries());
        AssertStored(name);
        AssertGuard(guard);
    }

    private string CreateTracked(string code = MCode, string? name = null)
    {
        name ??= UniqueName("Rename");
        RequireSuccess(_queries.Create(_fixture.BatchToken, name, code, PowerQueryLoadMode.ConnectionOnly));
        _fixture.RegisterPowerQueryForCleanup(name);
        PowerQueryStateAssertions.AssertStored(_fixture, name, code,
            PowerQueryLoadMode.ConnectionOnly, null, [], []);
        return name;
    }

    private string CreateGuard()
    {
        var name = UniqueName("RenameGuard");
        RequireSuccess(_queries.Create(_fixture.BatchToken, name, GuardMCode,
            PowerQueryLoadMode.LoadToBoth, name));
        _fixture.RegisterPowerQueryForCleanup(name);
        _fixture.RegisterSheetForCleanup(name);
        AssertGuard(name);
        return name;
    }

    private void AssertStored(string name) =>
        PowerQueryStateAssertions.AssertStored(_fixture, name, MCode,
            PowerQueryLoadMode.ConnectionOnly, null, ["Value"], [[17], [29]]);

    private void AssertGuard(string name) =>
        PowerQueryStateAssertions.AssertStored(_fixture, name, GuardMCode,
            PowerQueryLoadMode.LoadToBoth, name, ["Value"], [[47], [83]]);

    private string SnapshotQueries() =>
        JsonSerializer.Serialize(RequireSuccess(_queries.List(_fixture.BatchToken)).Queries);

    private static void AssertRenameMetadata(RenameResult result, string oldName, string newName)
    {
        Assert.Equal("power-query", result.ObjectType);
        Assert.Equal(oldName, result.OldName);
        Assert.Equal(newName, result.NewName);
        Assert.Equal(oldName.Trim(), result.NormalizedOldName);
        Assert.Equal(newName.Trim(), result.NormalizedNewName);
    }

    private static string UniqueName(string prefix) => $"{prefix}_{Guid.NewGuid():N}"[..(prefix.Length + 9)];
}
