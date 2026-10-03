using Sbroenne.ExcelMcp.Core.Commands;
using Sbroenne.ExcelMcp.Core.Models;
using Xunit;

namespace Sbroenne.ExcelMcp.Service.Tests;

[Collection("ServiceWorkflow")]
[Trait("Category", "Integration")]
[Trait("Layer", "Service")]
[Trait("Feature", "PowerQuery")]
[Trait("RequiresExcel", "true")]
public sealed partial class PersistentServicePowerQueryEvaluateTests(
    PersistentServiceWorkbookFixture fixture) :
    PersistentServiceWorkbookTestBase(fixture),
    IClassFixture<PersistentServiceWorkbookFixture>
{
    private const string GuardCode =
        "let Source = #table(type table [Guard = Int64.Type, Label = text], " +
        "{{47, \"KeepA\"}, {83, \"KeepB\"}}) in Source";
    private readonly IPowerQueryCommands _queries =
        fixture.CreateCommands<IPowerQueryCommands>();

    [Fact]
    public void Evaluate_SimpleTable_ReturnsData()
    {
        const string code = """
            let
                Source = #table({"Name", "Value"}, {{"Test1", 100}, {"Test2", 200}})
            in Source
            """;
        EvaluateChecked(code, ["Name", "Value"], [["Test1", 100], ["Test2", 200]]);
    }

    [Fact]
    public void Evaluate_SingleColumn_ReturnsData() =>
        EvaluateChecked("let Source = #table({\"SingleCol\"}, {{1}, {2}, {3}}) in Source",
            ["SingleCol"], [[1], [2], [3]]);

    [Fact]
    public void Evaluate_VariousDataTypes_ReturnsCorrectValues()
    {
        const string code = """
            let
                Source = #table({"Text", "Number", "Boolean"},
                    {{"Hello", 42, true}, {"World", 3.14, false}})
            in Source
            """;
        EvaluateChecked(code, ["Text", "Number", "Boolean"],
            [["Hello", 42, true], ["World", 3.14m, false]]);
    }

    [Fact]
    public void Evaluate_EmptyMCode_ThrowsArgumentException()
    {
        var guard = CreateEvaluationGuard();
        var before = SnapshotNativeArtifacts();
        var error = Assert.Throws<ArgumentException>(() => _queries.Evaluate(_fixture.BatchToken, ""));
        Assert.Contains("mCode", error.Message, StringComparison.OrdinalIgnoreCase);
        Assert.Equal(before, SnapshotNativeArtifacts());
        AssertEvaluationGuard(guard);
        var recovered = RequireSuccess(_queries.Evaluate(_fixture.BatchToken,
            "let Source = #table({\"Value\"}, {{37}}) in Source"));
        Assert.Equal(["Value"], recovered.Columns);
        Assert.Equal(1, recovered.RowCount);
        Assert.Equal(1, recovered.ColumnCount);
        PowerQueryStateAssertions.AssertRows([[37]], recovered.Rows);
        Assert.Equal(before, SnapshotNativeArtifacts());
        AssertEvaluationGuard(guard);
    }

    [Fact]
    public void Evaluate_WithTransformations_ReturnsTransformedData()
    {
        const string code = """
            let
                Source = #table({"Value"}, {{1}, {2}, {3}, {4}, {5}}),
                Filtered = Table.SelectRows(Source, each [Value] > 2),
                Added = Table.AddColumn(Filtered, "Doubled", each [Value] * 2)
            in Added
            """;
        EvaluateChecked(code, ["Value", "Doubled"], [[3, 6], [4, 8], [5, 10]]);
    }

    private void EvaluateChecked(string code, string[] columns, object[][] rows)
    {
        var guard = CreateEvaluationGuard();
        var before = SnapshotNativeArtifacts();
        var result = RequireSuccess(_queries.Evaluate(_fixture.BatchToken, code));
        Assert.Equal(code, result.MCode);
        Assert.Equal(columns, result.Columns);
        Assert.Equal(columns.Length, result.ColumnCount);
        Assert.Equal(rows.Length, result.RowCount);
        PowerQueryStateAssertions.AssertRows(rows, result.Rows);
        Assert.Equal(before, SnapshotNativeArtifacts());
        AssertEvaluationGuard(guard);
    }

    private string CreateEvaluationGuard()
    {
        var name = "EvalGuard_" + Guid.NewGuid().ToString("N")[..8];
        RequireSuccess(_queries.Create(_fixture.BatchToken, name, GuardCode,
            PowerQueryLoadMode.LoadToBoth, name));
        _fixture.RegisterPowerQueryForCleanup(name);
        _fixture.RegisterSheetForCleanup(name);
        AssertEvaluationGuard(name);
        return name;
    }

    private void AssertEvaluationGuard(string name) =>
        PowerQueryStateAssertions.AssertStored(_fixture, name, GuardCode,
            PowerQueryLoadMode.LoadToBoth, name, ["Guard", "Label"], [[47, "KeepA"], [83, "KeepB"]]);
}
