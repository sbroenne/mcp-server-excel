using System.Globalization;
using System.Text.Json;
using Sbroenne.ExcelMcp.Core.Commands;
using Sbroenne.ExcelMcp.Core.Models;
using Xunit;

namespace Sbroenne.ExcelMcp.Service.Tests;

[Trait("Layer", "Service")]
[Trait("Category", "Integration")]
[Trait("Feature", "PowerQuery")]
[Trait("Feature", "DataModel")]
[Trait("RequiresExcel", "true")]
[Trait("Speed", "Medium")]
[Collection("ServiceWorkflow")]
public class PersistentServicePowerQueryDataModelLoadingTests(
    PersistentServiceWorkbookFixture fixture) :
    PersistentServiceWorkbookTestBase(fixture),
    IClassFixture<PersistentServiceWorkbookFixture>
{
    private readonly IPowerQueryCommands _queries =
        fixture.CreateCommands<IPowerQueryCommands>();
    private readonly IDataModelCommands _model =
        fixture.CreateCommands<IDataModelCommands>();

    [Fact]
    public void Create_LoadToDataModel_TableAppearsInDataModel()
    {
        var name = CreateQuery(["ID", "Name"], [[1, "Alpha"], [2, "Beta"]]);
        AssertLoaded(name, ["ID", "Name"], [[1, "Alpha"], [2, "Beta"]]);
    }

    [Fact]
    public void Update_LoadedToDataModel_PreservesSettingsAndNoDuplicateTables()
    {
        string[] columns = ["ID", "Name", "Amount"];
        var name = CreateQuery(columns, [[1, "Alpha", 100], [2, "Beta", 200]]);
        UpdateAndAssert(name, columns,
            [[1, "Alpha", 150], [2, "Beta", 250], [3, "Gamma", 350]]);
    }

    [Fact]
    public void Update_MultipleUpdatesToDataModel_NoDuplicateTables()
    {
        var name = CreateQuery(["Val"], [[1]]);
        UpdateAndAssert(name, ["Val"], [[2], [12]]);
        UpdateAndAssert(name, ["Val"], [[3], [13], [23]]);
    }

    [Fact]
    public void Refresh_LoadedToDataModel_PreservesSettings() =>
        AssertDeferredRefresh(PowerQueryLoadMode.LoadToDataModel);

    [Fact]
    public void Update_LoadedToBoth_PreservesSettings()
    {
        var name = CreateQuery(["A"], [[1], [11]], PowerQueryLoadMode.LoadToBoth);
        UpdateAndAssert(name, ["A"], [[2], [22], [32]], PowerQueryLoadMode.LoadToBoth);
    }

    [Fact]
    public void Refresh_LoadedToBoth_PreservesSettings() =>
        AssertDeferredRefresh(PowerQueryLoadMode.LoadToBoth);

    [Fact]
    public void GetLoadConfig_MultipleQueriesDifferentModes_ReturnsCorrectModeForEach()
    {
        foreach (var (mode, seed) in new[]
        {
            (PowerQueryLoadMode.ConnectionOnly, 17),
            (PowerQueryLoadMode.LoadToTable, 29),
            (PowerQueryLoadMode.LoadToDataModel, 43),
            (PowerQueryLoadMode.LoadToBoth, 61)
        })
        {
            var name = CreateQuery(["Val"], [[seed], [seed + 1]], mode);
            AssertLoaded(name, ["Val"], [[seed], [seed + 1]], mode);
        }
    }

    [Fact]
    public void Update_LoadedToDataModel_AddColumn_HandlesSchemaChange()
    {
        var name = CreateQuery(["ID", "Name"], [[1, "Alpha"], [2, "Beta"]]);
        UpdateAndAssert(name, ["ID", "Name", "Amount"],
            [[1, "Alpha", 100], [2, "Beta", 200], [3, "Gamma", 300]]);
    }

    [Fact]
    public void Update_LoadedToDataModel_RemoveColumn_HandlesSchemaChange()
    {
        var name = CreateQuery(["ID", "Name", "Amount"],
            [[1, "Alpha", 100], [2, "Beta", 200]]);
        UpdateAndAssert(name, ["ID", "Name"], [[1, "Alpha"], [2, "Beta"], [3, "Gamma"]]);
    }

    [Fact]
    public void Update_LoadedToDataModel_ChangeColumnType_HandlesSchemaChange()
    {
        var name = CreateQuery(["ID", "Amount"], [[1, 100], [2, 200]]);
        var before = RequireSuccess(_model.ReadTable(_fixture.BatchToken, name));
        UpdateAndAssert(name, ["ID", "Amount"],
            [[1, "One Hundred"], [2, "Two Hundred"], [3, "Three Hundred"]]);
        var after = RequireSuccess(_model.ReadTable(_fixture.BatchToken, name));
        Assert.NotEqual(before.Columns.Single(column => column.Name == "Amount").DataType,
            after.Columns.Single(column => column.Name == "Amount").DataType);
    }

    [Fact]
    public void Update_LoadedToDataModel_WithDaxMeasure_AddColumn_HandlesSchemaChange()
    {
        var name = CreateQuery(["ID", "Amount"], [[1, 100], [2, 200], [3, 300]]);
        var measure = CreateMeasure(name, "Amount", 600);
        UpdateAndAssert(name, ["ID", "Amount", "Category"],
            [[1, 100, "A"], [2, 200, "B"], [3, 300, "A"], [4, 400, "C"]]);
        AssertMeasure(measure, $"SUM('{name}'[Amount])", 1000);
    }

    [Fact]
    public void Update_LoadedToDataModel_WithDaxMeasure_RemoveReferencedColumn_HandlesError()
    {
        var guard = CreateQuery(["Guard"], [[71], [89]]);
        string[] columns = ["ID", "Name", "Amount"];
        object[][] original = [[1, "Alpha", 100], [2, "Beta", 200]];
        var name = CreateQuery(columns, original);
        var formula = $"SUM('{name}'[Amount])";
        var measure = CreateMeasure(name, "Amount", 300);
        UpdateAndAssert(name, ["ID", "Name"], [[1, "Alpha"], [2, "Beta"], [3, "Gamma"]]);

        var definition = RequireSuccess(_model.Read(_fixture.BatchToken, measure));
        Assert.Equal(formula, definition.DaxFormula);
        var before = JsonSerializer.Serialize(RequireSuccess(_model.ListMeasures(_fixture.BatchToken)).Measures);
        var error = Assert.Throws<InvalidOperationException>(() => _model.Evaluate(
            _fixture.BatchToken, $"EVALUATE ROW(\"Value\", [{measure}])"));
        Assert.Contains("datamodel.evaluate failed [ComInterop/InvalidOperationException]", error.Message);
        Assert.Contains("DAX evaluation failed", error.Message);
        Assert.Equal(before,
            JsonSerializer.Serialize(RequireSuccess(_model.ListMeasures(_fixture.BatchToken)).Measures));
        AssertLoaded(name, ["ID", "Name"], [[1, "Alpha"], [2, "Beta"], [3, "Gamma"]]);
        AssertLoaded(guard, ["Guard"], [[71], [89]]);

        UpdateAndAssert(name, columns, original);
        AssertMeasure(measure, formula, 300);
        AssertLoaded(guard, ["Guard"], [[71], [89]]);
    }

    [Fact]
    public void Update_DaxMeasure_AfterSchemaChange_HandlesUpdate()
    {
        var name = CreateQuery(["ID", "Amount"], [[1, 100], [2, 200]]);
        var measure = CreateMeasure(name, "Amount", 300);
        UpdateAndAssert(name, ["ID", "Amount", "Quantity"],
            [[1, 100, 5], [2, 200, 10], [3, 300, 15]]);
        AssertMeasure(measure, $"SUM('{name}'[Amount])", 600);
        var formula = $"SUM('{name}'[Amount]) + SUM('{name}'[Quantity])";
        RequireSuccess(_model.UpdateMeasure(_fixture.BatchToken, measure, formula));
        AssertMeasure(measure, formula, 630);
        AssertLoaded(name, ["ID", "Amount", "Quantity"],
            [[1, 100, 5], [2, 200, 10], [3, 300, 15]]);
    }

    [Fact]
    public void Update_LoadedToDataModel_MultipleSchemaChanges_WithDaxMeasures()
    {
        var name = CreateQuery(["ID", "Value"], [[1, 100], [2, 200]]);
        var measure = CreateMeasure(name, "Value", 300);
        var formula = $"SUM('{name}'[Value])";
        UpdateAndAssert(name, ["ID", "Value", "Category"], [[1, 100, "A"], [2, 200, "B"]]);
        AssertMeasure(measure, formula, 300);
        UpdateAndAssert(name, ["ID", "Value", "Category", "Quantity"],
            [[1, 100, "A", 5], [2, 200, "B", 10], [3, 300, "C", 15]]);
        AssertMeasure(measure, formula, 600);
        UpdateAndAssert(name, ["ID", "Value", "Quantity"],
            [[1, 100, 5], [2, 200, 10], [3, 300, 15], [4, 400, 20]]);
        AssertMeasure(measure, formula, 1000);
    }

    [Fact]
    public void Update_LoadedToDataModel_ComplexMCode_SchemaChange()
    {
        const string initial = """
            let
                Source = #table({"ID", "RawValue"}, {{1, 100}, {2, 200}, {3, 300}}),
                AddedColumn = Table.AddColumn(Source, "DoubleValue", each [RawValue] * 2),
                ChangedType = Table.TransformColumnTypes(AddedColumn, {{"DoubleValue", type number}})
            in ChangedType
            """;
        const string updated = """
            let
                Source = #table({"ID", "RawValue"}, {{1, 100}, {2, 200}, {3, 300}, {4, 400}}),
                AddedColumn = Table.AddColumn(Source, "DoubleValue", each [RawValue] * 2),
                AddedTriple = Table.AddColumn(AddedColumn, "TripleValue", each [RawValue] * 3),
                ChangedType = Table.TransformColumnTypes(AddedTriple,
                    {{"DoubleValue", type number}, {"TripleValue", type number}})
            in ChangedType
            """;
        var name = CreateQuery(["ID", "RawValue", "DoubleValue"],
            [[1, 100, 200], [2, 200, 400], [3, 300, 600]], code: initial);
        var measureName = "Average_" + Guid.NewGuid().ToString("N")[..8];
        const string description = "Independent transformed-value control";
        var formula = $"AVERAGE('{name}'[DoubleValue])";
        RequireSuccess(_model.CreateMeasure(_fixture.BatchToken, name, measureName,
            formula, description: description));
        _fixture.RegisterDataModelMeasureForCleanup(measureName);
        AssertMeasure(measureName, formula, 400);
        RequireSuccess(_queries.Update(_fixture.BatchToken, name, updated));
        PowerQueryStateAssertions.AssertStored(_fixture, name, updated,
            PowerQueryLoadMode.LoadToDataModel, null, ["ID", "RawValue", "DoubleValue", "TripleValue"],
            [[1, 100, 200, 300], [2, 200, 400, 600], [3, 300, 600, 900], [4, 400, 800, 1200]]);
        AssertMeasure(measureName, formula, 500);
        Assert.Equal(description, RequireSuccess(_model.Read(_fixture.BatchToken, measureName)).Description);
    }

    private void AssertDeferredRefresh(PowerQueryLoadMode mode)
    {
        var name = CreateQuery(["Val"], [[42], [52]], mode);
        var code = TableCode(["Val"], [[73], [83], [93]]);
        RequireSuccess(_queries.Update(_fixture.BatchToken, name, code, refresh: false));
        PowerQueryStateAssertions.AssertStored(_fixture, name, code, mode,
            mode == PowerQueryLoadMode.LoadToBoth ? name : null, ["Val"], [[42], [52]]);
        RequireSuccess(_queries.Refresh(_fixture.BatchToken, name, TimeSpan.FromMinutes(5)));
        AssertLoaded(name, ["Val"], [[73], [83], [93]], mode);
    }

    private string CreateQuery(string[] columns, object[][] rows,
        PowerQueryLoadMode mode = PowerQueryLoadMode.LoadToDataModel, string? code = null)
    {
        var name = "PQ_Model_" + Guid.NewGuid().ToString("N")[..8];
        var sheet = mode is PowerQueryLoadMode.LoadToTable or PowerQueryLoadMode.LoadToBoth ? name : null;
        code ??= TableCode(columns, rows);
        RequireSuccess(_queries.Create(_fixture.BatchToken, name, code, mode, sheet));
        _fixture.RegisterPowerQueryForCleanup(name);
        if (sheet is not null) { _fixture.RegisterSheetForCleanup(sheet); }
        PowerQueryStateAssertions.AssertStored(_fixture, name, code, mode, sheet, columns, rows);
        return name;
    }

    private void UpdateAndAssert(string name, string[] columns, object[][] rows,
        PowerQueryLoadMode mode = PowerQueryLoadMode.LoadToDataModel)
    {
        RequireSuccess(_queries.Update(_fixture.BatchToken, name, TableCode(columns, rows)));
        AssertLoaded(name, columns, rows, mode);
    }

    private void AssertLoaded(string name, string[] columns, object[][] rows,
        PowerQueryLoadMode mode = PowerQueryLoadMode.LoadToDataModel) =>
        PowerQueryStateAssertions.AssertStored(_fixture, name, TableCode(columns, rows), mode,
            mode is PowerQueryLoadMode.LoadToTable or PowerQueryLoadMode.LoadToBoth ? name : null,
            columns, rows);

    private string CreateMeasure(string table, string column, int expected)
    {
        var name = "Sum_" + Guid.NewGuid().ToString("N")[..8];
        var formula = $"SUM('{table}'[{column}])";
        RequireSuccess(_model.CreateMeasure(_fixture.BatchToken, table, name, formula));
        _fixture.RegisterDataModelMeasureForCleanup(name);
        AssertMeasure(name, formula, expected);
        return name;
    }

    private void AssertMeasure(string name, string formula, int expected)
    {
        var definition = RequireSuccess(_model.Read(_fixture.BatchToken, name));
        Assert.Equal(name, definition.MeasureName);
        Assert.Equal(formula, definition.DaxFormula);
        var result = RequireSuccess(_model.Evaluate(_fixture.BatchToken,
            $"EVALUATE ROW(\"Value\", [{name}])"));
        Assert.Equal(["[Value]"], result.Columns);
        Assert.Equal(1, result.RowCount);
        Assert.Equal(1, result.ColumnCount);
        PowerQueryStateAssertions.AssertRows([[expected]], result.Rows);
    }

    private static string TableCode(string[] columns, object[][] rows)
    {
        static string Literal(object value) => value is string text
            ? "\"" + text.Replace("\"", "\"\"", StringComparison.Ordinal) + "\""
            : Convert.ToString(value, CultureInfo.InvariantCulture)
                ?? throw new InvalidOperationException("A test seed must have a value.");
        var headers = string.Join(", ", columns.Select(column => Literal(column)));
        var records = string.Join(", ", rows.Select(row => "{" + string.Join(", ", row.Select(Literal)) + "}"));
        var types = string.Join(", ", columns.Select((column, index) =>
            $"{{{Literal(column)}, {(rows[0][index] is string ? "type text" : "Int64.Type")}}}"));
        return $"let Source = #table({{{headers}}}, {{{records}}}), " +
            $"Typed = Table.TransformColumnTypes(Source, {{{types}}}) in Typed";
    }
}
