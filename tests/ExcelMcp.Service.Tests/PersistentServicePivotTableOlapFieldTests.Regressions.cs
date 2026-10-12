using Sbroenne.ExcelMcp.Core.Commands.Table;
using Sbroenne.ExcelMcp.Core.Commands;
using Sbroenne.ExcelMcp.Core.Models;
using Xunit;

namespace Sbroenne.ExcelMcp.Service.Tests;

public partial class PersistentServicePivotTableOlapFieldTests
{
    [Fact]
    public void AddValueField_UnknownOlapField_PreservesStrategyFailure()
    {
        var batch = _fixture.BatchToken;
        var (_, pivot) = CreateIsolatedPivot();
        var before = RequireSuccess(_pivotCommands.Read(batch, pivot)).PivotTable;

        var result = _pivotCommands.AddValueField(batch, pivot, "MissingMeasure");

        Assert.False(result.Success);
        Assert.Contains("Cannot determine table and column for field 'MissingMeasure'",
            result.ErrorMessage, StringComparison.Ordinal);
        var after = RequireSuccess(_pivotCommands.Read(batch, pivot)).PivotTable;
        Assert.Equal(before.Range, after.Range);
        Assert.Equal(before.ValueFieldCount, after.ValueFieldCount);
        Assert.Equal(before.LayoutError, after.LayoutError);
    }

    [Fact]
    public void AddValueField_DottedColumnPreservesTheExactDataModelColumn()
    {
        var batch = _fixture.BatchToken;
        var source = _fixture.CreateTestSheet(batch);
        var destination = _fixture.CreateTestSheet(batch);
        var tableName = "DottedSales" + Guid.NewGuid().ToString("N")[..8];
        RequireSuccess(_commands.SetValues(batch, source, "A1:B3",
            [["Category", "Net.Amount"], ["One", 10], ["Two", 20]]));
        var tables = _fixture.CreateCommands<ITableCommands>();
        RequireSuccess(tables.Create(batch, source, tableName, "A1:B3"));
        _fixture.RegisterTableForCleanup(tableName);
        RequireSuccess(tables.AddToDataModel(batch, tableName));
        _fixture.RegisterDataModelTableForCleanup(tableName);
        var pivotName = tableName + "_Pivot";
        RequireSuccess(_pivotCommands.CreateFromDataModel(batch, tableName, destination, "A3", pivotName));

        var added = RequireSuccess(_pivotCommands.AddValueField(batch, pivotName, $"[{tableName}].[Net.Amount]"));
        _fixture.RegisterDataModelMeasureForCleanup(added.FieldName);

        Assert.Equal("Net.Amount Sum", added.FieldName);
        Assert.Equal($"SUM('{tableName}'[Net.Amount])",
            Assert.Single(RequireSuccess(_dataModelCommands.ListMeasures(batch, tableName)).Measures).FormulaPreview);
        Assert.Equal(30d, Convert.ToDouble(RequireSuccess(_pivotCommands.GetData(batch, pivotName)).Values[^1][0],
            System.Globalization.CultureInfo.InvariantCulture));
    }

    [Fact]
    public void AddValueField_EscapedColumnPreservesFormulaThroughAggregationChange()
    {
        var batch = _fixture.BatchToken;
        var query = "O'Brien" + Guid.NewGuid().ToString("N")[..8];
        var table = query.Replace("'", " ", StringComparison.Ordinal);
        var destination = _fixture.CreateTestSheet(batch);
        var queries = _fixture.CreateCommands<IPowerQueryCommands>();
        RequireSuccess(queries.Create(batch, query,
            "#table(type table [Category = text, #\"Net]Amount\" = number], {{\"One\", 10}, {\"Two\", 20}})",
            PowerQueryLoadMode.LoadToDataModel));
        _fixture.RegisterPowerQueryForCleanup(query);
        Assert.Contains(RequireSuccess(_dataModelCommands.ListTables(batch)).Tables, model => model.Name == table);
        var modelTable = RequireSuccess(_dataModelCommands.ReadTable(batch, table));
        Assert.Contains(modelTable.Columns, column => column.Name == "Net]Amount");
        var pivot = "EscapedPivot" + Guid.NewGuid().ToString("N")[..8];
        RequireSuccess(_pivotCommands.CreateFromDataModel(batch, table, destination, "A3", pivot));

        var added = RequireSuccess(_pivotCommands.AddValueField(batch, pivot, $"[{table}].[Net]]Amount]"));
        _fixture.RegisterDataModelMeasureForCleanup(added.FieldName);

        Assert.Equal("Net]Amount Sum", added.FieldName);
        var escapedTable = table.Replace("'", "''", StringComparison.Ordinal);
        AssertFormulaAndValue($"SUM('{escapedTable}'[Net]]Amount])", 30);
        RequireSuccess(_pivotCommands.SetFieldFunction(
            batch, pivot, "[Measures].[Net]]Amount Sum]", AggregationFunction.Count));
        AssertFormulaAndValue($"COUNT('{escapedTable}'[Net]]Amount])", 2);
        RequireSuccess(_pivotCommands.Refresh(batch, pivot));
        Assert.Equal(2d, Convert.ToDouble(RequireSuccess(_pivotCommands.GetData(batch, pivot)).Values[^1][0],
            System.Globalization.CultureInfo.InvariantCulture));

        void AssertFormulaAndValue(string formula, double expected)
        {
            Assert.Equal(formula, Assert.Single(
                RequireSuccess(_dataModelCommands.ListMeasures(batch, table)).Measures).FormulaPreview);
            AssertMeasureValue(added.FieldName, expected);
        }
    }

    [Fact]
    public void SetFieldFunction_DottedMeasureNameUpdatesTheExactMeasure()
    {
        var batch = _fixture.BatchToken;
        var (_, pivot) = CreateIsolatedPivot();
        RequireSuccess(_dataModelCommands.CreateMeasure(
            batch, "RegionalSalesTable", "Revenue.Total", "SUM('RegionalSalesTable'[Sales])"));
        _fixture.RegisterDataModelMeasureForCleanup("Revenue.Total");
        RequireSuccess(_pivotCommands.Refresh(batch, pivot));
        RequireSuccess(_pivotCommands.AddValueField(batch, pivot, "[Measures].[Revenue.Total]"));
        AssertMeasureValue("Revenue.Total", 49000);

        RequireSuccess(_pivotCommands.SetFieldFunction(
            batch, pivot, "[Measures].[Revenue.Total]", AggregationFunction.Count));

        var measures = RequireSuccess(_dataModelCommands.ListMeasures(batch, "RegionalSalesTable"));
        Assert.Equal("COUNT('RegionalSalesTable'[Sales])",
            Assert.Single(measures.Measures, measure => measure.Name == "Revenue.Total").FormulaPreview);
        AssertMeasureValue("Revenue.Total", 8);
    }

    [Theory]
    [InlineData("Margin]Pct")]
    [InlineData("[Measures].[Margin]]Pct]")]
    public void AddValueField_EscapedMeasureNameResolvesAndUpdatesExactModelMeasure(string fieldName)
    {
        var batch = _fixture.BatchToken;
        var (_, pivot) = CreateIsolatedPivot();
        RequireSuccess(_dataModelCommands.CreateMeasure(
            batch, "RegionalSalesTable", "Margin]Pct", "SUM('RegionalSalesTable'[Sales])"));
        _fixture.RegisterDataModelMeasureForCleanup("Margin]Pct");
        RequireSuccess(_pivotCommands.Refresh(batch, pivot));

        var added = RequireSuccess(_pivotCommands.AddValueField(batch, pivot, fieldName));

        Assert.Equal("Margin]Pct", added.FieldName);
        AssertMeasureValue("Margin]Pct", 49000);
        RequireSuccess(_pivotCommands.SetFieldFunction(
            batch, pivot, "[Measures].[Margin]]Pct]", AggregationFunction.Count));
        Assert.Equal("COUNT('RegionalSalesTable'[Sales])",
            Assert.Single(RequireSuccess(_dataModelCommands.ListMeasures(batch, "RegionalSalesTable")).Measures,
                measure => measure.Name == "Margin]Pct").FormulaPreview);
        AssertMeasureValue("Margin]Pct", 8);
    }

    [Theory]
    [InlineData("SUM('RegionalSalesTable'[Sales]) * 2")]
    [InlineData("SUM('RegionalSalesTable'[Sales]) + SUM('RegionalSalesTable'[Units])")]
    public void SetFieldFunction_CompoundMeasureIsRejectedWithoutChangingFormula(string formula)
    {
        var batch = _fixture.BatchToken;
        var (_, pivot) = CreateIsolatedPivot();
        RequireSuccess(_dataModelCommands.CreateMeasure(batch, "RegionalSalesTable", "Compound Revenue", formula));
        _fixture.RegisterDataModelMeasureForCleanup("Compound Revenue");
        RequireSuccess(_pivotCommands.Refresh(batch, pivot));
        RequireSuccess(_pivotCommands.AddValueField(batch, pivot, "[Measures].[Compound Revenue]"));
        var before = RequireSuccess(_dataModelCommands.ListMeasures(batch, "RegionalSalesTable"))
            .Measures.Single(measure => measure.Name == "Compound Revenue").FormulaPreview;

        var result = _pivotCommands.SetFieldFunction(
            batch, pivot, "[Measures].[Compound Revenue]", AggregationFunction.Count);

        Assert.False(result.Success);
        Assert.Contains("single aggregation", result.ErrorMessage, StringComparison.OrdinalIgnoreCase);
        var after = RequireSuccess(_dataModelCommands.ListMeasures(batch, "RegionalSalesTable"))
            .Measures.Single(measure => measure.Name == "Compound Revenue").FormulaPreview;
        Assert.Equal(before, after);
        AssertMeasureValue("Compound Revenue", formula.EndsWith("* 2", StringComparison.Ordinal) ? 98000 : 49980);
    }
}
