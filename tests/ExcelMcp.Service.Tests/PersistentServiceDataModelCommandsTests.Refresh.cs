// Copyright (c) Sbroenne. All rights reserved.
// Licensed under the MIT License.

using System.Globalization;
using Sbroenne.ExcelMcp.Core.Commands.Table;
using Sbroenne.ExcelMcp.Core.Commands.Range;
using Xunit;

namespace Sbroenne.ExcelMcp.Service.Tests;

/// <summary>
/// Integration tests for Data Model Refresh operations.
/// Uses shared DataModelPivotTableFixture (non-destructive refresh).
/// </summary>
public partial class PersistentServiceDataModelCommandsTests
{
    #region Refresh Tests

    /// <summary>
    /// Refreshes the entire Data Model.
    /// LLM use case: "refresh the data model"
    /// </summary>
    [Fact]
    public void Refresh_EntireModel_Succeeds()
    {
        AssertRefreshUpdatesSource(specificTable: false);
    }

    /// <summary>
    /// Refreshes a specific Data Model table by name.
    /// LLM use case: "refresh the SalesTable in the data model"
    /// </summary>
    [Fact]
    public void Refresh_SpecificTable_Succeeds()
    {
        AssertRefreshUpdatesSource(specificTable: true);
    }

    /// <summary>
    /// Refreshing a non-existent table throws InvalidOperationException.
    /// LLM use case: error handling for typo in table name
    /// </summary>
    [Fact]
    public void Refresh_InvalidTableName_ThrowsInvalidOperationException()
    {
        var batch = _fixture.BatchToken;
        var before = ReadModelTotal("SalesTable");

        var ex = Assert.Throws<InvalidOperationException>(
            () => _dataModelCommands.Refresh(batch, tableName: "NonExistentTable"));

        Assert.Contains("NonExistentTable", ex.Message);
        Assert.Equal(before, ReadModelTotal("SalesTable"));
    }

    /// <summary>
    /// Refresh with explicit timeout succeeds when within time limit.
    /// </summary>
    [Fact]
    public void Refresh_WithExplicitTimeout_Succeeds()
    {
        AssertRefreshUpdatesSource(specificTable: false, TimeSpan.FromMinutes(5));
    }

    private void AssertRefreshUpdatesSource(bool specificTable, TimeSpan? timeout = null)
    {
        var batch = _fixture.BatchToken;
        var sheet = _fixture.CreateTestSheet(batch);
        var table = $"Refresh_{Guid.NewGuid():N}"[..24];
        var tables = _fixture.CreateCommands<ITableCommands>();
        RequireSuccess(_commands.SetValues(batch, sheet, "A1:B3", [["ID", "Amount"], [1, 10], [2, 20]]));
        RequireSuccess(tables.Create(batch, sheet, table, "A1:B3"));
        RequireSuccess(tables.AddToDataModel(batch, table));
        _fixture.RegisterDataModelTableForCleanup(table);
        Assert.Equal(30m, ReadModelTotal(table));
        var untouched = ReadModelTotal("SalesTable");

        RequireSuccess(_commands.SetValues(batch, sheet, "B2", [[70]], overwritePolicy: OverwritePolicy.Allow));
        Assert.Equal(30m, ReadModelTotal(table));

        var result = _dataModelCommands.Refresh(batch, specificTable ? table : null, timeout);

        RequireSuccess(result);
        Assert.Equal(_dataModelFile, result.FilePath);
        Assert.Equal(90m, ReadModelTotal(table));
        Assert.Equal(untouched, ReadModelTotal("SalesTable"));
        var tableResult = RequireSuccess(_dataModelCommands.ReadTable(batch, table));
        Assert.Equal(2, tableResult.RecordCount);
        Assert.Equal(["Amount", "ID"], tableResult.Columns.Select(column => column.Name).Order(StringComparer.Ordinal));
    }

    private decimal ReadModelTotal(string table)
    {
        var result = _dataModelCommands.Evaluate(_fixture.BatchToken,
            $"EVALUATE ROW(\"Total\", SUM('{table}'[Amount]))");
        RequireSuccess(result);
        Assert.Equal(1, result.RowCount);
        Assert.Equal(1, result.ColumnCount);
        Assert.Equal("[Total]", Assert.Single(result.Columns));
        return Convert.ToDecimal(Assert.Single(Assert.Single(result.Rows)), CultureInfo.InvariantCulture);
    }

    #endregion
}
