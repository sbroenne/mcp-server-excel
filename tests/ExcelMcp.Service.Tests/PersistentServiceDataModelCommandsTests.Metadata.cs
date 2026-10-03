// Copyright (c) Sbroenne. All rights reserved.
// Licensed under the MIT License.

using System.Globalization;
using Xunit;

namespace Sbroenne.ExcelMcp.Service.Tests;

public partial class PersistentServiceDataModelCommandsTests
{
    [Fact]
    public void ReadConnection_WithDataModel_ReturnsModelConnectionMetadata()
    {
        var batch = _fixture.BatchToken;

        var result = RequireSuccess(_dataModelCommands.ReadConnection(batch));

        Assert.Equal("ThisWorkbookDataModel", result.ConnectionName);
        Assert.Equal("MODEL", result.ConnectionType);
        Assert.Equal(7, result.ConnectionTypeValue);
        Assert.True(result.InModel);
        Assert.Equal("CUBE", result.CommandType);
        Assert.Equal(1, result.CommandTypeValue);
        Assert.Equal(
            ["CustomersTable", "DisambiguationTable", "ProductsTable", "RegionalSalesTable", "SalesTable"],
            result.TableNames.Order(StringComparer.Ordinal));
    }

    [Fact]
    public void RefreshThenReadTable_WithWorksheetSource_ReturnsSourceConnectionMetadata()
    {
        var batch = _fixture.BatchToken;

        RequireSuccess(_dataModelCommands.Refresh(batch, "SalesTable"));
        var result = RequireSuccess(_dataModelCommands.ReadTable(batch, "SalesTable"));
        var workbookName = Path.GetFileName(_dataModelFile);

        Assert.Equal($"WorkbookConnection_{workbookName}!SalesTable", result.SourceConnectionName);
        Assert.Equal("Excel Table: SalesTable", result.SourceConnectionDescription);
        Assert.Equal("WORKSHEET", result.SourceConnectionType);
        Assert.Equal(8, result.SourceConnectionTypeValue);
        Assert.True(result.SourceConnectionInModel);
    }

    [Fact]
    public void ListColumns_WithTypedPiaMetadata_ReturnsRawAndNamedDataTypes()
    {
        var batch = _fixture.BatchToken;

        var result = RequireSuccess(_dataModelCommands.ListColumns(batch, "SalesTable"));

        Assert.Equal(
            ["Amount", "CustomerID", "Date", "ProductID", "Quantity", "SalesID"],
            result.Columns.Select(column => column.Name).Order(StringComparer.Ordinal));
        var salesId = Assert.Single(result.Columns, column => column.Name == "SalesID");
        Assert.NotEqual(0, salesId.DataTypeValue);
        Assert.Equal(salesId.DataTypeValue.ToString(CultureInfo.InvariantCulture), salesId.DataType);
        Assert.All(result.Columns, column =>
            Assert.False(
                column.DataTypeName.StartsWith("Unknown", StringComparison.Ordinal),
                $"Excel returned unmapped DataType {column.DataTypeValue} for {column.Name}"));
    }
}
