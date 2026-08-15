// Copyright (c) Sbroenne. All rights reserved.
// Licensed under the MIT License.

using System.Globalization;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Commands.DataModel;

public partial class DataModelCommandsTests
{
    [Fact]
    public void ReadConnection_WithDataModel_ReturnsModelConnectionMetadata()
    {
        using var batch = ExcelSession.BeginBatch(_dataModelFile);

        var result = _dataModelCommands.ReadConnection(batch);

        Assert.True(result.Success, $"ReadConnection failed: {result.ErrorMessage}");
        Assert.Equal("ThisWorkbookDataModel", result.ConnectionName);
        Assert.Equal("MODEL", result.ConnectionType);
        Assert.Equal(7, result.ConnectionTypeValue);
        Assert.True(result.InModel);
        Assert.Equal("CUBE", result.CommandType);
        Assert.Equal(1, result.CommandTypeValue);
        Assert.Equal(5, result.TableNames.Count);
        Assert.Contains("SalesTable", result.TableNames);
        Assert.Contains("ProductsTable", result.TableNames);
    }

    [Fact]
    public void RefreshThenReadTable_WithWorksheetSource_ReturnsSourceConnectionMetadata()
    {
        using var batch = ExcelSession.BeginBatch(_dataModelFile);

        var refreshResult = _dataModelCommands.Refresh(batch, "SalesTable");
        var result = _dataModelCommands.ReadTable(batch, "SalesTable");
        var workbookName = Path.GetFileName(_dataModelFile);

        Assert.True(refreshResult.Success, $"Refresh failed: {refreshResult.ErrorMessage}");
        Assert.True(result.Success, $"ReadTable failed: {result.ErrorMessage}");
        Assert.Equal($"WorkbookConnection_{workbookName}!SalesTable", result.SourceConnectionName);
        Assert.Equal("Excel Table: SalesTable", result.SourceConnectionDescription);
        Assert.Equal("WORKSHEET", result.SourceConnectionType);
        Assert.Equal(8, result.SourceConnectionTypeValue);
        Assert.True(result.SourceConnectionInModel);
    }

    [Fact]
    public void ListColumns_WithTypedPiaMetadata_ReturnsRawAndNamedDataTypes()
    {
        using var batch = ExcelSession.BeginBatch(_dataModelFile);

        var result = _dataModelCommands.ListColumns(batch, "SalesTable");
        var salesId = Assert.Single(result.Columns, column => column.Name == "SalesID");

        Assert.True(result.Success, $"ListColumns failed: {result.ErrorMessage}");
        Assert.NotEqual(0, salesId.DataTypeValue);
        Assert.Equal(salesId.DataTypeValue.ToString(CultureInfo.InvariantCulture), salesId.DataType);
        Assert.All(result.Columns, column =>
            Assert.False(
                column.DataTypeName.StartsWith("Unknown", StringComparison.Ordinal),
                $"Excel returned unmapped DataType {column.DataTypeValue} for {column.Name}"));
    }
}
