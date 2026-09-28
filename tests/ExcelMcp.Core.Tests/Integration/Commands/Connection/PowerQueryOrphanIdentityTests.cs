using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Commands;
using Sbroenne.ExcelMcp.Core.Models;
using Sbroenne.ExcelMcp.Core.PowerQuery;
using Sbroenne.ExcelMcp.Core.Tests.Helpers;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Commands.Connection;

[Trait("Category", "Integration")]
[Trait("Layer", "Core")]
[Trait("Feature", "Connection")]
[Trait("RequiresExcel", "true")]
[Collection("Sequential")]
public sealed class PowerQueryOrphanIdentityTests(
    ConnectionTestsFixture fixture) :
    IClassFixture<ConnectionTestsFixture>
{
    [Fact]
    public void IsOrphanedPowerQueryConnection_GenericNamedConnection_ReturnsTrue()
    {
        var testFile = fixture.CreateTestFile();
        using var batch = ExcelSession.BeginBatch(testFile);
        AddMashupConnection(
            batch,
            "Connection",
            $"Missing_{Guid.NewGuid():N}");

        var result = IsOrphaned(batch, "Connection");

        Assert.True(result);
    }

    [Fact]
    public void IsOrphanedPowerQueryConnection_ValidConnection_ReturnsFalse()
    {
        var testFile = fixture.CreateTestFile();
        using var batch = ExcelSession.BeginBatch(testFile);
        var queryName = $"Valid_{Guid.NewGuid():N}"[..24];
        new PowerQueryCommands(new DataModelCommands()).Create(
            batch,
            queryName,
            "let Source = #table({\"Value\"}, {{1}}) in Source",
            PowerQueryLoadMode.LoadToTable,
            "Sheet1");

        var result = IsOrphaned(batch, $"Query - {queryName}");

        Assert.False(result);
    }

    private static bool IsOrphaned(
        IExcelBatch batch,
        string connectionName) =>
        batch.Execute((ctx, ct) =>
        {
            dynamic? connection = null;
            try
            {
                connection = ctx.Book.Connections[connectionName];
                return PowerQueryHelpers.IsOrphanedPowerQueryConnection(
                    ctx.Book,
                    connection);
            }
            finally
            {
                ComUtilities.Release(ref connection);
            }
        });

    private static void AddMashupConnection(
        IExcelBatch batch,
        string connectionName,
        string location) =>
        batch.Execute((ctx, ct) =>
        {
            dynamic? connections = null;
            dynamic? connection = null;
            try
            {
                connections = ctx.Book.Connections;
                connection = connections.Add2(
                    Name: connectionName,
                    Description: "Orphaned Power Query helper test",
                    ConnectionString:
                        "OLEDB;Provider=Microsoft.Mashup.OleDb.1;" +
                        $"Data Source=$Workbook$;Location={location};" +
                        "Extended Properties=\"\"",
                    CommandText: $"SELECT * FROM [{location}]",
                    lCmdtype: 2,
                    CreateModelConnection: false,
                    ImportRelationships: false);
            }
            finally
            {
                ComUtilities.Release(ref connection);
                ComUtilities.Release(ref connections);
            }
        });
}
