using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Commands;
using Sbroenne.ExcelMcp.Core.Models;
using Sbroenne.ExcelMcp.Core.Tests.Helpers;
using Xunit;
using Xunit.Abstractions;

namespace Sbroenne.ExcelMcp.Core.Tests.Commands.PowerQuery;

/// <summary>
/// Regression tests for workbooks with mixed connection types.
///
/// Root cause (Issue #323): When a workbook contains non-OLEDB connections
/// (Type=7 ThisWorkbookDataModel, Type=8 workbook data-model connections),
/// accessing .OLEDBConnection on them throws COMException 0x800A03EC.
///
/// The Update/View/Lifecycle code iterates all connections and must gracefully
/// skip non-OLEDB connection types. Loading a Power Query to the Data Model
/// creates a ThisWorkbookDataModel (Type=7) connection, which reproduces the bug.
/// </summary>
[Trait("Layer", "Core")]
[Trait("Category", "Integration")]
[Trait("RequiresExcel", "true")]
[Trait("Feature", "PowerQuery")]
[Trait("Speed", "Medium")]
public class PowerQueryMixedConnectionTypeTests : IClassFixture<TempDirectoryFixture>
{
    private readonly PowerQueryCommands _powerQueryCommands;
    private readonly DataModelCommands _dataModelCommands;
    private readonly TempDirectoryFixture _fixture;
    private readonly ITestOutputHelper _output;

    public PowerQueryMixedConnectionTypeTests(TempDirectoryFixture fixture, ITestOutputHelper output)
    {
        _dataModelCommands = new DataModelCommands();
        _powerQueryCommands = new PowerQueryCommands(_dataModelCommands);
        _fixture = fixture;
        _output = output;
    }

    /// <summary>
    /// Regression test for Issue #323: Update on a connection-only query crashes with
    /// COMException 0x800A03EC when the workbook also has a Data Model connection (Type=7).
    ///
    /// Scenario: Create query A loaded to Data Model (produces ThisWorkbookDataModel
    /// Type=7 connection), then Update query B (connection-only) with refresh.
    /// Before fix: COMException 0x800A03EC thrown iterating connections.
    /// After fix: Update succeeds, skipping non-OLEDB connections.
    /// </summary>
    [Fact]
    public void Update_WithDataModelConnection_DoesNotThrowCOMException()
    {
        // Arrange
        var testFile = _fixture.CreateTestFile();
        var dataModelQueryName = "PQ_DM_" + Guid.NewGuid().ToString("N")[..8];
        var connOnlyQueryName = "PQ_CO_" + Guid.NewGuid().ToString("N")[..8];

        var dataModelMCode = @"let Source = #table({""ID"", ""Value""}, {{1, 100}, {2, 200}}) in Source";
        var connOnlyMCode = @"let Source = #table({""A""}, {{1}}) in Source";
        var updatedMCode = @"let Source = #table({""A"", ""B""}, {{1, 2}}) in Source";

        using var batch = ExcelSession.BeginBatch(testFile);

        // STEP 1: Create a query loaded to Data Model → produces Type=7 connection
        _powerQueryCommands.Create(batch, dataModelQueryName, dataModelMCode, PowerQueryLoadMode.LoadToDataModel);

        // Verify Data Model connection exists
        var loadConfig = _powerQueryCommands.GetLoadConfig(batch, dataModelQueryName);
        Assert.True(loadConfig.Success, $"GetLoadConfig failed: {loadConfig.ErrorMessage}");
        Assert.Equal(PowerQueryLoadMode.LoadToDataModel, loadConfig.LoadMode);
        _output.WriteLine($"Data Model query '{dataModelQueryName}' created with LoadToDataModel");

        // STEP 2: Create a connection-only query
        _powerQueryCommands.Create(batch, connOnlyQueryName, connOnlyMCode, PowerQueryLoadMode.ConnectionOnly);
        _output.WriteLine($"Connection-only query '{connOnlyQueryName}' created");

        // STEP 3: Update the connection-only query - this is what crashed before the fix
        // The Update code iterates all connections including ThisWorkbookDataModel (Type=7)
        var updateResult = _powerQueryCommands.Update(batch, connOnlyQueryName, updatedMCode);

        // Assert - no COMException thrown, update succeeds
        Assert.True(updateResult.Success, $"Update failed: {updateResult.ErrorMessage}");
        _output.WriteLine($"Update succeeded on '{connOnlyQueryName}' in presence of Data Model connection");

        // Verify code was actually updated
        var viewResult = _powerQueryCommands.View(batch, connOnlyQueryName);
        Assert.True(viewResult.Success, $"View failed: {viewResult.ErrorMessage}");
        Assert.Contains("\"B\"", viewResult.MCode);
    }

    /// <summary>
    /// Regression test: View on a connection-only query should not crash when
    /// the workbook has a Data Model connection (Type=7).
    /// </summary>
    [Fact]
    public void View_WithDataModelConnection_DoesNotThrowCOMException()
    {
        // Arrange
        var testFile = _fixture.CreateTestFile();
        var dataModelQueryName = "PQ_DM_" + Guid.NewGuid().ToString("N")[..8];
        var connOnlyQueryName = "PQ_CO_" + Guid.NewGuid().ToString("N")[..8];

        var dataModelMCode = @"let Source = #table({""X""}, {{1}}) in Source";
        var connOnlyMCode = @"let Source = #table({""Y""}, {{2}}) in Source";

        using var batch = ExcelSession.BeginBatch(testFile);

        // Create Data Model query (Type=7 connection) + connection-only query
        _powerQueryCommands.Create(batch, dataModelQueryName, dataModelMCode, PowerQueryLoadMode.LoadToDataModel);
        _powerQueryCommands.Create(batch, connOnlyQueryName, connOnlyMCode, PowerQueryLoadMode.ConnectionOnly);

        // Act - View should work without COMException
        var viewResult = _powerQueryCommands.View(batch, connOnlyQueryName);

        // Assert
        Assert.True(viewResult.Success, $"View failed: {viewResult.ErrorMessage}");
        Assert.Contains("\"Y\"", viewResult.MCode);
        _output.WriteLine("View succeeded in presence of Data Model connection");
    }

    /// <summary>
    /// Regression test: List should work when the workbook has mixed connection types.
    /// </summary>
    [Fact]
    public void List_WithDataModelConnection_Succeeds()
    {
        // Arrange
        var testFile = _fixture.CreateTestFile();
        var dataModelQueryName = "PQ_DM_" + Guid.NewGuid().ToString("N")[..8];
        var connOnlyQueryName = "PQ_CO_" + Guid.NewGuid().ToString("N")[..8];

        var dataModelMCode = @"let Source = #table({""X""}, {{1}}) in Source";
        var connOnlyMCode = @"let Source = #table({""Y""}, {{2}}) in Source";

        using var batch = ExcelSession.BeginBatch(testFile);

        _powerQueryCommands.Create(batch, dataModelQueryName, dataModelMCode, PowerQueryLoadMode.LoadToDataModel);
        _powerQueryCommands.Create(batch, connOnlyQueryName, connOnlyMCode, PowerQueryLoadMode.ConnectionOnly);

        // Act
        var listResult = _powerQueryCommands.List(batch);

        // Assert
        Assert.True(listResult.Success, $"List failed: {listResult.ErrorMessage}");
        Assert.NotNull(listResult.Queries);
        var queryNames = listResult.Queries.Select(q => q.Name).ToList();
        Assert.Contains(dataModelQueryName, queryNames);
        Assert.Contains(connOnlyQueryName, queryNames);
        _output.WriteLine($"List returned {listResult.Queries.Count} queries in presence of Data Model connection");
    }

    /// <summary>
    /// Regression test: Delete (lifecycle) should work when the workbook has
    /// mixed connection types including Type=7.
    /// </summary>
    [Fact]
    public void Delete_WithDataModelConnection_Succeeds()
    {
        // Arrange
        var testFile = _fixture.CreateTestFile();
        var dataModelQueryName = "PQ_DM_" + Guid.NewGuid().ToString("N")[..8];
        var targetQueryName = "PQ_Del_" + Guid.NewGuid().ToString("N")[..8];

        var dataModelMCode = @"let Source = #table({""X""}, {{1}}) in Source";
        var targetMCode = @"let Source = #table({""Y""}, {{2}}) in Source";

        using var batch = ExcelSession.BeginBatch(testFile);

        _powerQueryCommands.Create(batch, dataModelQueryName, dataModelMCode, PowerQueryLoadMode.LoadToDataModel);
        _powerQueryCommands.Create(batch, targetQueryName, targetMCode, PowerQueryLoadMode.ConnectionOnly);

        // Act - Delete should iterate connections without crashing on Type=7
        var deleteResult = _powerQueryCommands.Delete(batch, targetQueryName);

        // Assert
        Assert.True(deleteResult.Success, $"Delete failed: {deleteResult.ErrorMessage}");

        // Verify query is gone
        var listResult = _powerQueryCommands.List(batch);
        Assert.True(listResult.Success);
        var queryNames = listResult.Queries!.Select(q => q.Name).ToList();
        Assert.DoesNotContain(targetQueryName, queryNames);
        Assert.Contains(dataModelQueryName, queryNames);
        _output.WriteLine("Delete succeeded in presence of Data Model connection");
    }

    /// <summary>
    /// Regression test: Update a worksheet-loaded query when a Data Model
    /// connection also exists. Tests the QueryTable + ListObject iteration paths.
    /// </summary>
    [Fact]
    public void Update_WorksheetQuery_WithDataModelConnection_Succeeds()
    {
        // Arrange
        var testFile = _fixture.CreateTestFile();
        var dataModelQueryName = "PQ_DM_" + Guid.NewGuid().ToString("N")[..8];
        var worksheetQueryName = "PQ_WS_" + Guid.NewGuid().ToString("N")[..8];

        var dataModelMCode = @"let Source = #table({""X""}, {{1}}) in Source";
        var worksheetMCode = @"let Source = #table({""Col1""}, {{10}}) in Source";
        var updatedMCode = @"let Source = #table({""Col1"", ""Col2""}, {{10, 20}}) in Source";

        using var batch = ExcelSession.BeginBatch(testFile);

        // Create Data Model query (Type=7 connection)
        _powerQueryCommands.Create(batch, dataModelQueryName, dataModelMCode, PowerQueryLoadMode.LoadToDataModel);

        // Create worksheet-loaded query (has QueryTable/ListObject)
        _powerQueryCommands.Create(batch, worksheetQueryName, worksheetMCode, PowerQueryLoadMode.LoadToTable);

        // Act - Update the worksheet query; iterates connections including Type=7
        var updateResult = _powerQueryCommands.Update(batch, worksheetQueryName, updatedMCode);

        // Assert
        Assert.True(updateResult.Success, $"Update failed: {updateResult.ErrorMessage}");

        var viewResult = _powerQueryCommands.View(batch, worksheetQueryName);
        Assert.True(viewResult.Success, $"View failed: {viewResult.ErrorMessage}");
        Assert.Contains("\"Col2\"", viewResult.MCode);
        _output.WriteLine("Update of worksheet query succeeded in presence of Data Model connection");
    }

    /// <summary>
    /// Verifies that the workbook actually has a non-OLEDB connection (Type=7)
    /// after loading a query to the Data Model. This validates our test setup
    /// actually reproduces the precondition for the bug.
    /// </summary>
    [Fact]
    public void DataModelLoad_CreatesNonOledbConnection()
    {
        // Arrange
        var testFile = _fixture.CreateTestFile();
        var queryName = "PQ_Verify_" + Guid.NewGuid().ToString("N")[..8];
        var mCode = @"let Source = #table({""V""}, {{1}}) in Source";

        using var batch = ExcelSession.BeginBatch(testFile);

        _powerQueryCommands.Create(batch, queryName, mCode, PowerQueryLoadMode.LoadToDataModel);

        // Act - enumerate connection types
        bool hasNonOledbConnection = false;
        batch.Execute((ctx, ct) =>
        {
            dynamic? connections = null;
            try
            {
                connections = ctx.Book.Connections;
                int count = connections.Count;
                _output.WriteLine($"Total connections: {count}");

                for (int i = 1; i <= count; i++)
                {
                    dynamic? conn = null;
                    try
                    {
                        conn = connections.Item(i);
                        string name = conn.Name?.ToString() ?? "(unknown)";
                        int connType = -1;
                        try { connType = (int)conn.Type; } catch { /* ignore */ }

                        _output.WriteLine($"  Connection {i}: '{name}' Type={connType}");

                        if (connType != 1) // Not OLEDB
                        {
                            hasNonOledbConnection = true;

                            // Verify that accessing OLEDBConnection throws
                            try
                            {
                                dynamic? oledb = conn.OLEDBConnection;
                                _output.WriteLine($"    OLEDBConnection accessible (unexpected for Type={connType})");
                                if (oledb != null) ComUtilities.Release(ref oledb!);
                            }
                            catch (System.Runtime.InteropServices.COMException ex)
                            {
                                _output.WriteLine($"    OLEDBConnection throws COMException 0x{ex.HResult:X8} (expected)");
                            }
                        }
                    }
                    finally
                    {
                        ComUtilities.Release(ref conn!);
                    }
                }
            }
            finally
            {
                ComUtilities.Release(ref connections!);
            }

            return 0;
        });

        // Assert - our test setup must produce at least one non-OLEDB connection
        Assert.True(hasNonOledbConnection,
            "Expected at least one non-OLEDB connection (Type != 1) after LoadToDataModel. " +
            "If this fails, the test setup doesn't reproduce the bug precondition.");
    }
}
