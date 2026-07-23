using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.PowerQuery;
using Xunit;
using IOFile = System.IO.File;

namespace Sbroenne.ExcelMcp.Core.Tests.Commands.Connection;

/// <summary>
/// Tests for Connection Delete operations
/// </summary>
public partial class ConnectionCommandsTests
{
    [Fact]
    public void Delete_ExistingTextConnection_ReturnsSuccess()
    {
        // Arrange
        var testFile = _fixture.CreateTestFile();

        // Use ODBC connection (doesn't need actual DSN for delete test)
        string connectionString = "ODBC;DSN=TestDSN;DBQ=C:\\temp\\test.xlsx";
        string connectionName = "DeleteTestConnection";

        using var batch = ExcelSession.BeginBatch(testFile);

        // Create connection first
        _commands.Create(batch, connectionName, connectionString);

        // Verify connection exists
        var listResultBefore = _commands.List(batch);
        Assert.True(listResultBefore.Success);
        Assert.Contains(listResultBefore.Connections, c => c.Name == connectionName);

        // Act - Delete the connection
        // Assert
        _commands.Delete(batch, connectionName);

        // Verify connection no longer exists
        var listResultAfter = _commands.List(batch);
        Assert.True(listResultAfter.Success);
        Assert.DoesNotContain(listResultAfter.Connections, c => c.Name == connectionName);
    }

    [Fact]
    public void Delete_NonExistentConnection_ThrowsException()
    {
        // Arrange
        var testFile = _fixture.CreateTestFile();

        string connectionName = "NonExistentConnection";

        using var batch = ExcelSession.BeginBatch(testFile);

        // Act & Assert - Attempting to delete non-existent connection should throw
        var exception = Assert.Throws<InvalidOperationException>(() =>
        {
            _commands.Delete(batch, connectionName);
        });

        Assert.Contains("not found", exception.Message);
    }

    [Fact]
    public void Delete_AfterCreatingMultiple_RemovesOnlySpecified()
    {
        // Arrange
        var testFile = _fixture.CreateTestFile();

        // Use ODBC connections (don't need actual DSNs for delete test)
        string conn1Name = "Connection1";
        string conn2Name = "Connection2";
        string conn3Name = "Connection3";

        using var batch = ExcelSession.BeginBatch(testFile);

        // Create three connections
        _commands.Create(batch, conn1Name, "ODBC;DSN=TestDSN1;DBQ=C:\\temp\\test1.xlsx");
        _commands.Create(batch, conn2Name, "ODBC;DSN=TestDSN2;DBQ=C:\\temp\\test2.xlsx");
        _commands.Create(batch, conn3Name, "ODBC;DSN=TestDSN3;DBQ=C:\\temp\\test3.xlsx");

        // Act - Delete only the second connection
        // Assert
        _commands.Delete(batch, conn2Name);

        // Verify only conn2 is deleted
        var listResult = _commands.List(batch);
        Assert.True(listResult.Success);
        Assert.Contains(listResult.Connections, c => c.Name == conn1Name);
        Assert.DoesNotContain(listResult.Connections, c => c.Name == conn2Name);
        Assert.Contains(listResult.Connections, c => c.Name == conn3Name);
    }

    [Fact]
    public void Delete_ConnectionWithDescription_RemovesSuccessfully()
    {
        // Arrange
        var testFile = _fixture.CreateTestFile();

        string connectionName = "DescribedConnection";
        string description = "Test connection with description";
        string connectionString = "ODBC;DSN=DescribedDSN;DBQ=C:\\temp\\described.xlsx";

        using var batch = ExcelSession.BeginBatch(testFile);

        // Create connection with description
        _commands.Create(batch, connectionName, connectionString, null, description);

        // Act - Delete connection
        // Assert
        _commands.Delete(batch, connectionName);

        var listResult = _commands.List(batch);
        Assert.DoesNotContain(listResult.Connections, c => c.Name == connectionName);
    }

    [Fact]
    public void Delete_ImmediatelyAfterCreate_WorksCorrectly()
    {
        // Arrange
        var testFile = _fixture.CreateTestFile();

        string connectionName = "ImmediateDeleteTest";
        string connectionString = "ODBC;DSN=ImmediateDSN;DBQ=C:\\temp\\immediate.xlsx";

        using var batch = ExcelSession.BeginBatch(testFile);

        // Act - Create and immediately delete
        _commands.Create(batch, connectionName, connectionString);

        // Assert
        _commands.Delete(batch, connectionName);

        var listResult = _commands.List(batch);
        Assert.DoesNotContain(listResult.Connections, c => c.Name == connectionName);
    }

    [Fact]
    public void Delete_ConnectionAfterViewOperation_RemovesSuccessfully()
    {
        // Arrange
        var testFile = _fixture.CreateTestFile();

        string connectionName = "ViewThenDelete";
        string connectionString = "ODBC;DSN=ViewDeleteDSN;DBQ=C:\\temp\\viewdelete.xlsx";

        using var batch = ExcelSession.BeginBatch(testFile);

        // Create and view connection
        _commands.Create(batch, connectionName, connectionString);

        var viewResult = _commands.View(batch, connectionName);
        Assert.True(viewResult.Success);
        Assert.Equal(connectionName, viewResult.ConnectionName);

        // Act - Delete after viewing
        // Assert
        _commands.Delete(batch, connectionName);

        var listResult = _commands.List(batch);
        Assert.DoesNotContain(listResult.Connections, c => c.Name == connectionName);
    }

    [Fact]
    public void Delete_EmptyConnectionName_ThrowsException()
    {
        // Arrange
        var testFile = _fixture.CreateTestFile();

        using var batch = ExcelSession.BeginBatch(testFile);

        // Act & Assert - Empty connection name should throw
        var exception = Assert.Throws<InvalidOperationException>(() =>
        {
            _commands.Delete(batch, string.Empty);
        });

        Assert.Contains("not found", exception.Message);
    }

    [Fact]
    public void Delete_RepeatedDeleteAttempts_SecondAttemptFails()
    {
        // Arrange
        var testFile = _fixture.CreateTestFile();

        string connectionName = "DoubleDeleteTest";
        string connectionString = "ODBC;DSN=DoubleDeleteDSN;DBQ=C:\\temp\\doubledelete.xlsx";

        using var batch = ExcelSession.BeginBatch(testFile);

        // Create connection
        _commands.Create(batch, connectionName, connectionString);

        // Act - First delete
        _commands.Delete(batch, connectionName);

        // Act & Assert - Second delete should fail
        var exception = Assert.Throws<InvalidOperationException>(() =>
        {
            _commands.Delete(batch, connectionName);
        });

        Assert.Contains("not found", exception.Message);
    }

    #region Orphaned Power Query Connection Tests

    /// <summary>
    /// Tests that orphaned Power Query connections (generic names like "Connection", "Connection1")
    /// can be deleted via the connection API even though they use the Mashup provider.
    /// These connections don't follow the standard "Query - {name}" pattern.
    /// </summary>
    [Fact]
    public void Delete_OrphanedPowerQueryConnection_GenericName_Succeeds()
    {
        // Arrange - Use the test file that has orphaned connections
        var sourceFile = Path.Combine(AppContext.BaseDirectory, "TestData", "MSXI Baseline.xlsx");

        if (!IOFile.Exists(sourceFile))
        {
            // Skip if test data file doesn't exist
            return;
        }

        var testFile = Path.Combine(_fixture.TempDir, $"OrphanedPQ_{Guid.NewGuid():N}.xlsx");
        IOFile.Copy(sourceFile, testFile);

        using var batch = ExcelSession.BeginBatch(testFile);

        // Verify the orphaned connection exists
        var listBefore = _commands.List(batch);
        var orphanedConn = listBefore.Connections.FirstOrDefault(c => c.Name == "Connection");
        Assert.NotNull(orphanedConn);
        Assert.True(orphanedConn.IsPowerQuery, "Connection should be detected as Power Query");

        // Act - Delete the orphaned connection
        _commands.Delete(batch, "Connection");

        // Assert - Connection should be removed
        var listAfter = _commands.List(batch);
        Assert.DoesNotContain(listAfter.Connections, c => c.Name == "Connection");
    }

    /// <summary>
    /// Tests that a Power Query connection following the "Query - {name}" pattern
    /// but with no corresponding query can be deleted.
    /// </summary>
    [Fact]
    public void Delete_OrphanedPowerQueryConnection_StandardNameMissingQuery_Succeeds()
    {
        // Arrange - Use the test file that has orphaned connections
        var sourceFile = Path.Combine(AppContext.BaseDirectory, "TestData", "MSXI Baseline.xlsx");

        if (!IOFile.Exists(sourceFile))
        {
            // Skip if test data file doesn't exist
            return;
        }

        var testFile = Path.Combine(_fixture.TempDir, $"OrphanedPQ2_{Guid.NewGuid():N}.xlsx");
        IOFile.Copy(sourceFile, testFile);

        using var batch = ExcelSession.BeginBatch(testFile);

        // Verify the orphaned connection exists (Query - 2 ReservationYearsBaseline has no matching query)
        var listBefore = _commands.List(batch);
        var orphanedConn = listBefore.Connections.FirstOrDefault(c => c.Name == "Query - 2 ReservationYearsBaseline");
        Assert.NotNull(orphanedConn);
        Assert.True(orphanedConn.IsPowerQuery, "Connection should be detected as Power Query");

        // Act - Delete the orphaned connection
        _commands.Delete(batch, "Query - 2 ReservationYearsBaseline");

        // Assert - Connection should be removed
        var listAfter = _commands.List(batch);
        Assert.DoesNotContain(listAfter.Connections, c => c.Name == "Query - 2 ReservationYearsBaseline");
    }

    /// <summary>
    /// Tests that a valid Power Query connection (with matching query) cannot be deleted
    /// via the connection API - should redirect to powerquery.
    /// </summary>
    [Fact]
    public void Delete_ValidPowerQueryConnection_ThrowsWithRedirect()
    {
        // Arrange - Use the test file that has valid Power Query connections
        var sourceFile = Path.Combine(AppContext.BaseDirectory, "TestData", "MSXI Baseline.xlsx");

        if (!IOFile.Exists(sourceFile))
        {
            // Skip if test data file doesn't exist
            return;
        }

        var testFile = Path.Combine(_fixture.TempDir, $"ValidPQ_{Guid.NewGuid():N}.xlsx");
        IOFile.Copy(sourceFile, testFile);

        using var batch = ExcelSession.BeginBatch(testFile);

        // "Query - Milestones" has a matching query named "Milestones"
        var listBefore = _commands.List(batch);
        var validConn = listBefore.Connections.FirstOrDefault(c => c.Name == "Query - Milestones");
        Assert.NotNull(validConn);
        Assert.True(validConn.IsPowerQuery, "Connection should be detected as Power Query");

        // Act & Assert - Should throw with redirect message
        var exception = Assert.Throws<InvalidOperationException>(() =>
        {
            _commands.Delete(batch, "Query - Milestones");
        });

        Assert.Contains("powerquery", exception.Message);
    }

    /// <summary>
    /// Verifies that IsOrphanedPowerQueryConnection correctly identifies orphaned connections.
    /// </summary>
    [Fact]
    public void IsOrphanedPowerQueryConnection_GenericNamedConnection_ReturnsTrue()
    {
        // Arrange - Use the test file that has orphaned connections
        var sourceFile = Path.Combine(AppContext.BaseDirectory, "TestData", "MSXI Baseline.xlsx");

        if (!IOFile.Exists(sourceFile))
        {
            // Skip if test data file doesn't exist
            return;
        }

        var testFile = Path.Combine(_fixture.TempDir, $"IsOrphaned_{Guid.NewGuid():N}.xlsx");
        IOFile.Copy(sourceFile, testFile);

        using var batch = ExcelSession.BeginBatch(testFile);

        // Act - Check if generic-named connections are orphaned
        var result = batch.Execute((ctx, ct) =>
        {
            dynamic? conn = null;
            try
            {
                // "Connection" is a generic-named Power Query connection
                conn = ctx.Book.Connections["Connection"];
                return PowerQueryHelpers.IsOrphanedPowerQueryConnection(ctx.Book, conn);
            }
            finally
            {
                if (conn != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(conn);
                }
            }
        });

        // Assert - Generic-named Power Query connections are always orphaned
        Assert.True(result);
    }

    /// <summary>
    /// Verifies that IsOrphanedPowerQueryConnection correctly identifies valid connections.
    /// </summary>
    [Fact]
    public void IsOrphanedPowerQueryConnection_ValidConnection_ReturnsFalse()
    {
        // Arrange - Use the test file that has valid Power Query connections
        var sourceFile = Path.Combine(AppContext.BaseDirectory, "TestData", "MSXI Baseline.xlsx");

        if (!IOFile.Exists(sourceFile))
        {
            // Skip if test data file doesn't exist
            return;
        }

        var testFile = Path.Combine(_fixture.TempDir, $"IsValid_{Guid.NewGuid():N}.xlsx");
        IOFile.Copy(sourceFile, testFile);

        using var batch = ExcelSession.BeginBatch(testFile);

        // Act - Check if "Query - Milestones" (has matching query) is orphaned
        var result = batch.Execute((ctx, ct) =>
        {
            dynamic? conn = null;
            try
            {
                // "Query - Milestones" has a matching "Milestones" query
                conn = ctx.Book.Connections["Query - Milestones"];
                return PowerQueryHelpers.IsOrphanedPowerQueryConnection(ctx.Book, conn);
            }
            finally
            {
                if (conn != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(conn);
                }
            }
        });

        // Assert - This is NOT orphaned because the query exists
        Assert.False(result);
    }

    #endregion
}




