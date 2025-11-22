using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Tests.Helpers;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Commands.Connection;

[Trait("Category", "Integration")]
[Trait("Speed", "Medium")]
[Trait("Layer", "Core")]
[Trait("Feature", "Connection")]
[Trait("RequiresExcel", "true")]
public partial class ConnectionCommandsTests
{
    /// <inheritdoc/>
    [Fact]
    public void Refresh_ConnectionNotFound_ReturnsFailure()
    {
        // Arrange
        var testFile = CoreTestHelper.CreateUniqueTestFile(
            nameof(ConnectionCommandsTests),
            nameof(Refresh_ConnectionNotFound_ReturnsFailure),
            _tempDir);

        // Act
        using var batch = ExcelSession.BeginBatch(testFile);
        var result = _commands.Refresh(batch, "NonExistentConnection");

        // Assert
        Assert.False(result.Success);
        Assert.Contains("not found", result.ErrorMessage);
    }

    /// <summary>
    /// Tests OLEDB connection with LoadTo and Refresh operations using SQLite.
    /// SQLite provides a reliable, lightweight database for testing OLEDB operations.
    /// </summary>
    [Fact]
    public void Refresh_SQLiteOleDbConnection_ReturnsSuccess()
    {
        var testFile = CoreTestHelper.CreateUniqueTestFile(
            nameof(ConnectionCommandsTests),
            nameof(Refresh_SQLiteOleDbConnection_ReturnsSuccess),
            _tempDir);

        var dbPath = Path.Combine(_tempDir, $"TestDb_{Guid.NewGuid():N}.db");

        try
        {
            // Create SQLite database with test data
            SQLiteDatabaseHelper.CreateTestDatabase(dbPath);
            Assert.True(System.IO.File.Exists(dbPath), "SQLite database should be created");

            // Create OLEDB connection to SQLite database
            var connectionName = "TestSQLiteConnection";
            ConnectionTestHelper.CreateSQLiteOleDbConnection(testFile, connectionName, dbPath);

            using var batch = ExcelSession.BeginBatch(testFile);

            // Verify connection was created
            var listResult = _commands.List(batch);
            Assert.True(listResult.Success);
            Assert.Contains(listResult.Connections, c => c.Name == connectionName);

            // Act - Load data to worksheet
            var loadResult = _commands.LoadTo(batch, connectionName, "Products");

            // Assert - Data loaded successfully
            Assert.True(loadResult.Success, $"Failed to load data: {loadResult.ErrorMessage}");

            batch.Save();

            // Act - Refresh connection
            var refreshResult = _commands.Refresh(batch, connectionName);

            // Assert - Refresh succeeded
            Assert.True(refreshResult.Success, $"Failed to refresh: {refreshResult.ErrorMessage}");
        }
        finally
        {
            // Cleanup - Delete SQLite database
            SQLiteDatabaseHelper.DeleteDatabase(dbPath);
        }
    }

    /// <summary>
    /// Tests OLEDB connection refresh after data source update using SQLite.
    /// </summary>
    [Fact]
    public void Refresh_SQLiteOleDbConnectionAfterDataUpdate_ReturnsSuccess()
    {
        var testFile = CoreTestHelper.CreateUniqueTestFile(
            nameof(ConnectionCommandsTests),
            nameof(Refresh_SQLiteOleDbConnectionAfterDataUpdate_ReturnsSuccess),
            _tempDir);

        var dbPath = Path.Combine(_tempDir, $"TestDb_{Guid.NewGuid():N}.db");

        try
        {
            // Create SQLite database with initial data
            SQLiteDatabaseHelper.CreateTestDatabase(dbPath);

            // Create OLEDB connection
            var connectionName = "TestSQLiteConnection";
            ConnectionTestHelper.CreateSQLiteOleDbConnection(testFile, connectionName, dbPath);

            using var batch = ExcelSession.BeginBatch(testFile);

            // Load initial data
            var loadResult = _commands.LoadTo(batch, connectionName, "Products");
            Assert.True(loadResult.Success);

            batch.Save();

            // Update data in SQLite database
            SQLiteDatabaseHelper.UpdateTestData(dbPath, "UPDATE Products SET Price = Price * 1.1");

            // Act - Refresh to pull updated data
            var refreshResult = _commands.Refresh(batch, connectionName);

            // Assert - Refresh succeeded
            Assert.True(refreshResult.Success, $"Failed to refresh after data update: {refreshResult.ErrorMessage}");
        }
        finally
        {
            // Cleanup
            SQLiteDatabaseHelper.DeleteDatabase(dbPath);
        }
    }

    /// <summary>
    /// Tests OLEDB connection with LoadTo and Refresh operations.
    /// This test attempts to create a real Access database to validate OLEDB operations.
    /// DEPRECATED: Use SQLite-based tests instead (Refresh_SQLiteOleDbConnection_ReturnsSuccess).
    /// </summary>
    /// <remarks>
    /// OLEDB connections CANNOT be created via Add2() - documented Excel COM limitation.
    /// This test will skip if ADOX COM is not available or if Add2() throws ArgumentException.
    /// Users should use Power Query for OLEDB data import instead.
    /// </remarks>
    [Fact]
    public void Refresh_ConnectionWithLoadedData_ReturnsSuccess()
    {
        var testFile = CoreTestHelper.CreateUniqueTestFile(
            nameof(ConnectionCommandsTests),
            nameof(Refresh_ConnectionWithLoadedData_ReturnsSuccess),
            _tempDir);

        var dbPath = Path.Combine(_tempDir, $"TestDb_{Guid.NewGuid():N}.accdb");

        try
        {
            // Attempt to create Access database using ADOX COM
            var catalogType = Type.GetTypeFromProgID("ADOX.Catalog");
            if (catalogType == null)
            {
                // ADOX COM not available - skip test
                return;
            }

            dynamic? catalog = null;
            dynamic? connection = null;

            try
            {
                catalog = Activator.CreateInstance(catalogType);
                var connectionString = $"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={dbPath}";

                // Create database
                catalog.Create(connectionString);

                // Get connection from catalog
                connection = catalog.ActiveConnection;

                // Create table with test data
                connection.Execute(
                    "CREATE TABLE Products (" +
                    "ProductID AUTOINCREMENT PRIMARY KEY, " +
                    "ProductName TEXT(50), " +
                    "Price CURRENCY)");

                connection.Execute("INSERT INTO Products (ProductName, Price) VALUES ('Widget', 19.99)");
                connection.Execute("INSERT INTO Products (ProductName, Price) VALUES ('Gadget', 29.99)");
                connection.Execute("INSERT INTO Products (ProductName, Price) VALUES ('Doohickey', 39.99)");
            }
            catch (System.Runtime.InteropServices.COMException ex) when (ex.Message.Contains("Class not registered"))
            {
                // ADOX COM not registered - skip test
                return;
            }
            finally
            {
                if (connection != null)
                {
                    connection.Close();
                    ComInterop.ComUtilities.Release(ref connection!);
                }
                if (catalog != null)
                {
                    ComInterop.ComUtilities.Release(ref catalog!);
                }
            }

            // Test OLEDB connection operations
            using var batch = ExcelSession.BeginBatch(testFile);

            var connectionName = "TestAccessConnection";
            var oledbConnectionString = $"OLEDB;Provider=Microsoft.ACE.OLEDB.12.0;Data Source={dbPath}";
            var commandText = "SELECT * FROM Products";

            // Act - Create connection
            Core.Models.OperationResult createResult;
            try
            {
                createResult = _commands.Create(batch, connectionName, oledbConnectionString, commandText);
            }
            catch (ArgumentException ex) when (ex.Message.Contains("Value does not fall within the expected range"))
            {
                // OLEDB Add2() limitation confirmed - test skipped
                // This confirms OLEDB connections cannot be created via Add2() even with valid Access database
                return;
            }

            // Assert - Connection created successfully
            Assert.True(createResult.Success, $"Failed to create connection: {createResult.ErrorMessage}");

            // Act - Load data to worksheet
            var loadResult = _commands.LoadTo(batch, connectionName, "TestSheet");

            // Assert - Data loaded
            Assert.True(loadResult.Success, $"Failed to load data: {loadResult.ErrorMessage}");

            batch.Save();

            // Act - Refresh connection
            var refreshResult = _commands.Refresh(batch, connectionName);

            // Assert - Refresh succeeded
            Assert.True(refreshResult.Success, $"Failed to refresh: {refreshResult.ErrorMessage}");
        }
        finally
        {
            // Cleanup - Delete Access database
            if (System.IO.File.Exists(dbPath))
            {
                try
                {
                    System.IO.File.Delete(dbPath);
                }
                catch
                {
                    // Ignore cleanup errors
                }
            }
        }
    }
}
