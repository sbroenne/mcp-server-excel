using Sbroenne.ExcelMcp.ComInterop.Session;

namespace Sbroenne.ExcelMcp.Core.Tests.Helpers;

/// <summary>
/// Helper class to create test Excel connections via COM interop.
/// Used by integration and round trip tests.
/// </summary>
public static class ConnectionTestHelper
{
    /// <summary>
    /// Creates an OLEDB connection to a SQLite database in an Excel workbook.
    /// This creates an actual Excel connection object that can be managed by ConnectionCommands.
    /// SQLite is used as a reliable, lightweight database for testing OLEDB operations.
    /// </summary>
    public static void CreateSQLiteOleDbConnection(string excelFilePath, string connectionName, string sqliteDbPath)
    {
        // Create SQLite database if it doesn't exist
        if (!File.Exists(sqliteDbPath))
        {
            SQLiteDatabaseHelper.CreateTestDatabase(sqliteDbPath);
        }

        // Get OLEDB connection string for SQLite
        var oledbConnectionString = SQLiteDatabaseHelper.GetOleDbConnectionString(sqliteDbPath);

        using var batch = ExcelSession.BeginBatch(excelFilePath);
        batch.Execute((ctx, ct) =>
        {
            try
            {
                // Get connections collection
                dynamic connections = ctx.Book.Connections;

                // Create OLEDB connection using Add2() (current method, Add() is deprecated)
                // Per instructions: Must use Connections.Add2() for OLEDB/ODBC connections
                dynamic newConnection = connections.Add2(
                    Name: connectionName,
                    Description: $"Test SQLite OLEDB connection created by {nameof(CreateSQLiteOleDbConnection)}",
                    ConnectionString: oledbConnectionString,
                    CommandText: "",
                    lCmdtype: Type.Missing,            // Let Excel auto-detect
                    CreateModelConnection: false,       // Don't create Data Model connection
                    ImportRelationships: false          // Don't import relationships
                );

                // Configure OLEDB connection properties
                if (newConnection.Type == 1) // OLEDB
                {
                    dynamic oledb = newConnection.OLEDBConnection;
                    if (oledb != null)
                    {
                        oledb.BackgroundQuery = true;
                        oledb.RefreshOnFileOpen = false;
                        oledb.SavePassword = false;
                    }
                }

                return 0; // Success
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException($"Failed to create SQLite OLEDB connection '{connectionName}': {ex.Message}", ex);
            }
        });
    }

    /// <summary>
    /// Creates a simple OLEDB connection to a test database in an Excel workbook.
    /// This creates an actual Excel connection object that can be managed by ConnectionCommands.
    /// </summary>
    public static void CreateOleDbConnection(string filePath, string connectionName, string connectionString)
    {
        using var batch = ExcelSession.BeginBatch(filePath);
        batch.Execute((ctx, ct) =>
        {
            try
            {
                // Get connections collection
                dynamic connections = ctx.Book.Connections;

                // Create OLEDB connection using Add2() (current method, Add() is deprecated)
                // Per instructions: Must use Connections.Add2() for OLEDB/ODBC connections
                dynamic newConnection = connections.Add2(
                    Name: connectionName,
                    Description: $"Test OLEDB connection created by {nameof(CreateOleDbConnection)}",
                    ConnectionString: connectionString,
                    CommandText: "",
                    lCmdtype: Type.Missing,            // Let Excel auto-detect
                    CreateModelConnection: false,       // Don't create Data Model connection
                    ImportRelationships: false          // Don't import relationships
                );

                // Configure OLEDB connection properties
                if (newConnection.Type == 1) // OLEDB
                {
                    dynamic oledb = newConnection.OLEDBConnection;
                    if (oledb != null)
                    {
                        oledb.BackgroundQuery = true;
                        oledb.RefreshOnFileOpen = false;
                        oledb.SavePassword = false;
                    }
                }

                return 0; // Success
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException($"Failed to create OLEDB connection '{connectionName}': {ex.Message}", ex);
            }
        });
    }

    /// <summary>
    /// Creates a simple ODBC connection in an Excel workbook.
    /// </summary>
    public static void CreateOdbcConnection(string filePath, string connectionName, string connectionString)
    {
        using var batch = ExcelSession.BeginBatch(filePath);
        batch.Execute((ctx, ct) =>
        {
            try
            {
                dynamic connections = ctx.Book.Connections;

                // Create ODBC connection using NAMED parameters (Excel COM requires this)
                connections.Add(
                    Name: connectionName,
                    Description: $"Test ODBC connection created by {nameof(CreateOdbcConnection)}",
                    ConnectionString: connectionString,
                    CommandText: ""
                );

                return 0; // Success
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException($"Failed to create ODBC connection '{connectionName}': {ex.Message}", ex);
            }
        });
    }



    /// <summary>
    /// Creates multiple test connections of different types for multi-connection tests.
    /// </summary>
    public static void CreateMultipleConnections(string filePath, params (string name, string type, string connectionString)[] connections)
    {
        using var batch = ExcelSession.BeginBatch(filePath);
        batch.Execute((ctx, ct) =>
        {
            try
            {
                dynamic connectionsCollection = ctx.Book.Connections;

                foreach (var (name, type, connectionString) in connections)
                {
                    // Use positional parameters
                    connectionsCollection.Add(
                        name,
                        $"Test {type} connection",
                        connectionString,
                        ""
                    );
                }

                return 0; // Success
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException($"Failed to create multiple connections: {ex.Message}", ex);
            }
        });
        batch.Save();
    }
}
