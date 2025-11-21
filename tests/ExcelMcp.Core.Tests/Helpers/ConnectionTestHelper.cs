using Sbroenne.ExcelMcp.ComInterop.Session;

namespace Sbroenne.ExcelMcp.Core.Tests.Helpers;

/// <summary>
/// Helper class to create test Excel connections via COM interop.
/// Used by integration and round trip tests.
/// </summary>
public static class ConnectionTestHelper
{
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

                // Create OLEDB connection using NAMED parameters (Excel COM requires this)
                // Per Microsoft docs: https://learn.microsoft.com/en-us/office/vba/api/excel.connections.add
                dynamic newConnection = connections.Add(
                    Name: connectionName,
                    Description: $"Test OLEDB connection created by {nameof(CreateOleDbConnection)}",
                    ConnectionString: connectionString,
                    CommandText: ""
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
    /// Creates a text file connection in an Excel workbook.
    /// This is useful for testing connection refresh and data loading.
    /// </summary>
    public static void CreateTextFileConnection(string filePath, string connectionName, string textFilePath)
    {
        using var batch = ExcelSession.BeginBatch(filePath);
        batch.Execute((ctx, ct) =>
        {
            try
            {
                // Ensure text file exists
                if (!File.Exists(textFilePath))
                {
                    // Create a simple CSV file for testing
                    File.WriteAllText(textFilePath, "Column1,Column2,Column3\nValue1,Value2,Value3\nTest1,Test2,Test3");
                }

                dynamic connections = ctx.Book.Connections;

                // Create text file connection using the SAME approach as Import
                string connectionString = $"TEXT;{textFilePath}";

                // Use Connections.Add() with named parameters like Import does
                connections.Add(
                    Name: connectionName,
                    Description: $"Test text file connection created by {nameof(CreateTextFileConnection)}",
                    ConnectionString: connectionString,
                    CommandText: ""
                );

                // Connection created - Excel should handle the rest
                // If Import works, this should too

                return 0; // Success
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException($"Failed to create text file connection '{connectionName}': {ex.Message}", ex);
            }
        });
        batch.Save();
    }

    /// <summary>
    /// Creates a simple web connection (URL) in an Excel workbook.
    /// Uses the same approach as ConnectionCommands.Import() for consistency.
    /// </summary>
    public static void CreateWebConnection(string filePath, string connectionName, string url)
    {
        using var batch = ExcelSession.BeginBatch(filePath);
        batch.Execute((ctx, ct) =>
        {
            try
            {
                dynamic connections = ctx.Book.Connections;

                // Create web connection using the SAME approach as Import and CreateTextFileConnection
                // Use URL; prefix in connection string to indicate web connection type
                string connectionString = $"URL;{url}";

                // Use Connections.Add() with named parameters like Import does
                connections.Add(
                    Name: connectionName,
                    Description: $"Test web connection created by {nameof(CreateWebConnection)}",
                    ConnectionString: connectionString,
                    CommandText: ""
                );

                // Connection created - Excel should handle the rest
                // With the URL; prefix, Excel should recognize this as a Web connection (type 5)

                return 0; // Success
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException($"Failed to create web connection '{connectionName}': {ex.Message}", ex);
            }
        });
        batch.Save();
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
