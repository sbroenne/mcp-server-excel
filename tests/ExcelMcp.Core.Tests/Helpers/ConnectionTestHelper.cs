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
    public static async Task CreateOleDbConnectionAsync(string filePath, string connectionName, string connectionString)
    {
        await using var batch = await ExcelSession.BeginBatchAsync(filePath);
        await batch.ExecuteAsync<int>((ctx, ct) =>
        {
            try
            {
                // Get connections collection
                dynamic connections = ctx.Book.Connections;

                // Create OLEDB connection using NAMED parameters (Excel COM requires this)
                // Per Microsoft docs: https://learn.microsoft.com/en-us/office/vba/api/excel.connections.add
                dynamic newConnection = connections.Add(
                    Name: connectionName,
                    Description: $"Test OLEDB connection created by {nameof(CreateOleDbConnectionAsync)}",
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

                return ValueTask.FromResult(0); // Success
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException($"Failed to create OLEDB connection '{connectionName}': {ex.Message}", ex);
            }
        });
        await batch.SaveAsync();
    }

    /// <summary>
    /// Creates a simple ODBC connection in an Excel workbook.
    /// </summary>
    public static async Task CreateOdbcConnectionAsync(string filePath, string connectionName, string connectionString)
    {
        await using var batch = await ExcelSession.BeginBatchAsync(filePath);
        await batch.ExecuteAsync<int>((ctx, ct) =>
        {
            try
            {
                dynamic connections = ctx.Book.Connections;

                // Create ODBC connection using NAMED parameters (Excel COM requires this)
                dynamic newConnection = connections.Add(
                    Name: connectionName,
                    Description: $"Test ODBC connection created by {nameof(CreateOdbcConnectionAsync)}",
                    ConnectionString: connectionString,
                    CommandText: ""
                );

                return ValueTask.FromResult(0); // Success
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException($"Failed to create ODBC connection '{connectionName}': {ex.Message}", ex);
            }
        });
        await batch.SaveAsync();
    }

    /// <summary>
    /// Creates a text file connection in an Excel workbook.
    /// This is useful for testing connection refresh and data loading.
    /// </summary>
    public static async Task CreateTextFileConnectionAsync(string filePath, string connectionName, string textFilePath)
    {
        await using var batch = await ExcelSession.BeginBatchAsync(filePath);
        await batch.ExecuteAsync<int>((ctx, ct) =>
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
                dynamic newConnection = connections.Add(
                    Name: connectionName,
                    Description: $"Test text file connection created by {nameof(CreateTextFileConnectionAsync)}",
                    ConnectionString: connectionString,
                    CommandText: ""
                );

                // Connection created - Excel should handle the rest
                // If Import works, this should too

                return ValueTask.FromResult(0); // Success
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException($"Failed to create text file connection '{connectionName}': {ex.Message}", ex);
            }
        });
        await batch.SaveAsync();
    }

    /// <summary>
    /// Creates a simple web connection (URL) in an Excel workbook.
    /// Uses the same approach as ConnectionCommands.Import() for consistency.
    /// </summary>
    public static async Task CreateWebConnectionAsync(string filePath, string connectionName, string url)
    {
        await using var batch = await ExcelSession.BeginBatchAsync(filePath);
        await batch.ExecuteAsync<int>((ctx, ct) =>
        {
            try
            {
                dynamic connections = ctx.Book.Connections;

                // Create web connection using the SAME approach as Import and CreateTextFileConnection
                // Use URL; prefix in connection string to indicate web connection type
                string connectionString = $"URL;{url}";

                // Use Connections.Add() with named parameters like Import does
                dynamic newConnection = connections.Add(
                    Name: connectionName,
                    Description: $"Test web connection created by {nameof(CreateWebConnectionAsync)}",
                    ConnectionString: connectionString,
                    CommandText: ""
                );

                // Connection created - Excel should handle the rest
                // With the URL; prefix, Excel should recognize this as a Web connection (type 5)

                return ValueTask.FromResult(0); // Success
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException($"Failed to create web connection '{connectionName}': {ex.Message}", ex);
            }
        });
        await batch.SaveAsync();
    }

    /// <summary>
    /// Creates multiple test connections of different types for multi-connection tests.
    /// </summary>
    public static async Task CreateMultipleConnectionsAsync(string filePath, params (string name, string type, string connectionString)[] connections)
    {
        await using var batch = await ExcelSession.BeginBatchAsync(filePath);
        await batch.ExecuteAsync<int>((ctx, ct) =>
        {
            try
            {
                dynamic connectionsCollection = ctx.Book.Connections;

                foreach (var (name, type, connectionString) in connections)
                {
                    // Use positional parameters
                    dynamic newConnection = connectionsCollection.Add(
                        name,
                        $"Test {type} connection",
                        connectionString,
                        ""
                    );
                }

                return ValueTask.FromResult(0); // Success
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException($"Failed to create multiple connections: {ex.Message}", ex);
            }
        });
        await batch.SaveAsync();
    }
}
