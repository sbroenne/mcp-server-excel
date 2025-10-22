using Sbroenne.ExcelMcp.Core.Commands;
using System.Runtime.InteropServices;

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
        ExcelHelper.WithExcel(filePath, save: true, (excel, workbook) =>
        {
            try
            {
                // Get connections collection
                dynamic connections = workbook.Connections;

                // Create OLEDB connection using positional parameters
                // Connections.Add(Name, Description, ConnectionString, CommandText)
                dynamic newConnection = connections.Add(
                    connectionName,
                    $"Test OLEDB connection created by {nameof(CreateOleDbConnection)}",
                    connectionString,
                    ""
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
        ExcelHelper.WithExcel(filePath, save: true, (excel, workbook) =>
        {
            try
            {
                dynamic connections = workbook.Connections;

                // Create ODBC connection using positional parameters
                dynamic newConnection = connections.Add(
                    connectionName,
                    $"Test ODBC connection created by {nameof(CreateOdbcConnection)}",
                    connectionString,
                    ""
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
        ExcelHelper.WithExcel(filePath, save: true, (excel, workbook) =>
        {
            try
            {
                // Ensure text file exists
                if (!File.Exists(textFilePath))
                {
                    // Create a simple CSV file for testing
                    File.WriteAllText(textFilePath, "Column1,Column2,Column3\nValue1,Value2,Value3\nTest1,Test2,Test3");
                }

                dynamic connections = workbook.Connections;

                // Create text file connection using the SAME approach as Import
                string connectionString = $"TEXT;{textFilePath}";

                // Use Connections.Add() with named parameters like Import does
                dynamic newConnection = connections.Add(
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
    }

    /// <summary>
    /// Creates a simple web connection (URL) in an Excel workbook.
    /// Uses the same approach as ConnectionCommands.Import() for consistency.
    /// </summary>
    public static void CreateWebConnection(string filePath, string connectionName, string url)
    {
        ExcelHelper.WithExcel(filePath, save: true, (excel, workbook) =>
        {
            try
            {
                dynamic connections = workbook.Connections;

                // Create web connection using the SAME approach as Import and CreateTextFileConnection
                // Use URL; prefix in connection string to indicate web connection type
                string connectionString = $"URL;{url}";

                // Use Connections.Add() with named parameters like Import does
                dynamic newConnection = connections.Add(
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
    }

    /// <summary>
    /// Creates multiple test connections of different types for multi-connection tests.
    /// </summary>
    public static void CreateMultipleConnections(string filePath, params (string name, string type, string connectionString)[] connections)
    {
        ExcelHelper.WithExcel(filePath, save: true, (excel, workbook) =>
        {
            try
            {
                dynamic connectionsCollection = workbook.Connections;

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

                return 0; // Success
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException($"Failed to create multiple connections: {ex.Message}", ex);
            }
        });
    }
}
