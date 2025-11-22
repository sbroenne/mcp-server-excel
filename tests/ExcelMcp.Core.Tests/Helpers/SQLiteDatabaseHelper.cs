using Microsoft.Data.Sqlite;

namespace Sbroenne.ExcelMcp.Core.Tests.Helpers;

/// <summary>
/// Helper class to create SQLite databases for OLEDB connection testing.
/// SQLite is used as a lightweight, reliable database engine for integration tests.
/// </summary>
public static class SQLiteDatabaseHelper
{
    /// <summary>
    /// Creates a SQLite database with a simple Products table for testing.
    /// </summary>
    /// <param name="dbPath">Full path where the SQLite database file should be created.</param>
    /// <returns>The connection string for the created database.</returns>
    public static string CreateTestDatabase(string dbPath)
    {
        var connectionString = $"Data Source={dbPath}";

        using var conn = new SqliteConnection(connectionString);
        conn.Open();

        // Create Products table
        using var cmd = conn.CreateCommand();
        cmd.CommandText = @"
            CREATE TABLE Products (
                ProductID INTEGER PRIMARY KEY AUTOINCREMENT,
                ProductName TEXT NOT NULL,
                Price REAL NOT NULL,
                InStock INTEGER NOT NULL DEFAULT 1
            )";
        cmd.ExecuteNonQuery();

        // Insert test data
        cmd.CommandText = @"
            INSERT INTO Products (ProductName, Price, InStock) VALUES
            ('Widget', 19.99, 1),
            ('Gadget', 29.99, 1),
            ('Doohickey', 39.99, 0)";
        cmd.ExecuteNonQuery();

        return connectionString;
    }

    /// <summary>
    /// Creates a SQLite database with custom schema and data.
    /// </summary>
    /// <param name="dbPath">Full path where the SQLite database file should be created.</param>
    /// <param name="createTableSql">SQL statement to create table(s).</param>
    /// <param name="insertDataSql">SQL statement to insert test data (optional).</param>
    /// <returns>The connection string for the created database.</returns>
    public static string CreateCustomDatabase(string dbPath, string createTableSql, string? insertDataSql = null)
    {
        var connectionString = $"Data Source={dbPath}";

        using var conn = new SqliteConnection(connectionString);
        conn.Open();

        using var cmd = conn.CreateCommand();

        // Create tables
#pragma warning disable CA2100 // Review SQL queries for security vulnerabilities - Test code uses hardcoded SQL
        cmd.CommandText = createTableSql;
#pragma warning restore CA2100
        cmd.ExecuteNonQuery();

        // Insert data if provided
        if (!string.IsNullOrEmpty(insertDataSql))
        {
#pragma warning disable CA2100 // Review SQL queries for security vulnerabilities - Test code uses hardcoded SQL
            cmd.CommandText = insertDataSql;
#pragma warning restore CA2100
            cmd.ExecuteNonQuery();
        }

        return connectionString;
    }

    /// <summary>
    /// Gets the OLEDB connection string for a SQLite database.
    /// This format is used by Excel OLEDB connections.
    /// </summary>
    /// <param name="dbPath">Full path to the SQLite database file.</param>
    /// <returns>OLEDB connection string in Excel format.</returns>
    public static string GetOleDbConnectionString(string dbPath)
    {
        // Excel OLEDB connection string format for SQLite
        // Note: Requires SQLite OLEDB provider to be installed on the system
        return $"OLEDB;Provider=System.Data.SQLite;Data Source={dbPath}";
    }

    /// <summary>
    /// Updates data in the SQLite database to simulate data source changes.
    /// Useful for testing connection refresh operations.
    /// </summary>
    /// <param name="dbPath">Full path to the SQLite database file.</param>
    /// <param name="updateSql">SQL UPDATE statement to execute.</param>
    public static void UpdateTestData(string dbPath, string updateSql)
    {
        var connectionString = $"Data Source={dbPath}";

        using var conn = new SqliteConnection(connectionString);
        conn.Open();

        using var cmd = conn.CreateCommand();
#pragma warning disable CA2100 // Review SQL queries for security vulnerabilities - Test code uses hardcoded SQL
        cmd.CommandText = updateSql;
#pragma warning restore CA2100
        cmd.ExecuteNonQuery();
    }

    /// <summary>
    /// Verifies that a SQLite database exists and is accessible.
    /// </summary>
    /// <param name="dbPath">Full path to the SQLite database file.</param>
    /// <returns>True if the database exists and can be opened.</returns>
    public static bool DatabaseExists(string dbPath)
    {
        if (!File.Exists(dbPath))
            return false;

        try
        {
            var connectionString = $"Data Source={dbPath}";
            using var conn = new SqliteConnection(connectionString);
            conn.Open();
            return true;
        }
        catch
        {
            return false;
        }
    }

    /// <summary>
    /// Cleans up a SQLite database file if it exists.
    /// </summary>
    /// <param name="dbPath">Full path to the SQLite database file.</param>
    public static void DeleteDatabase(string dbPath)
    {
        if (File.Exists(dbPath))
        {
            try
            {
                File.Delete(dbPath);
            }
            catch
            {
                // Ignore cleanup errors
            }
        }
    }
}
