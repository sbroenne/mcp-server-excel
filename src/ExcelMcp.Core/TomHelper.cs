using Microsoft.AnalysisServices.Tabular;

namespace Sbroenne.ExcelMcp.Core;

/// <summary>
/// Helper utilities for TOM (Tabular Object Model) API operations.
/// Provides connection management and common TOM operations for Excel Data Models.
/// </summary>
public static class TomHelper
{
    /// <summary>
    /// Executes an action within a TOM server connection context.
    /// Automatically handles connection, cleanup, and error handling.
    /// </summary>
    /// <typeparam name="T">Return type of the operation</typeparam>
    /// <param name="filePath">Path to Excel file with Data Model</param>
    /// <param name="action">Action to execute with server and model access</param>
    /// <param name="saveChanges">Whether to save changes to the model (default: false)</param>
    /// <returns>Result of the action</returns>
    public static T WithTomServer<T>(string filePath, Func<Server, Model, T> action, bool saveChanges = false)
    {
        Server? server = null;
        try
        {
            server = new Server();
            
            // Try different connection string formats for Excel compatibility
            string[] connectionFormats = new[]
            {
                $"Provider=MSOLAP;Data Source={filePath};",
                $"Data Source={filePath};",
                $"Provider=MSOLAP.8;Data Source={filePath};",
                $"DataSource={filePath};Provider=MSOLAP;"
            };

            Exception? lastException = null;
            foreach (var connString in connectionFormats)
            {
                try
                {
                    server.Connect(connString);
                    
                    if (server.Connected)
                    {
                        if (server.Databases.Count == 0)
                        {
                            throw new InvalidOperationException(
                                "Connected to Excel file but no Data Model database found. " +
                                "Ensure the file has Power Pivot / Data Model enabled.");
                        }

                        Database db = server.Databases[0];
                        Model model = db.Model;

                        if (model == null)
                        {
                            throw new InvalidOperationException(
                                "Connected to database but no model found. " +
                                "Ensure the Data Model is properly initialized.");
                        }

                        // Execute the action
                        var result = action(server, model);

                        // Save changes if requested
                        if (saveChanges)
                        {
                            model.SaveChanges();
                        }

                        return result;
                    }
                }
                catch (Exception ex)
                {
                    lastException = ex;
                    if (server.Connected)
                    {
                        server.Disconnect();
                    }
                }
            }

            // If we get here, all connection attempts failed
            throw new InvalidOperationException(
                $"Could not connect to Excel Data Model at '{filePath}'. " +
                "Ensure the file exists, has Data Model enabled, and is not locked by Excel. " +
                $"Last error: {lastException?.Message}",
                lastException);
        }
        finally
        {
            if (server?.Connected == true)
            {
                try
                {
                    server.Disconnect();
                }
                catch
                {
                    // Best effort cleanup
                }
            }
        }
    }

    /// <summary>
    /// Validates a DAX formula syntax without executing it.
    /// </summary>
    /// <param name="filePath">Path to Excel file with Data Model</param>
    /// <param name="daxFormula">DAX formula to validate</param>
    /// <returns>True if valid, false otherwise with error message</returns>
    public static (bool IsValid, string? ErrorMessage) ValidateDaxFormula(string filePath, string? daxFormula)
    {
        try
        {
            return WithTomServer(filePath, (server, model) =>
            {
                // Basic validation - check for common issues
                if (string.IsNullOrWhiteSpace(daxFormula))
                {
                    return (false, "DAX formula cannot be empty");
                }

                // Check for balanced brackets
                int openParens = daxFormula.Count(c => c == '(');
                int closeParens = daxFormula.Count(c => c == ')');
                if (openParens != closeParens)
                {
                    return (false, $"Unbalanced parentheses: {openParens} open, {closeParens} close");
                }

                int openSquare = daxFormula.Count(c => c == '[');
                int closeSquare = daxFormula.Count(c => c == ']');
                if (openSquare != closeSquare)
                {
                    return (false, $"Unbalanced square brackets: {openSquare} open, {closeSquare} close");
                }

                // TOM API will validate the formula when we try to create/update a measure
                // Additional validation happens during SaveChanges()
                return (true, (string?)null);
            }, saveChanges: false);
        }
        catch (Exception ex)
        {
            return (false, $"Validation error: {ex.Message}");
        }
    }

    /// <summary>
    /// Finds a table in the model by name (case-insensitive).
    /// </summary>
    public static Table? FindTable(Model model, string tableName)
    {
        return model.Tables.FirstOrDefault(t => 
            t.Name.Equals(tableName, StringComparison.OrdinalIgnoreCase));
    }

    /// <summary>
    /// Finds a measure in the model by name (case-insensitive).
    /// </summary>
    public static Measure? FindMeasure(Model model, string measureName)
    {
        foreach (var table in model.Tables)
        {
            var measure = table.Measures.FirstOrDefault(m => 
                m.Name.Equals(measureName, StringComparison.OrdinalIgnoreCase));
            if (measure != null)
            {
                return measure;
            }
        }
        return null;
    }

    /// <summary>
    /// Finds a column in a table by name (case-insensitive).
    /// </summary>
    public static Column? FindColumn(Table table, string columnName)
    {
        return table.Columns.FirstOrDefault(c => 
            c.Name.Equals(columnName, StringComparison.OrdinalIgnoreCase));
    }

    /// <summary>
    /// Finds a relationship in the model by endpoint columns.
    /// </summary>
    public static SingleColumnRelationship? FindRelationship(
        Model model,
        string fromTableName,
        string fromColumnName,
        string toTableName,
        string toColumnName)
    {
        foreach (var rel in model.Relationships.OfType<SingleColumnRelationship>())
        {
            var fromCol = rel.FromColumn;
            var toCol = rel.ToColumn;

            if (fromCol?.Parent is Table fromTable &&
                toCol?.Parent is Table toTable &&
                fromTable.Name.Equals(fromTableName, StringComparison.OrdinalIgnoreCase) &&
                fromCol.Name.Equals(fromColumnName, StringComparison.OrdinalIgnoreCase) &&
                toTable.Name.Equals(toTableName, StringComparison.OrdinalIgnoreCase) &&
                toCol.Name.Equals(toColumnName, StringComparison.OrdinalIgnoreCase))
            {
                return rel;
            }
        }

        return null;
    }

    /// <summary>
    /// Gets all table names in the model.
    /// </summary>
    public static List<string> GetTableNames(Model model)
    {
        return model.Tables.Select(t => t.Name).ToList();
    }

    /// <summary>
    /// Gets all measure names in the model.
    /// </summary>
    public static List<string> GetMeasureNames(Model model)
    {
        var names = new List<string>();
        foreach (var table in model.Tables)
        {
            names.AddRange(table.Measures.Select(m => m.Name));
        }
        return names;
    }
}
