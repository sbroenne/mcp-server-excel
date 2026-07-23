using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Models;

namespace Sbroenne.ExcelMcp.Core.Commands.Table;

/// <summary>
/// Table Data Model integration operations (AddToDataModel)
/// </summary>
public partial class TableCommands
{
    /// <inheritdoc />
    public AddToDataModelResult AddToDataModel(IExcelBatch batch, string tableName, bool stripBracketColumnNames = false)
    {
        // Security: Validate table name
        ValidateTableName(tableName);

        return batch.Execute((ctx, ct) =>
        {
            dynamic? table = null;
            dynamic? model = null;
            dynamic? modelTables = null;
            try
            {
                table = FindTable(ctx.Book, tableName);

                // Data Model is always available in Excel 2013+ (no need to check)
                model = ctx.Book.Model;
                modelTables = model.ModelTables;

                // Check if table is already in the Data Model via ModelTables
                for (int i = 1; i <= modelTables.Count; i++)
                {
                    dynamic? modelTable = null;
                    try
                    {
                        modelTable = modelTables.Item(i);
                        string sourceTableName = modelTable.SourceName;
                        if (sourceTableName == tableName || sourceTableName.EndsWith($"[{tableName}]", StringComparison.Ordinal))
                        {
                            throw new InvalidOperationException($"Table '{tableName}' is already in the Data Model");
                        }
                    }
                    finally
                    {
                        ComUtilities.Release(ref modelTable);
                    }
                }

                // Detect and optionally strip bracket-escaped column names
                // Columns with literal brackets (e.g., [ACR_CM1]) cannot be referenced in DAX
                var bracketColumnNames = FindBracketColumnNames(table);
                string[]? bracketColumnsFound = null;
                string[]? bracketColumnsRenamed = null;

                if (bracketColumnNames.Count > 0)
                {
                    if (stripBracketColumnNames)
                    {
                        // Rename columns in source table to remove brackets before adding to Data Model
                        StripBracketColumnNames(table, bracketColumnNames);
                        bracketColumnsRenamed = bracketColumnNames.ToArray();
                    }
                    else
                    {
                        bracketColumnsFound = bracketColumnNames.ToArray();
                    }
                }

                // Create a connection for the table using the sigma_coding VBA pattern
                // ConnectionString: "WORKSHEET;{DirectoryPath}" (directory only, not full file path!)
                // CommandText: "{WorkbookName}!{TableName}" (not SQL query!)
                // lCmdtype: xlCmdExcel = 7 (THE KEY - not 4 or 8!)
                const int xlCmdExcel = 7;
                string connectionName = $"WorkbookConnection_{ctx.Book.Name}!{tableName}";

                // Add table to Data Model using sigma_coding VBA pattern
                dynamic? workbookConnections = null;
                dynamic? newConnection = null;
                try
                {
                    workbookConnections = ctx.Book.Connections;

                    // Double-check: Connection name shouldn't exist
                    for (int i = 1; i <= workbookConnections.Count; i++)
                    {
                        dynamic? conn = null;
                        try
                        {
                            conn = workbookConnections.Item(i);
                            if (conn.Name == connectionName)
                            {
                                throw new InvalidOperationException($"Table '{tableName}' is already in the Data Model");
                            }
                        }
                        finally
                        {
                            ComUtilities.Release(ref conn);
                        }
                    }

                    // Create the connection using EXACT pattern from sigma_coding VBA
                    newConnection = workbookConnections.Add2(
                        connectionName,                          // Name
                        $"Excel Table: {tableName}",             // Description
                        $"WORKSHEET;{ctx.Book.Path}",            // ConnectionString: "WORKSHEET;{DirectoryPath}"
                        $"{ctx.Book.Name}!{tableName}",          // CommandText: "{WorkbookName}!{TableName}"
                        xlCmdExcel,                              // lCmdtype: 7 (THE CRITICAL DIFFERENCE!)
                        true,                                    // CreateModelConnection: true
                        false                                    // ImportRelationships: false
                    );
                }
                finally
                {
                    ComUtilities.Release(ref newConnection);
                    ComUtilities.Release(ref workbookConnections);
                }

                // Table is immediately available in Data Model - no refresh needed
                // Connections.Add2() makes the table accessible for relationships/measures instantly

                var result = new AddToDataModelResult { Success = true, FilePath = batch.WorkbookPath };
                if (bracketColumnsFound?.Length > 0)
                    result.BracketColumnsFound = bracketColumnsFound;
                if (bracketColumnsRenamed?.Length > 0)
                    result.BracketColumnsRenamed = bracketColumnsRenamed;
                return result;
            }
            finally
            {
                // Release COM objects
                ComUtilities.Release(ref modelTables);
                ComUtilities.Release(ref model);
                ComUtilities.Release(ref table);
            }
        });
    }

    /// <summary>
    /// Finds column names in the Excel Table that contain literal bracket characters.
    /// Such columns cannot be referenced in DAX formulas after being added to the Data Model.
    /// </summary>
    private static List<string> FindBracketColumnNames(dynamic table)
    {
        var bracketColumns = new List<string>();
        dynamic? listColumns = null;
        try
        {
            listColumns = table.ListColumns;
            int count = listColumns.Count;
            for (int i = 1; i <= count; i++)
            {
                dynamic? col = null;
                try
                {
                    col = listColumns.Item(i);
                    string name = col.Name?.ToString() ?? string.Empty;
                    if (name.Contains('[') || name.Contains(']'))
                    {
                        bracketColumns.Add(name);
                    }
                }
                finally
                {
                    ComUtilities.Release(ref col);
                }
            }
        }
        finally
        {
            ComUtilities.Release(ref listColumns);
        }
        return bracketColumns;
    }

    /// <summary>
    /// Strips literal bracket characters from the names of columns in the source Excel Table.
    /// Modifies the worksheet table column headers in place.
    /// </summary>
    private static void StripBracketColumnNames(dynamic table, List<string> bracketColumnNames)
    {
        dynamic? listColumns = null;
        try
        {
            listColumns = table.ListColumns;
            int count = listColumns.Count;
            for (int i = 1; i <= count; i++)
            {
                dynamic? col = null;
                try
                {
                    col = listColumns.Item(i);
                    string name = col.Name?.ToString() ?? string.Empty;
                    if (bracketColumnNames.Contains(name))
                    {
                        string stripped = name.Replace("[", string.Empty).Replace("]", string.Empty);
                        if (!string.IsNullOrWhiteSpace(stripped))
                        {
                            col.Name = stripped;
                        }
                    }
                }
                finally
                {
                    ComUtilities.Release(ref col);
                }
            }
        }
        finally
        {
            ComUtilities.Release(ref listColumns);
        }
    }
}



