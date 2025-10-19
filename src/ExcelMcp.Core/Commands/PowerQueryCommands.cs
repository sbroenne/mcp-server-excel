using Sbroenne.ExcelMcp.Core.Models;
using static Sbroenne.ExcelMcp.Core.ExcelHelper;

namespace Sbroenne.ExcelMcp.Core.Commands;

/// <summary>
/// Power Query management commands - Core data layer (no console output)
/// </summary>
public class PowerQueryCommands : IPowerQueryCommands
{
    /// <summary>
    /// Finds the closest matching string using simple Levenshtein distance
    /// </summary>
    private static string? FindClosestMatch(string target, List<string> candidates)
    {
        if (candidates.Count == 0) return null;
        
        int minDistance = int.MaxValue;
        string? bestMatch = null;
        
        foreach (var candidate in candidates)
        {
            int distance = ComputeLevenshteinDistance(target.ToLowerInvariant(), candidate.ToLowerInvariant());
            if (distance < minDistance && distance <= Math.Max(target.Length, candidate.Length) / 2)
            {
                minDistance = distance;
                bestMatch = candidate;
            }
        }
        
        return bestMatch;
    }
    
    /// <summary>
    /// Computes Levenshtein distance between two strings
    /// </summary>
    private static int ComputeLevenshteinDistance(string s1, string s2)
    {
        int[,] d = new int[s1.Length + 1, s2.Length + 1];
        
        for (int i = 0; i <= s1.Length; i++)
            d[i, 0] = i;
        for (int j = 0; j <= s2.Length; j++)
            d[0, j] = j;
            
        for (int i = 1; i <= s1.Length; i++)
        {
            for (int j = 1; j <= s2.Length; j++)
            {
                int cost = s1[i - 1] == s2[j - 1] ? 0 : 1;
                d[i, j] = Math.Min(Math.Min(d[i - 1, j] + 1, d[i, j - 1] + 1), d[i - 1, j - 1] + cost);
            }
        }
        
        return d[s1.Length, s2.Length];
    }

    /// <inheritdoc />
    public PowerQueryListResult List(string filePath)
    {
        var result = new PowerQueryListResult { FilePath = filePath };

        if (!File.Exists(filePath))
        {
            result.Success = false;
            result.ErrorMessage = $"File not found: {filePath}";
            return result;
        }

        WithExcel(filePath, false, (excel, workbook) =>
        {
            try
            {
                dynamic queriesCollection = workbook.Queries;
                int count = queriesCollection.Count;
                
                for (int i = 1; i <= count; i++)
                {
                    try
                    {
                        dynamic query = queriesCollection.Item(i);
                        string name = query.Name ?? $"Query{i}";
                        string formula = query.Formula ?? "";
                        
                        string preview = formula.Length > 80 ? formula[..77] + "..." : formula;
                        
                        // Check if connection only
                        bool isConnectionOnly = true;
                        try
                        {
                            dynamic connections = workbook.Connections;
                            for (int c = 1; c <= connections.Count; c++)
                            {
                                dynamic conn = connections.Item(c);
                                string connName = conn.Name?.ToString() ?? "";
                                if (connName.Equals(name, StringComparison.OrdinalIgnoreCase) ||
                                    connName.Equals($"Query - {name}", StringComparison.OrdinalIgnoreCase))
                                {
                                    isConnectionOnly = false;
                                    break;
                                }
                            }
                        }
                        catch { }
                        
                        result.Queries.Add(new PowerQueryInfo
                        {
                            Name = name,
                            Formula = formula,
                            FormulaPreview = preview,
                            IsConnectionOnly = isConnectionOnly
                        });
                    }
                    catch (Exception queryEx)
                    {
                        result.Queries.Add(new PowerQueryInfo
                        {
                            Name = $"Error Query {i}",
                            Formula = "",
                            FormulaPreview = $"Error: {queryEx.Message}",
                            IsConnectionOnly = false
                        });
                    }
                }
                
                result.Success = true;
                return 0;
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.ErrorMessage = $"Error accessing Power Queries: {ex.Message}";
                
                string extension = Path.GetExtension(filePath).ToLowerInvariant();
                if (extension == ".xls")
                {
                    result.ErrorMessage += " (.xls files don't support Power Query)";
                }
                
                return 1;
            }
        });

        return result;
    }

    /// <inheritdoc />
    public PowerQueryViewResult View(string filePath, string queryName)
    {
        var result = new PowerQueryViewResult 
        { 
            FilePath = filePath,
            QueryName = queryName
        };

        if (!File.Exists(filePath))
        {
            result.Success = false;
            result.ErrorMessage = $"File not found: {filePath}";
            return result;
        }

        WithExcel(filePath, false, (excel, workbook) =>
        {
            try
            {
                dynamic query = FindQuery(workbook, queryName);
                if (query == null)
                {
                    var queryNames = GetQueryNames(workbook);
                    string? suggestion = FindClosestMatch(queryName, queryNames);
                    
                    result.Success = false;
                    result.ErrorMessage = $"Query '{queryName}' not found";
                    if (suggestion != null)
                    {
                        result.ErrorMessage += $". Did you mean '{suggestion}'?";
                    }
                    return 1;
                }

                string mCode = query.Formula;
                result.MCode = mCode;
                result.CharacterCount = mCode.Length;
                
                // Check if connection only
                bool isConnectionOnly = true;
                try
                {
                    dynamic connections = workbook.Connections;
                    for (int c = 1; c <= connections.Count; c++)
                    {
                        dynamic conn = connections.Item(c);
                        string connName = conn.Name?.ToString() ?? "";
                        if (connName.Equals(queryName, StringComparison.OrdinalIgnoreCase) ||
                            connName.Equals($"Query - {queryName}", StringComparison.OrdinalIgnoreCase))
                        {
                            isConnectionOnly = false;
                            break;
                        }
                    }
                }
                catch { }
                
                result.IsConnectionOnly = isConnectionOnly;
                result.Success = true;
                return 0;
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.ErrorMessage = $"Error viewing query: {ex.Message}";
                return 1;
            }
        });

        return result;
    }

    /// <inheritdoc />
    public async Task<OperationResult> Update(string filePath, string queryName, string mCodeFile)
    {
        var result = new OperationResult 
        { 
            FilePath = filePath, 
            Action = "pq-update"
        };

        if (!File.Exists(filePath))
        {
            result.Success = false;
            result.ErrorMessage = $"File not found: {filePath}";
            return result;
        }

        if (!File.Exists(mCodeFile))
        {
            result.Success = false;
            result.ErrorMessage = $"M code file not found: {mCodeFile}";
            return result;
        }

        string mCode = await File.ReadAllTextAsync(mCodeFile);

        WithExcel(filePath, true, (excel, workbook) =>
        {
            try
            {
                dynamic query = FindQuery(workbook, queryName);
                if (query == null)
                {
                    var queryNames = GetQueryNames(workbook);
                    string? suggestion = FindClosestMatch(queryName, queryNames);
                    
                    result.Success = false;
                    result.ErrorMessage = $"Query '{queryName}' not found";
                    if (suggestion != null)
                    {
                        result.ErrorMessage += $". Did you mean '{suggestion}'?";
                    }
                    return 1;
                }

                query.Formula = mCode;
                result.Success = true;
                return 0;
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.ErrorMessage = $"Error updating query: {ex.Message}";
                return 1;
            }
        });

        return result;
    }

    /// <inheritdoc />
    public async Task<OperationResult> Export(string filePath, string queryName, string outputFile)
    {
        var result = new OperationResult 
        { 
            FilePath = filePath, 
            Action = "pq-export"
        };

        if (!File.Exists(filePath))
        {
            result.Success = false;
            result.ErrorMessage = $"File not found: {filePath}";
            return result;
        }

        WithExcel(filePath, false, (excel, workbook) =>
        {
            try
            {
                dynamic query = FindQuery(workbook, queryName);
                if (query == null)
                {
                    var queryNames = GetQueryNames(workbook);
                    string? suggestion = FindClosestMatch(queryName, queryNames);
                    
                    result.Success = false;
                    result.ErrorMessage = $"Query '{queryName}' not found";
                    if (suggestion != null)
                    {
                        result.ErrorMessage += $". Did you mean '{suggestion}'?";
                    }
                    return 1;
                }

                string mCode = query.Formula;
                File.WriteAllText(outputFile, mCode);
                
                result.Success = true;
                return 0;
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.ErrorMessage = $"Error exporting query: {ex.Message}";
                return 1;
            }
        });

        return await Task.FromResult(result);
    }

    /// <inheritdoc />
    public async Task<OperationResult> Import(string filePath, string queryName, string mCodeFile)
    {
        var result = new OperationResult 
        { 
            FilePath = filePath, 
            Action = "pq-import"
        };

        if (!File.Exists(filePath))
        {
            result.Success = false;
            result.ErrorMessage = $"File not found: {filePath}";
            return result;
        }

        if (!File.Exists(mCodeFile))
        {
            result.Success = false;
            result.ErrorMessage = $"M code file not found: {mCodeFile}";
            return result;
        }

        string mCode = await File.ReadAllTextAsync(mCodeFile);

        WithExcel(filePath, true, (excel, workbook) =>
        {
            try
            {
                // Check if query already exists
                dynamic existingQuery = FindQuery(workbook, queryName);
                if (existingQuery != null)
                {
                    result.Success = false;
                    result.ErrorMessage = $"Query '{queryName}' already exists. Use pq-update to modify it.";
                    return 1;
                }

                // Add new query
                dynamic queriesCollection = workbook.Queries;
                dynamic newQuery = queriesCollection.Add(queryName, mCode);
                
                result.Success = true;
                return 0;
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.ErrorMessage = $"Error importing query: {ex.Message}";
                return 1;
            }
        });

        return result;
    }

    /// <inheritdoc />
    public OperationResult Refresh(string filePath, string queryName)
    {
        var result = new OperationResult 
        { 
            FilePath = filePath, 
            Action = "pq-refresh"
        };

        if (!File.Exists(filePath))
        {
            result.Success = false;
            result.ErrorMessage = $"File not found: {filePath}";
            return result;
        }

        WithExcel(filePath, true, (excel, workbook) =>
        {
            try
            {
                dynamic query = FindQuery(workbook, queryName);
                if (query == null)
                {
                    var queryNames = GetQueryNames(workbook);
                    string? suggestion = FindClosestMatch(queryName, queryNames);
                    
                    result.Success = false;
                    result.ErrorMessage = $"Query '{queryName}' not found";
                    if (suggestion != null)
                    {
                        result.ErrorMessage += $". Did you mean '{suggestion}'?";
                    }
                    return 1;
                }

                // Check if query has a connection to refresh
                dynamic? targetConnection = null;
                try
                {
                    dynamic connections = workbook.Connections;
                    for (int i = 1; i <= connections.Count; i++)
                    {
                        dynamic conn = connections.Item(i);
                        string connName = conn.Name?.ToString() ?? "";
                        if (connName.Equals(queryName, StringComparison.OrdinalIgnoreCase) ||
                            connName.Equals($"Query - {queryName}", StringComparison.OrdinalIgnoreCase))
                        {
                            targetConnection = conn;
                            break;
                        }
                    }
                }
                catch { }

                if (targetConnection != null)
                {
                    targetConnection.Refresh();
                    result.Success = true;
                }
                else
                {
                    result.Success = true;
                    result.ErrorMessage = "Query is connection-only or function - no data to refresh";
                }
                
                return 0;
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.ErrorMessage = $"Error refreshing query: {ex.Message}";
                return 1;
            }
        });

        return result;
    }

    /// <inheritdoc />
    public PowerQueryViewResult Errors(string filePath, string queryName)
    {
        var result = new PowerQueryViewResult 
        { 
            FilePath = filePath,
            QueryName = queryName
        };

        if (!File.Exists(filePath))
        {
            result.Success = false;
            result.ErrorMessage = $"File not found: {filePath}";
            return result;
        }

        WithExcel(filePath, false, (excel, workbook) =>
        {
            try
            {
                dynamic query = FindQuery(workbook, queryName);
                if (query == null)
                {
                    result.Success = false;
                    result.ErrorMessage = $"Query '{queryName}' not found";
                    return 1;
                }

                // Try to get error information if available
                try
                {
                    dynamic connections = workbook.Connections;
                    for (int i = 1; i <= connections.Count; i++)
                    {
                        dynamic conn = connections.Item(i);
                        string connName = conn.Name?.ToString() ?? "";
                        if (connName.Equals(queryName, StringComparison.OrdinalIgnoreCase) ||
                            connName.Equals($"Query - {queryName}", StringComparison.OrdinalIgnoreCase))
                        {
                            // Connection found - query has been loaded
                            result.MCode = "No error information available through Excel COM interface";
                            result.Success = true;
                            return 0;
                        }
                    }
                }
                catch { }

                result.MCode = "Query is connection-only - no error information available";
                result.Success = true;
                return 0;
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.ErrorMessage = $"Error checking query errors: {ex.Message}";
                return 1;
            }
        });

        return result;
    }

    /// <inheritdoc />
    public OperationResult LoadTo(string filePath, string queryName, string sheetName)
    {
        var result = new OperationResult 
        { 
            FilePath = filePath, 
            Action = "pq-loadto"
        };

        if (!File.Exists(filePath))
        {
            result.Success = false;
            result.ErrorMessage = $"File not found: {filePath}";
            return result;
        }

        WithExcel(filePath, true, (excel, workbook) =>
        {
            try
            {
                dynamic query = FindQuery(workbook, queryName);
                if (query == null)
                {
                    result.Success = false;
                    result.ErrorMessage = $"Query '{queryName}' not found";
                    return 1;
                }

                // Find or create target sheet
                dynamic sheets = workbook.Worksheets;
                dynamic? targetSheet = null;
                
                for (int i = 1; i <= sheets.Count; i++)
                {
                    dynamic sheet = sheets.Item(i);
                    if (sheet.Name == sheetName)
                    {
                        targetSheet = sheet;
                        break;
                    }
                }

                if (targetSheet == null)
                {
                    targetSheet = sheets.Add();
                    targetSheet.Name = sheetName;
                }

                // Load query to worksheet using QueryTables
                dynamic queryTables = targetSheet.QueryTables;
                string connectionString = $"OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location={queryName}";
                string commandText = $"SELECT * FROM [{queryName}]";

                dynamic queryTable = queryTables.Add(connectionString, targetSheet.Range["A1"], commandText);
                queryTable.Name = queryName.Replace(" ", "_");
                queryTable.RefreshStyle = 1; // xlInsertDeleteCells
                queryTable.Refresh(false);

                result.Success = true;
                return 0;
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.ErrorMessage = $"Error loading query to worksheet: {ex.Message}";
                return 1;
            }
        });

        return result;
    }

    /// <inheritdoc />
    public OperationResult Delete(string filePath, string queryName)
    {
        var result = new OperationResult 
        { 
            FilePath = filePath, 
            Action = "pq-delete"
        };

        if (!File.Exists(filePath))
        {
            result.Success = false;
            result.ErrorMessage = $"File not found: {filePath}";
            return result;
        }

        WithExcel(filePath, true, (excel, workbook) =>
        {
            try
            {
                dynamic query = FindQuery(workbook, queryName);
                if (query == null)
                {
                    result.Success = false;
                    result.ErrorMessage = $"Query '{queryName}' not found";
                    return 1;
                }

                dynamic queriesCollection = workbook.Queries;
                queriesCollection.Item(queryName).Delete();

                result.Success = true;
                return 0;
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.ErrorMessage = $"Error deleting query: {ex.Message}";
                return 1;
            }
        });

        return result;
    }

    /// <summary>
    /// Helper to get all query names
    /// </summary>
    private static List<string> GetQueryNames(dynamic workbook)
    {
        var names = new List<string>();
        try
        {
            dynamic queriesCollection = workbook.Queries;
            for (int i = 1; i <= queriesCollection.Count; i++)
            {
                names.Add(queriesCollection.Item(i).Name);
            }
        }
        catch { }
        return names;
    }

    /// <inheritdoc />
    public WorksheetListResult Sources(string filePath)
    {
        var result = new WorksheetListResult { FilePath = filePath };

        if (!File.Exists(filePath))
        {
            result.Success = false;
            result.ErrorMessage = $"File not found: {filePath}";
            return result;
        }

        WithExcel(filePath, false, (excel, workbook) =>
        {
            try
            {
                // Get all tables from all worksheets
                dynamic worksheets = workbook.Worksheets;
                for (int ws = 1; ws <= worksheets.Count; ws++)
                {
                    dynamic worksheet = worksheets.Item(ws);
                    string wsName = worksheet.Name;
                    
                    dynamic tables = worksheet.ListObjects;
                    for (int i = 1; i <= tables.Count; i++)
                    {
                        dynamic table = tables.Item(i);
                        result.Worksheets.Add(new WorksheetInfo
                        {
                            Name = table.Name,
                            Index = i,
                            Visible = true
                        });
                    }
                }

                // Get all named ranges
                dynamic names = workbook.Names;
                int namedRangeIndex = result.Worksheets.Count + 1;
                for (int i = 1; i <= names.Count; i++)
                {
                    dynamic name = names.Item(i);
                    string nameValue = name.Name;
                    if (!nameValue.StartsWith("_"))
                    {
                        result.Worksheets.Add(new WorksheetInfo
                        {
                            Name = nameValue,
                            Index = namedRangeIndex++,
                            Visible = true
                        });
                    }
                }

                result.Success = true;
                return 0;
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.ErrorMessage = $"Error listing sources: {ex.Message}";
                return 1;
            }
        });

        return result;
    }

    /// <inheritdoc />
    public OperationResult Test(string filePath, string sourceName)
    {
        var result = new OperationResult 
        { 
            FilePath = filePath, 
            Action = "pq-test"
        };

        if (!File.Exists(filePath))
        {
            result.Success = false;
            result.ErrorMessage = $"File not found: {filePath}";
            return result;
        }

        WithExcel(filePath, true, (excel, workbook) =>
        {
            try
            {
                // Create a test query to load the source
                string testQuery = $@"
let
    Source = Excel.CurrentWorkbook(){{[Name=""{sourceName.Replace("\"", "\"\"")}""]]}}[Content]
in
    Source";

                dynamic queriesCollection = workbook.Queries;
                dynamic tempQuery = queriesCollection.Add("_TestQuery", testQuery);

                // Try to refresh
                bool refreshSuccess = false;
                try
                {
                    tempQuery.Refresh();
                    refreshSuccess = true;
                }
                catch { }

                // Clean up
                tempQuery.Delete();

                result.Success = true;
                if (!refreshSuccess)
                {
                    result.ErrorMessage = "Source exists but could not refresh (may need data source configuration)";
                }
                
                return 0;
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.ErrorMessage = $"Source '{sourceName}' not found or cannot be loaded: {ex.Message}";
                return 1;
            }
        });

        return result;
    }

    /// <inheritdoc />
    public WorksheetDataResult Peek(string filePath, string sourceName)
    {
        var result = new WorksheetDataResult 
        { 
            FilePath = filePath,
            SheetName = sourceName
        };

        if (!File.Exists(filePath))
        {
            result.Success = false;
            result.ErrorMessage = $"File not found: {filePath}";
            return result;
        }

        WithExcel(filePath, false, (excel, workbook) =>
        {
            try
            {
                // Check if it's a named range (single value)
                dynamic names = workbook.Names;
                for (int i = 1; i <= names.Count; i++)
                {
                    dynamic name = names.Item(i);
                    string nameValue = name.Name;
                    if (nameValue == sourceName)
                    {
                        try
                        {
                            var value = name.RefersToRange.Value;
                            result.Data.Add(new List<object?> { value });
                            result.RowCount = 1;
                            result.ColumnCount = 1;
                            result.Success = true;
                            return 0;
                        }
                        catch
                        {
                            result.Success = false;
                            result.ErrorMessage = "Named range found but value cannot be read (may be #REF!)";
                            return 1;
                        }
                    }
                }

                // Check if it's a table
                dynamic worksheets = workbook.Worksheets;
                for (int ws = 1; ws <= worksheets.Count; ws++)
                {
                    dynamic worksheet = worksheets.Item(ws);
                    dynamic tables = worksheet.ListObjects;
                    for (int i = 1; i <= tables.Count; i++)
                    {
                        dynamic table = tables.Item(i);
                        if (table.Name == sourceName)
                        {
                            result.RowCount = table.ListRows.Count;
                            result.ColumnCount = table.ListColumns.Count;

                            // Get column names
                            dynamic listCols = table.ListColumns;
                            for (int c = 1; c <= Math.Min(result.ColumnCount, 10); c++)
                            {
                                result.Headers.Add(listCols.Item(c).Name);
                            }

                            result.Success = true;
                            return 0;
                        }
                    }
                }

                result.Success = false;
                result.ErrorMessage = $"Source '{sourceName}' not found";
                return 1;
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.ErrorMessage = $"Error peeking source: {ex.Message}";
                return 1;
            }
        });

        return result;
    }

    /// <inheritdoc />
    public PowerQueryViewResult Eval(string filePath, string mExpression)
    {
        var result = new PowerQueryViewResult 
        { 
            FilePath = filePath,
            QueryName = "_EvalExpression"
        };

        if (!File.Exists(filePath))
        {
            result.Success = false;
            result.ErrorMessage = $"File not found: {filePath}";
            return result;
        }

        WithExcel(filePath, true, (excel, workbook) =>
        {
            try
            {
                // Create a temporary query with the expression
                string evalQuery = $@"
let
    Result = {mExpression}
in
    Result";

                dynamic queriesCollection = workbook.Queries;
                dynamic tempQuery = queriesCollection.Add("_EvalQuery", evalQuery);

                result.MCode = evalQuery;
                result.CharacterCount = evalQuery.Length;

                // Try to refresh
                try
                {
                    tempQuery.Refresh();
                    result.Success = true;
                    result.ErrorMessage = null;
                }
                catch (Exception refreshEx)
                {
                    result.Success = false;
                    result.ErrorMessage = $"Expression syntax is valid but refresh failed: {refreshEx.Message}";
                }

                // Clean up
                tempQuery.Delete();

                return 0;
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.ErrorMessage = $"Expression evaluation failed: {ex.Message}";
                return 1;
            }
        });

        return result;
    }
}
