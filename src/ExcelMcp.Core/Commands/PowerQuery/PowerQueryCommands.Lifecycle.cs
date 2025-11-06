using System.Runtime.InteropServices;
using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Models;
using Sbroenne.ExcelMcp.Core.PowerQuery;
using Sbroenne.ExcelMcp.Core.Security;

namespace Sbroenne.ExcelMcp.Core.Commands;

/// <summary>
/// Power Query lifecycle operations (List, View, Import, Export, Update, Delete)
/// </summary>
public partial class PowerQueryCommands
{
    /// <inheritdoc />
    public async Task<PowerQueryListResult> ListAsync(IExcelBatch batch)
    {
        var result = new PowerQueryListResult { FilePath = batch.WorkbookPath };

        return await batch.Execute((ctx, ct) =>
        {
            dynamic? queriesCollection = null;
            try
            {
                queriesCollection = ctx.Book.Queries;
                int count = queriesCollection.Count;

                for (int i = 1; i <= count; i++)
                {
                    dynamic? query = null;
                    try
                    {
                        query = queriesCollection.Item(i);
                        string name = query.Name ?? $"Query{i}";
                        string formula = query.Formula ?? "";

                        string preview = formula.Length > 80 ? formula[..77] + "..." : formula;

                        // Check if connection only
                        bool isConnectionOnly = true;
                        dynamic? connections = null;
                        try
                        {
                            connections = ctx.Book.Connections;
                            for (int c = 1; c <= connections.Count; c++)
                            {
                                dynamic? conn = null;
                                try
                                {
                                    conn = connections.Item(c);
                                    string connName = conn.Name?.ToString() ?? "";
                                    if (connName.Equals(name, StringComparison.OrdinalIgnoreCase) ||
                                        connName.Equals($"Query - {name}", StringComparison.OrdinalIgnoreCase))
                                    {
                                        isConnectionOnly = false;
                                        break;
                                    }
                                }
                                finally
                                {
                                    ComUtilities.Release(ref conn);
                                }
                            }
                        }
                        catch { }
                        finally
                        {
                            ComUtilities.Release(ref connections);
                        }

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
                    finally
                    {
                        ComUtilities.Release(ref query);
                    }
                }

                result.Success = true;
                return result;
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.ErrorMessage = $"Error accessing Power Queries: {ex.Message}";

                string extension = Path.GetExtension(batch.WorkbookPath).ToLowerInvariant();
                if (extension == ".xls")
                {
                    result.ErrorMessage += " (.xls files don't support Power Query)";
                }

                return result;
            }
            finally
            {
                ComUtilities.Release(ref queriesCollection);
            }
        });
    }

    /// <inheritdoc />
    public async Task<PowerQueryViewResult> ViewAsync(IExcelBatch batch, string queryName)
    {
        var result = new PowerQueryViewResult
        {
            FilePath = batch.WorkbookPath,
            QueryName = queryName
        };

        // Validate query name
        if (!ValidateQueryName(queryName, out string? validationError))
        {
            result.Success = false;
            result.ErrorMessage = validationError;
            return result;
        }

        return await batch.Execute<PowerQueryViewResult>((ctx, ct) =>
        {
            dynamic? query = null;
            try
            {
                query = ComUtilities.FindQuery(ctx.Book, queryName);
                if (query == null)
                {
                    var queryNames = GetQueryNames(ctx.Book);
                    string? suggestion = FindClosestMatch(queryName, queryNames);

                    result.Success = false;
                    result.ErrorMessage = $"Query '{queryName}' not found";
                    if (suggestion != null)
                    {
                        result.ErrorMessage += $". Did you mean '{suggestion}'?";
                    }
                    return result;
                }

                string mCode = query.Formula;
                result.MCode = mCode;
                result.CharacterCount = mCode.Length;

                // Check if connection only
                bool isConnectionOnly = true;
                dynamic? connections = null;
                try
                {
                    connections = ctx.Book.Connections;
                    for (int c = 1; c <= connections.Count; c++)
                    {
                        dynamic? conn = null;
                        try
                        {
                            conn = connections.Item(c);
                            string connName = conn.Name?.ToString() ?? "";
                            if (connName.Equals(queryName, StringComparison.OrdinalIgnoreCase) ||
                                connName.Equals($"Query - {queryName}", StringComparison.OrdinalIgnoreCase))
                            {
                                isConnectionOnly = false;
                                break;
                            }
                        }
                        finally
                        {
                            ComUtilities.Release(ref conn);
                        }
                    }
                }
                catch { }
                finally
                {
                    ComUtilities.Release(ref connections);
                }

                result.IsConnectionOnly = isConnectionOnly;
                result.Success = true;
                return result;
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.ErrorMessage = $"Error viewing query: {ex.Message}";
                return result;
            }
            finally
            {
                ComUtilities.Release(ref query);
            }
        });
    }

    /// <inheritdoc />
    public async Task<OperationResult> ExportAsync(IExcelBatch batch, string queryName, string outputFile)
    {
        var result = new OperationResult
        {
            FilePath = batch.WorkbookPath,
            Action = "pq-export"
        };

        // Validate query name
        if (!ValidateQueryName(queryName, out string? validationError))
        {
            result.Success = false;
            result.ErrorMessage = validationError;
            return result;
        }

        // Validate and normalize the output file path to prevent path traversal attacks
        try
        {
            outputFile = PathValidator.ValidateOutputFile(outputFile, nameof(outputFile), allowOverwrite: true);
        }
        catch (Exception ex)
        {
            result.Success = false;
            result.ErrorMessage = $"Invalid output file path: {ex.Message}";
            return result;
        }

        return await batch.ExecuteAsync<OperationResult>(async (ctx, ct) =>
        {
            dynamic? query = null;
            try
            {
                query = ComUtilities.FindQuery(ctx.Book, queryName);
                if (query == null)
                {
                    var queryNames = GetQueryNames(ctx.Book);
                    string? suggestion = FindClosestMatch(queryName, queryNames);

                    result.Success = false;
                    result.ErrorMessage = $"Query '{queryName}' not found";
                    if (suggestion != null)
                    {
                        result.ErrorMessage += $". Did you mean '{suggestion}'?";
                    }
                    return result;
                }

                string mCode = query.Formula;
                await File.WriteAllTextAsync(outputFile, mCode, ct);

                result.Success = true;
                return result;
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.ErrorMessage = $"Error exporting query: {ex.Message}";
                return result;
            }
            finally
            {
                ComUtilities.Release(ref query);
            }
        });
    }

    /// <inheritdoc />
    public async Task<PowerQueryLoadConfigResult> GetLoadConfigAsync(IExcelBatch batch, string queryName)
    {
        var result = new PowerQueryLoadConfigResult
        {
            FilePath = batch.WorkbookPath,
            QueryName = queryName
        };

        // Validate query name
        if (!ValidateQueryName(queryName, out string? validationError))
        {
            result.Success = false;
            result.ErrorMessage = validationError;
            return result;
        }

        return await batch.Execute<PowerQueryLoadConfigResult>((ctx, ct) =>
        {
            dynamic? query = null;
            dynamic? worksheets = null;
            dynamic? connections = null;
            dynamic? names = null;
            try
            {
                query = ComUtilities.FindQuery(ctx.Book, queryName);
                if (query == null)
                {
                    result.Success = false;
                    result.ErrorMessage = $"Query '{queryName}' not found";
                    return result;
                }

                // Check for QueryTables first (table loading)
                bool hasTableConnection = false;
                bool hasDataModelConnection = false;
                string? targetSheet = null;

                worksheets = ctx.Book.Worksheets;
                for (int ws = 1; ws <= worksheets.Count; ws++)
                {
                    dynamic? worksheet = null;
                    dynamic? queryTables = null;
                    try
                    {
                        worksheet = worksheets.Item(ws);
                        queryTables = worksheet.QueryTables;

                        for (int qt = 1; qt <= queryTables.Count; qt++)
                        {
                            dynamic? queryTable = null;
                            try
                            {
                                queryTable = queryTables.Item(qt);
                                string qtName = queryTable.Name?.ToString() ?? "";

                                // Check if this QueryTable is for our query
                                if (qtName.Equals(queryName.Replace(" ", "_"), StringComparison.OrdinalIgnoreCase) ||
                                    qtName.Contains(queryName.Replace(" ", "_")))
                                {
                                    hasTableConnection = true;
                                    targetSheet = worksheet.Name;
                                    ComUtilities.Release(ref queryTable);
                                    break;
                                }
                            }
                            finally
                            {
                                ComUtilities.Release(ref queryTable);
                            }
                        }
                    }
                    finally
                    {
                        ComUtilities.Release(ref queryTables);
                        ComUtilities.Release(ref worksheet);
                    }
                    if (hasTableConnection) break;
                }

                // Check for connections (for data model or other types)
                connections = ctx.Book.Connections;
                for (int i = 1; i <= connections.Count; i++)
                {
                    dynamic? conn = null;
                    try
                    {
                        conn = connections.Item(i);
                        string connName = conn.Name?.ToString() ?? "";

                        if (connName.Equals(queryName, StringComparison.OrdinalIgnoreCase) ||
                            connName.Equals($"Query - {queryName}", StringComparison.OrdinalIgnoreCase))
                        {
                            result.HasConnection = true;

                            // If we don't have a table connection but have a workbook connection,
                            // it's likely a data model connection
                            if (!hasTableConnection)
                            {
                                hasDataModelConnection = true;
                            }
                        }
                        else if (connName.Equals($"DataModel_{queryName}", StringComparison.OrdinalIgnoreCase))
                        {
                            // This is our explicit data model connection marker
                            result.HasConnection = true;
                            hasDataModelConnection = true;
                        }
                    }
                    finally
                    {
                        ComUtilities.Release(ref conn);
                    }
                }

                // Always check for named range markers that indicate data model loading
                // (even if we have table connections, for LoadToBoth mode)
                if (!hasDataModelConnection)
                {
                    // Check for our data model marker
                    try
                    {
                        names = ctx.Book.Names;
                        string markerName = $"DataModel_Query_{queryName}";

                        for (int i = 1; i <= names.Count; i++)
                        {
                            dynamic? existingName = null;
                            try
                            {
                                existingName = names.Item(i);
                                if (existingName.Name.ToString() == markerName)
                                {
                                    hasDataModelConnection = true;
                                    ComUtilities.Release(ref existingName);
                                    break;
                                }
                            }
                            finally
                            {
                                ComUtilities.Release(ref existingName);
                            }
                        }
                    }
                    catch
                    {
                        // Cannot check names
                    }

                    // Fallback: Check if the query has data model indicators
                    if (!hasDataModelConnection)
                    {
                        hasDataModelConnection = CheckQueryDataModelConfiguration(query, ctx.Book);
                    }
                }

                // Determine load mode
                if (hasTableConnection && hasDataModelConnection)
                {
                    result.LoadMode = PowerQueryLoadMode.LoadToBoth;
                }
                else if (hasTableConnection)
                {
                    result.LoadMode = PowerQueryLoadMode.LoadToTable;
                }
                else if (hasDataModelConnection)
                {
                    result.LoadMode = PowerQueryLoadMode.LoadToDataModel;
                }
                else
                {
                    result.LoadMode = PowerQueryLoadMode.ConnectionOnly;
                }

                result.TargetSheet = targetSheet;
                result.IsLoadedToDataModel = hasDataModelConnection;
                result.Success = true;
                return result;
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.ErrorMessage = $"Error getting load config: {ex.Message}";
                return result;
            }
            finally
            {
                ComUtilities.Release(ref names);
                ComUtilities.Release(ref connections);
                ComUtilities.Release(ref worksheets);
                ComUtilities.Release(ref query);
            }
        });
    }

    /// <inheritdoc />
    public async Task<OperationResult> DeleteAsync(IExcelBatch batch, string queryName)
    {
        var result = new OperationResult
        {
            FilePath = batch.WorkbookPath,
            Action = "pq-delete"
        };

        // Validate query name
        if (!ValidateQueryName(queryName, out string? validationError))
        {
            result.Success = false;
            result.ErrorMessage = validationError;
            return result;
        }

        return await batch.Execute<OperationResult>((ctx, ct) =>
        {
            dynamic? query = null;
            dynamic? queriesCollection = null;
            try
            {
                query = ComUtilities.FindQuery(ctx.Book, queryName);
                if (query == null)
                {
                    result.Success = false;
                    result.ErrorMessage = $"Query '{queryName}' not found";
                    return result;
                }

                queriesCollection = ctx.Book.Queries;
                queriesCollection.Item(queryName).Delete();

                result.Success = true;
                return result;
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.ErrorMessage = $"Error deleting query: {ex.Message}";
                return result;
            }
            finally
            {
                ComUtilities.Release(ref queriesCollection);
                ComUtilities.Release(ref query);
            }
        });
    }

    /// <summary>
    /// Helper to get all query names
    /// </summary>
    private static List<string> GetQueryNames(dynamic workbook)
    {
        var names = new List<string>();
        dynamic? queriesCollection = null;
        try
        {
            queriesCollection = workbook.Queries;
            for (int i = 1; i <= queriesCollection.Count; i++)
            {
                dynamic? query = null;
                try
                {
                    query = queriesCollection.Item(i);
                    names.Add(query.Name);
                }
                finally
                {
                    ComUtilities.Release(ref query);
                }
            }
        }
        catch { }
        finally
        {
            ComUtilities.Release(ref queriesCollection);
        }
        return names;
    }
}
