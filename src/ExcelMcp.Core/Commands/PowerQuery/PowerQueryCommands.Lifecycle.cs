using System.Runtime.InteropServices;
using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Models;
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
    public async Task<OperationResult> UpdateAsync(IExcelBatch batch, string queryName, string mCodeFile)
    {
        var result = new OperationResult
        {
            FilePath = batch.WorkbookPath,
            Action = "pq-update"
        };

        // Validate query name length (Excel limit: 120 characters)
        if (string.IsNullOrWhiteSpace(queryName))
        {
            result.Success = false;
            result.ErrorMessage = "Query name cannot be empty or whitespace";
            return result;
        }

        if (queryName.Length > 120)
        {
            result.Success = false;
            result.ErrorMessage = $"Query name exceeds Excel's 120-character limit (current length: {queryName.Length})";
            return result;
        }

        // Validate and normalize the M code file path to prevent path traversal attacks
        try
        {
            mCodeFile = PathValidator.ValidateExistingFile(mCodeFile, nameof(mCodeFile));
        }
        catch (Exception ex)
        {
            result.Success = false;
            result.ErrorMessage = $"Invalid M code file path: {ex.Message}";
            return result;
        }

        string mCode = await File.ReadAllTextAsync(mCodeFile);

        // Update the query M code
        // NOTE: Excel preserves load configuration when updating query.Formula property
        result = await batch.Execute<OperationResult>((ctx, ct) =>
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

                // Update M code
                query.Formula = mCode;
                result.Success = true;

                return result;
            }
            catch (COMException comEx) when (comEx.Message.Contains("Information is needed in order to combine data") ||
                                             comEx.Message.Contains("Formula.Firewall", StringComparison.OrdinalIgnoreCase))
            {
                // Privacy level error - must be configured manually in Excel UI
                result.Success = false;
                result.ErrorMessage = "Privacy level error: This query combines data from multiple sources. " +
                                    "Open the file in Excel and configure privacy levels manually: " +
                                    "File → Options → Privacy. See COMMANDS.md for details.";
                return result;
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.ErrorMessage = $"Error updating query: {ex.Message}";
                return result;
            }
            finally
            {
                ComUtilities.Release(ref query);
            }
        });

        return result;
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
    public async Task<OperationResult> ImportAsync(IExcelBatch batch, string queryName, string mCodeFile, string loadDestination = "worksheet", string? worksheetName = null)
    {
        var result = new OperationResult
        {
            FilePath = batch.WorkbookPath,
            Action = "pq-import"
        };

        // Validate query name length (Excel limit: 120 characters)
        if (string.IsNullOrWhiteSpace(queryName))
        {
            result.Success = false;
            result.ErrorMessage = "Query name cannot be empty or whitespace";
            return result;
        }

        if (queryName.Length > 120)
        {
            result.Success = false;
            result.ErrorMessage = $"Query name exceeds Excel's 120-character limit (current length: {queryName.Length})";
            return result;
        }

        // Validate loadDestination parameter
        var validDestinations = new[] { "worksheet", "data-model", "both", "connection-only" };
        if (!validDestinations.Contains(loadDestination.ToLowerInvariant()))
        {
            result.Success = false;
            result.ErrorMessage = $"Invalid loadDestination: '{loadDestination}'. Valid values: {string.Join(", ", validDestinations)}";
            return result;
        }

        // Validate and normalize the M code file path to prevent path traversal attacks
        try
        {
            mCodeFile = PathValidator.ValidateExistingFile(mCodeFile, nameof(mCodeFile));
        }
        catch (Exception ex)
        {
            result.Success = false;
            result.ErrorMessage = $"Invalid M code file path: {ex.Message}";
            return result;
        }

        string mCode = await File.ReadAllTextAsync(mCodeFile);

        result = await batch.Execute<OperationResult>((ctx, ct) =>
        {
            dynamic? existingQuery = null;
            dynamic? queriesCollection = null;
            dynamic? newQuery = null;
            try
            {
                // Check if query already exists
                existingQuery = ComUtilities.FindQuery(ctx.Book, queryName);
                if (existingQuery != null)
                {
                    result.Success = false;
                    result.ErrorMessage = $"Query '{queryName}' already exists. Use pq-update to modify it.";
                    return result;
                }

                // Add new query
                queriesCollection = ctx.Book.Queries;
                newQuery = queriesCollection.Add(queryName, mCode);

                result.Success = true;
                return result;
            }
            catch (COMException comEx) when (comEx.Message.Contains("Information is needed in order to combine data") ||
                                             comEx.Message.Contains("Formula.Firewall", StringComparison.OrdinalIgnoreCase))
            {
                // Privacy level error - must be configured manually in Excel UI
                result.Success = false;
                result.ErrorMessage = "Privacy level error: This query combines data from multiple sources. " +
                                    "Open the file in Excel and configure privacy levels manually: " +
                                    "File → Options → Privacy. See COMMANDS.md for details.";
                return result;
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.ErrorMessage = $"Error importing query: {ex.Message}";
                return result;
            }
            finally
            {
                ComUtilities.Release(ref newQuery);
                ComUtilities.Release(ref queriesCollection);
                ComUtilities.Release(ref existingQuery);
            }
        });

        // Auto-load based on loadDestination parameter
        if (result.Success)
        {
            var destination = loadDestination.ToLowerInvariant();
            OperationResult? loadResult = null;

            switch (destination)
            {
                case "worksheet":
                    string targetSheet = worksheetName ?? queryName;
                    loadResult = await SetLoadToTableAsync(batch, queryName, targetSheet);
                    break;

                case "data-model":
                    var dmResult = await SetLoadToDataModelAsync(batch, queryName);
                    loadResult = new OperationResult
                    {
                        Success = dmResult.Success,
                        ErrorMessage = dmResult.ErrorMessage,
                        FilePath = dmResult.FilePath
                    };
                    break;

                case "both":
                    string targetSheetBoth = worksheetName ?? queryName;
                    var bothResult = await SetLoadToBothAsync(batch, queryName, targetSheetBoth);
                    loadResult = new OperationResult
                    {
                        Success = bothResult.Success,
                        ErrorMessage = bothResult.ErrorMessage,
                        FilePath = bothResult.FilePath
                    };
                    break;

                case "connection-only":
                    // No loading - query imported but not executed
                    return result;
            }

            // Handle loading result
            if (loadResult != null && !loadResult.Success)
            {
                // Loading failed - this is a FAILURE, not success
                result.Success = false;
                result.ErrorMessage = $"Query imported but failed to load to {destination}: {loadResult.ErrorMessage}";
                return result;
            }
            else if (loadResult != null && loadResult.Success)
            {
                // CRITICAL: Save the workbook to persist changes
                await batch.SaveAsync();

                // Query was loaded successfully
                return result;
            }
        }

        return result;
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
