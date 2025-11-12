using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Models;

namespace Sbroenne.ExcelMcp.Core.Commands;

/// <summary>
/// Power Query operations - FilePath-based API using direct FileHandleManager integration
/// </summary>
public partial class PowerQueryCommands
{
    /// <inheritdoc />
    public async Task<PowerQueryListResult> ListAsync(string filePath)
    {
        var result = new PowerQueryListResult { FilePath = filePath };

        try
        {
            var handle = await FileHandleManager.Instance.OpenOrGetAsync(filePath);

            await Task.Run(() =>
            {
                dynamic? queriesCollection = null;
                try
                {
                    queriesCollection = handle.Workbook.Queries;
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
                                connections = handle.Workbook.Connections;
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
                }
                finally
                {
                    ComUtilities.Release(ref queriesCollection);
                }
            });

            return result;
        }
        catch (Exception ex)
        {
            result.Success = false;
            result.ErrorMessage = $"Failed to access workbook: {ex.Message}";
            return result;
        }
    }

    /// <inheritdoc />
    public async Task<PowerQueryViewResult> ViewAsync(string filePath, string queryName)
    {
        // For now, delegate to batch method to avoid massive duplication
        // TODO: Implement direct FileHandleManager version
        await using var batch = await ExcelSession.BeginBatchAsync(filePath);
        return await ViewAsync(batch, queryName);
    }

    /// <inheritdoc />
    public async Task<OperationResult> ExportAsync(string filePath, string queryName, string outputFile)
    {
        await using var batch = await ExcelSession.BeginBatchAsync(filePath);
        return await ExportAsync(batch, queryName, outputFile);
    }

    /// <inheritdoc />
    public async Task<PowerQueryRefreshResult> RefreshAsync(string filePath, string queryName)
    {
        await using var batch = await ExcelSession.BeginBatchAsync(filePath);
        return await RefreshAsync(batch, queryName);
    }

    /// <inheritdoc />
    public async Task<PowerQueryRefreshResult> RefreshAsync(string filePath, string queryName, TimeSpan? timeout)
    {
        await using var batch = await ExcelSession.BeginBatchAsync(filePath);
        return await RefreshAsync(batch, queryName, timeout);
    }

    /// <inheritdoc />
    public async Task<PowerQueryLoadConfigResult> GetLoadConfigAsync(string filePath, string queryName)
    {
        await using var batch = await ExcelSession.BeginBatchAsync(filePath);
        return await GetLoadConfigAsync(batch, queryName);
    }

    /// <inheritdoc />
    public async Task<OperationResult> DeleteAsync(string filePath, string queryName)
    {
        await using var batch = await ExcelSession.BeginBatchAsync(filePath);
        return await DeleteAsync(batch, queryName);
    }

    /// <inheritdoc />
    public async Task<WorksheetListResult> ListExcelSourcesAsync(string filePath)
    {
        await using var batch = await ExcelSession.BeginBatchAsync(filePath);
        return await ListExcelSourcesAsync(batch);
    }

    /// <inheritdoc />
    public async Task<PowerQueryCreateResult> CreateAsync(string filePath, string queryName, string mCodeFile, PowerQueryLoadMode loadMode = PowerQueryLoadMode.LoadToTable, string? targetSheet = null)
    {
        await using var batch = await ExcelSession.BeginBatchAsync(filePath);
        var result = await CreateAsync(batch, queryName, mCodeFile, loadMode, targetSheet);
        await batch.SaveAsync();
        return result;
    }

    /// <inheritdoc />
    public async Task<OperationResult> UpdateAsync(string filePath, string queryName, string mCodeFile)
    {
        await using var batch = await ExcelSession.BeginBatchAsync(filePath);
        var result = await UpdateAsync(batch, queryName, mCodeFile);
        await batch.SaveAsync();
        return result;
    }

    /// <inheritdoc />
    public async Task<PowerQueryLoadResult> LoadToAsync(string filePath, string queryName, PowerQueryLoadMode loadMode, string? targetSheet = null)
    {
        await using var batch = await ExcelSession.BeginBatchAsync(filePath);
        var result = await LoadToAsync(batch, queryName, loadMode, targetSheet);
        await batch.SaveAsync();
        return result;
    }

    /// <inheritdoc />
    public async Task<OperationResult> UnloadAsync(string filePath, string queryName)
    {
        await using var batch = await ExcelSession.BeginBatchAsync(filePath);
        var result = await UnloadAsync(batch, queryName);
        await batch.SaveAsync();
        return result;
    }

    /// <inheritdoc />
    public async Task<OperationResult> RefreshAllAsync(string filePath)
    {
        await using var batch = await ExcelSession.BeginBatchAsync(filePath);
        return await RefreshAllAsync(batch);
    }
}
