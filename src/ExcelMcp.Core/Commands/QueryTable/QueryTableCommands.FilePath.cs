using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Models;

namespace Sbroenne.ExcelMcp.Core.Commands.QueryTable;

/// <summary>
/// QueryTable operations - FilePath-based API implementations
/// These methods internally use temporary batch sessions to leverage existing QueryTable logic
/// </summary>
public partial class QueryTableCommands
{
    /// <inheritdoc />
    public async Task<QueryTableListResult> ListAsync(string filePath)
    {
        await using var batch = await ExcelSession.BeginBatchAsync(filePath);
        return await ListAsync(batch);
    }

    /// <inheritdoc />
    public async Task<QueryTableInfoResult> GetAsync(string filePath, string queryTableName)
    {
        await using var batch = await ExcelSession.BeginBatchAsync(filePath);
        return await GetAsync(batch, queryTableName);
    }

    /// <inheritdoc />
    public async Task<OperationResult> CreateFromConnectionAsync(string filePath, string sheetName,
        string queryTableName, string connectionName, string range = "A1",
        QueryTableCreateOptions? options = null)
    {
        await using var batch = await ExcelSession.BeginBatchAsync(filePath);
        var result = await CreateFromConnectionAsync(batch, sheetName, queryTableName, connectionName, range, options);
        await batch.SaveAsync();
        return result;
    }

    /// <inheritdoc />
    public async Task<OperationResult> CreateFromQueryAsync(string filePath, string sheetName,
        string queryTableName, string queryName, string range = "A1",
        QueryTableCreateOptions? options = null)
    {
        await using var batch = await ExcelSession.BeginBatchAsync(filePath);
        var result = await CreateFromQueryAsync(batch, sheetName, queryTableName, queryName, range, options);
        await batch.SaveAsync();
        return result;
    }

    /// <inheritdoc />
    public async Task<OperationResult> RefreshAsync(string filePath, string queryTableName, TimeSpan? timeout = null)
    {
        await using var batch = await ExcelSession.BeginBatchAsync(filePath);
        return await RefreshAsync(batch, queryTableName, timeout);
    }

    /// <inheritdoc />
    public async Task<OperationResult> UpdatePropertiesAsync(string filePath, string queryTableName,
        QueryTableUpdateOptions options)
    {
        await using var batch = await ExcelSession.BeginBatchAsync(filePath);
        var result = await UpdatePropertiesAsync(batch, queryTableName, options);
        await batch.SaveAsync();
        return result;
    }

    /// <inheritdoc />
    public async Task<OperationResult> DeleteAsync(string filePath, string queryTableName)
    {
        await using var batch = await ExcelSession.BeginBatchAsync(filePath);
        var result = await DeleteAsync(batch, queryTableName);
        await batch.SaveAsync();
        return result;
    }

    /// <inheritdoc />
    public async Task<OperationResult> RefreshAllAsync(string filePath, TimeSpan? timeout = null)
    {
        await using var batch = await ExcelSession.BeginBatchAsync(filePath);
        return await RefreshAllAsync(batch, timeout);
    }
}
