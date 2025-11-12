using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Models;

namespace Sbroenne.ExcelMcp.Core.Commands;

/// <summary>
/// Power Query operations - FilePath-based API implementations
/// These methods internally use temporary batch sessions to leverage existing Power Query logic
/// </summary>
public partial class PowerQueryCommands
{
    /// <inheritdoc />
    public async Task<PowerQueryListResult> ListAsync(string filePath)
    {
        await using var batch = await ExcelSession.BeginBatchAsync(filePath);
        return await ListAsync(batch);
    }

    /// <inheritdoc />
    public async Task<PowerQueryViewResult> ViewAsync(string filePath, string queryName)
    {
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
