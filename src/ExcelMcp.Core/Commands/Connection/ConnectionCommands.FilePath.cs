using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Models;

namespace Sbroenne.ExcelMcp.Core.Commands;

/// <summary>
/// Connection operations - FilePath-based API implementations
/// These methods internally use temporary batch sessions to leverage existing connection logic
/// </summary>
public partial class ConnectionCommands
{
    /// <inheritdoc />
    public async Task<ConnectionListResult> ListAsync(string filePath)
    {
        await using var batch = await ExcelSession.BeginBatchAsync(filePath);
        return await ListAsync(batch);
    }

    /// <inheritdoc />
    public async Task<ConnectionViewResult> ViewAsync(string filePath, string connectionName)
    {
        await using var batch = await ExcelSession.BeginBatchAsync(filePath);
        return await ViewAsync(batch, connectionName);
    }

    /// <inheritdoc />
    public async Task<OperationResult> CreateAsync(string filePath, string connectionName,
        string connectionString, string? commandText = null, string? description = null)
    {
        await using var batch = await ExcelSession.BeginBatchAsync(filePath);
        var result = await CreateAsync(batch, connectionName, connectionString, commandText, description);
        await batch.SaveAsync();
        return result;
    }

    /// <inheritdoc />
    public async Task<OperationResult> ImportAsync(string filePath, string connectionName, string jsonFilePath)
    {
        await using var batch = await ExcelSession.BeginBatchAsync(filePath);
        var result = await ImportAsync(batch, connectionName, jsonFilePath);
        await batch.SaveAsync();
        return result;
    }

    /// <inheritdoc />
    public async Task<OperationResult> ExportAsync(string filePath, string connectionName, string jsonFilePath)
    {
        await using var batch = await ExcelSession.BeginBatchAsync(filePath);
        return await ExportAsync(batch, connectionName, jsonFilePath);
    }

    /// <inheritdoc />
    public async Task<OperationResult> UpdatePropertiesAsync(string filePath, string connectionName, string jsonFilePath)
    {
        await using var batch = await ExcelSession.BeginBatchAsync(filePath);
        var result = await UpdatePropertiesAsync(batch, connectionName, jsonFilePath);
        await batch.SaveAsync();
        return result;
    }

    /// <inheritdoc />
    public async Task<OperationResult> RefreshAsync(string filePath, string connectionName)
    {
        await using var batch = await ExcelSession.BeginBatchAsync(filePath);
        return await RefreshAsync(batch, connectionName);
    }

    /// <inheritdoc />
    public async Task<OperationResult> RefreshAsync(string filePath, string connectionName, TimeSpan? timeout)
    {
        await using var batch = await ExcelSession.BeginBatchAsync(filePath);
        return await RefreshAsync(batch, connectionName, timeout);
    }

    /// <inheritdoc />
    public async Task<OperationResult> DeleteAsync(string filePath, string connectionName)
    {
        await using var batch = await ExcelSession.BeginBatchAsync(filePath);
        var result = await DeleteAsync(batch, connectionName);
        await batch.SaveAsync();
        return result;
    }

    /// <inheritdoc />
    public async Task<OperationResult> LoadToAsync(string filePath, string connectionName, string sheetName)
    {
        await using var batch = await ExcelSession.BeginBatchAsync(filePath);
        var result = await LoadToAsync(batch, connectionName, sheetName);
        await batch.SaveAsync();
        return result;
    }

    /// <inheritdoc />
    public async Task<ConnectionPropertiesResult> GetPropertiesAsync(string filePath, string connectionName)
    {
        await using var batch = await ExcelSession.BeginBatchAsync(filePath);
        return await GetPropertiesAsync(batch, connectionName);
    }

    /// <inheritdoc />
    public async Task<OperationResult> SetPropertiesAsync(string filePath, string connectionName,
        bool? backgroundQuery = null, bool? refreshOnFileOpen = null,
        bool? savePassword = null, int? refreshPeriod = null)
    {
        await using var batch = await ExcelSession.BeginBatchAsync(filePath);
        var result = await SetPropertiesAsync(batch, connectionName, backgroundQuery, refreshOnFileOpen, savePassword, refreshPeriod);
        await batch.SaveAsync();
        return result;
    }

    /// <inheritdoc />
    public async Task<OperationResult> TestAsync(string filePath, string connectionName)
    {
        await using var batch = await ExcelSession.BeginBatchAsync(filePath);
        return await TestAsync(batch, connectionName);
    }
}
