namespace Sbroenne.ExcelMcp.Service.Mac;

internal sealed class MacExcelSession
{
    public required string SessionId { get; init; }
    public required string FilePath { get; init; }
    public required TimeSpan OperationTimeout { get; init; }
    public required bool IsVisible { get; set; }
    public DateTime CreatedAt { get; } = DateTime.UtcNow;
    public SemaphoreSlim OperationLock { get; } = new(1, 1);
    public int ActiveOperations;
}
