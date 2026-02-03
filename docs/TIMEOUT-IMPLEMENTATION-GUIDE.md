# Timeout Implementation Guide

> **Complete guide for using and enhancing the timeout protection system**

## Overview

Excel batch operations now have timeout protection to prevent indefinite hangs when Excel becomes unresponsive (modal dialogs, data source issues, COM deadlocks).

**Key Features:**
- Default 2-minute timeout for all operations
- Maximum 5-minute timeout (prevents infinite waits)
- Operation-specific timeout overrides
- Rich error messages with LLM guidance
- Automatic progress logging to stderr (MCP protocol)

---

## Core Implementation

### Constants (ExcelBatch.cs)

```csharp
/// <summary>
/// Default timeout for Excel operations (2 minutes)
/// </summary>
private static readonly TimeSpan DefaultOperationTimeout = TimeSpan.FromMinutes(2);

/// <summary>
/// Maximum allowed timeout (5 minutes) - prevents infinite waits
/// </summary>
private static readonly TimeSpan MaxOperationTimeout = TimeSpan.FromMinutes(5);
```

### IExcelBatch Interface

```csharp
Task<T> Execute<T>(
    Func<ExcelContext, CancellationToken, T> operation,
    CancellationToken cancellationToken = default,
    TimeSpan? timeout = null);  // ← NEW: Optional timeout parameter

Task<T> ExecuteAsync<T>(
    Func<ExcelContext, CancellationToken, Task<T>> operation,
    CancellationToken cancellationToken = default,
    TimeSpan? timeout = null);  // ← NEW: Optional timeout parameter
```

### Timeout Enforcement Pattern

```csharp
// Clamp timeout between default and max
var effectiveTimeout = timeout.HasValue
    ? TimeSpan.FromMilliseconds(Math.Min(timeout.Value.TotalMilliseconds, MaxOperationTimeout.TotalMilliseconds))
    : DefaultOperationTimeout;

// Create linked cancellation token (operation + timeout)
using var timeoutCts = new CancellationTokenSource(effectiveTimeout);
using var linkedCts = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken, timeoutCts.Token);

// Use Task.WaitAsync for timeout enforcement
return await tcs.Task.WaitAsync(linkedCts.Token);
```

### Error Handling Pattern

```csharp
catch (OperationCanceledException) when (timeoutCts.IsCancellationRequested && !cancellationToken.IsCancellationRequested)
{
    var duration = DateTime.UtcNow - startTime;
    var usedMaxTimeout = effectiveTimeout >= MaxOperationTimeout;

    Console.Error.WriteLine($"[EXCEL-BATCH] TIMEOUT after {duration.TotalSeconds:F1}s (limit: {effectiveTimeout.TotalMinutes:F1}min, max: {usedMaxTimeout})");

    var message = usedMaxTimeout
        ? $"Excel operation exceeded maximum timeout of {MaxOperationTimeout.TotalMinutes} minutes (actual: {duration.TotalMinutes:F1} min). " +
          "This indicates Excel is hung, unresponsive, or the operation is too complex. " +
          "Check if Excel is showing a dialog or consider breaking the operation into smaller steps."
        : $"Excel operation timed out after {effectiveTimeout.TotalMinutes} minutes (actual: {duration.TotalMinutes:F1} min). " +
          $"For large datasets or complex operations, more time may be needed (maximum: {MaxOperationTimeout.TotalMinutes} min).";

    throw new TimeoutException(message);
}
```

---

## Enhanced Result Types

### ResultBase Additions

All result types now inherit LLM guidance fields:

```csharp
public abstract class ResultBase
{
    public bool Success { get; set; }
    public string? ErrorMessage { get; set; }
    public string? FilePath { get; set; }

    // ✨ NEW: LLM Guidance Fields
    public List<string>? SuggestedNextActions { get; set; }
    public Dictionary<string, object>? OperationContext { get; set; }
    public bool IsRetryable { get; set; } = true;
    public string? RetryGuidance { get; set; }
}
```

---

## Usage Patterns

### Pattern 1: Core Commands (Heavy Operations)

For operations that typically take longer (refresh, data loading), Core commands can pass custom timeout:

```csharp
// In PowerQueryCommands.cs
public async Task<PowerQueryRefreshResult> RefreshAsync(IExcelBatch batch, string queryName)
{
    // Heavy operation: request extended timeout (5 minutes)
    return await batch.Execute<PowerQueryRefreshResult>(
        (ctx, ct) => {
            // Refresh logic...
            return result;
        },
        timeout: TimeSpan.FromMinutes(5)  // ← Request 5 min (will be clamped to max)
    );
}
```

### Pattern 2: MCP Tool Timeout Exception Handling

MCP tools should catch `TimeoutException` and enrich with operation-specific guidance:

```csharp
// In ExcelPowerQueryTool.cs
private static async Task<string> RefreshPowerQueryAsync(...)
{
    try
    {
        var result = await ExcelToolsBase.WithBatchAsync(
            batchId,
            excelPath,
            save: true,
            async (batch) => await commands.RefreshAsync(batch, queryName));

        return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions);
    }
    catch (TimeoutException ex)
    {
        // Enrich with operation-specific guidance
        var result = new PowerQueryRefreshResult
        {
            Success = false,
            ErrorMessage = ex.Message,
            QueryName = queryName,
            FilePath = excelPath,
            
            // ✨ Add LLM guidance
            SuggestedNextActions = new List<string>
            {
                "Check if Excel is showing a 'Privacy Level' dialog or other modal dialogs",
                "Verify the data source is accessible (network connection, database availability)",
                "Consider breaking query into smaller steps if processing large datasets",
                "Use batch mode (begin_excel_batch) if not already using it"
            },
            
            OperationContext = new Dictionary<string, object>
            {
                { "OperationType", "PowerQuery.Refresh" },
                { "QueryName", queryName },
                { "TimeoutReached", true },
                { "UsedMaxTimeout", ex.Message.Contains("maximum timeout") }
            },
            
            IsRetryable = !ex.Message.Contains("maximum timeout"),  // Don't retry max timeout
            
            RetryGuidance = ex.Message.Contains("maximum timeout")
                ? "Operation reached maximum timeout - do not retry. Check for Excel dialogs or break into smaller operations."
                : "Operation can be retried with longer timeout (up to 5 minutes) if data source is slow."
        };

        return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions);
    }
}
```

### Pattern 3: Light Operations (Use Default)

For quick operations (list, get, create), use default 2-minute timeout:

```csharp
// In PowerQueryCommands.cs
public async Task<PowerQueryListResult> ListAsync(IExcelBatch batch)
{
    // Light operation: use default timeout (no parameter needed)
    return await batch.Execute<PowerQueryListResult>((ctx, ct) =>
    {
        // List logic...
        return result;
    });  // ← Default 2 min timeout
}
```

---

## Operation-Specific Timeout Recommendations

| Operation Type | Recommended Timeout | Rationale |
|----------------|-------------------|-----------|
| **List operations** | Default (2 min) | Reading metadata is fast |
| **Get/View operations** | Default (2 min) | Reading single item is fast |
| **Create operations** | Default (2 min) | Creating objects is typically fast |
| **Delete operations** | Default (2 min) | Deleting is fast |
| **Refresh operations** | 5 minutes (max) | Data source queries can be slow |
| **Import operations** | 5 minutes (max) | Loading large datasets |
| **Data Model refresh** | 5 minutes (max) | Processing millions of rows |
| **VBA execution** | Context-dependent | User code complexity varies |

---

## Stderr Logging (MCP Protocol)

Operations automatically log progress to stderr (visible in MCP clients):

```
[EXCEL-BATCH] Starting operation (timeout: 2.0min)
[EXCEL-BATCH] Completed in 3.5s
```

Or on timeout:

```
[EXCEL-BATCH] Starting operation (timeout: 5.0min)
[EXCEL-BATCH] TIMEOUT after 300.2s (limit: 5.0min, max: true)
```

---

## Integration Checklist

### For New Core Commands

- [ ] Determine if operation is heavy or light
- [ ] Pass `timeout: TimeSpan.FromMinutes(5)` for heavy operations
- [ ] Use default timeout for light operations
- [ ] Add timeout context to result on failure

### For MCP Tools

- [ ] Wrap Core command calls in try-catch for `TimeoutException`
- [ ] Create enriched result with:
  - [ ] `SuggestedNextActions` (operation-specific steps)
  - [ ] `OperationContext` (timeout details, operation type)
  - [ ] `IsRetryable = false` for max timeout cases
  - [ ] `RetryGuidance` with actionable advice
- [ ] Return JSON serialized result (don't re-throw)

### For CLI Commands

- [ ] Display timeout errors with guidance
- [ ] Show suggested next actions in readable format
- [ ] Indicate if operation is retryable

---

## Example: Complete Timeout-Aware Implementation

```csharp
// Core Command (PowerQueryCommands.Refresh.cs)
public async Task<PowerQueryRefreshResult> RefreshAsync(IExcelBatch batch, string queryName)
{
    var result = new PowerQueryRefreshResult
    {
        FilePath = batch.WorkbookPath,
        QueryName = queryName,
        RefreshTime = DateTime.Now
    };

    return await batch.Execute<PowerQueryRefreshResult>(
        (ctx, ct) =>
        {
            // Refresh logic (omitted for brevity)
            result.Success = true;
            return result;
        },
        timeout: TimeSpan.FromMinutes(5)  // ← Request extended timeout
    );
}

// MCP Tool (ExcelPowerQueryTool.cs)
private static async Task<string> RefreshPowerQueryAsync(
    PowerQueryCommands commands, 
    string excelPath, 
    string? queryName, 
    string? batchId)
{
    if (string.IsNullOrEmpty(queryName))
        throw new ModelContextProtocol.McpException("queryName is required for refresh action");

    try
    {
        var result = await ExcelToolsBase.WithBatchAsync(
            batchId,
            excelPath,
            save: true,
            async (batch) => await commands.RefreshAsync(batch, queryName));

        return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions);
    }
    catch (TimeoutException ex)
    {
        var result = new PowerQueryRefreshResult
        {
            Success = false,
            ErrorMessage = ex.Message,
            QueryName = queryName,
            FilePath = excelPath,
            RefreshTime = DateTime.Now,
            
            SuggestedNextActions = new List<string>
            {
                "Check if Excel is showing a 'Privacy Level' dialog",
                "Verify the data source is accessible",
                "Consider using smaller data ranges if processing large datasets",
                "Use batch mode (begin_excel_batch) if not already"
            },
            
            OperationContext = new Dictionary<string, object>
            {
                { "OperationType", "PowerQuery.Refresh" },
                { "QueryName", queryName },
                { "TimeoutReached", true },
                { "UsedMaxTimeout", ex.Message.Contains("maximum timeout") }
            },
            
            IsRetryable = !ex.Message.Contains("maximum timeout"),
            
            RetryGuidance = ex.Message.Contains("maximum timeout")
                ? "Maximum timeout reached - do not retry automatically. Manual intervention needed."
                : "Retry with same timeout acceptable if transient issue suspected."
        };

        return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions);
    }
}
```

---

## Testing Strategy

### Unit Tests (Fast)

Test timeout clamping logic:

```csharp
[Fact]
public void Execute_TimeoutClamping_LimitsToMax()
{
    var requestedTimeout = TimeSpan.FromMinutes(10);  // Request 10 min
    var effectiveTimeout = Math.Min(requestedTimeout.TotalMilliseconds, MaxOperationTimeout.TotalMilliseconds);
    Assert.Equal(TimeSpan.FromMinutes(5), TimeSpan.FromMilliseconds(effectiveTimeout));  // Clamped to 5 min
}
```

### Integration Tests (Requires Excel)

Test actual timeout behavior:

```csharp
[Fact]
public async Task Execute_ExceedsTimeout_ThrowsTimeoutException()
{
    await using var batch = await ExcelSession.BeginBatchAsync(testFile);
    
    // Operation that takes 3 seconds with 1-second timeout
    await Assert.ThrowsAsync<TimeoutException>(async () =>
    {
        await batch.Execute<int>((ctx, ct) =>
        {
            Thread.Sleep(3000);  // Simulate slow operation
            return 0;
        }, timeout: TimeSpan.FromSeconds(1));
    });
}
```

---

## Troubleshooting

### Symptom: All operations timing out

**Likely Cause:** Excel showing modal dialog (Privacy Level, Credentials, etc.)

**Solution:**
1. Check if Excel process is visible (should be hidden background process)
2. Kill all Excel processes: `taskkill /F /IM excel.exe`
3. Re-run operation
4. If persistent, manually open Excel, configure privacy levels, save workbook

### Symptom: Max timeout reached frequently

**Likely Cause:** Operation genuinely too complex or data source very slow

**Solution:**
1. Break operation into smaller chunks (filter data, process in batches)
2. Optimize data source (add indexes, reduce query complexity)
3. Consider Power Query query folding to push processing to source

### Symptom: Timeout too aggressive (legitimate operations failing)

**Likely Cause:** Default timeout too short for specific scenario

**Solution:**
1. Core Commands: Pass explicit timeout: `timeout: TimeSpan.FromMinutes(5)`
2. Cannot exceed 5-minute max (architectural decision to prevent infinite hangs)

---

## Future Enhancements

### Potential Improvements

1. **Configurable Max Timeout**: Allow per-workbook or per-session max timeout override (currently hardcoded 5 min)
2. **Progress Callbacks**: Provide progress updates during long operations
3. **Timeout Metrics**: Collect stats on timeout frequency to identify problematic operations
4. **Adaptive Timeout**: Automatically increase timeout for operations that consistently hit limit

---

## Related Documentation

- [Excel COM Interop Patterns](.github/instructions/excel-com-interop.instructions.md)
- [MCP Server Guide](.github/instructions/mcp-server-guide.instructions.md)
- [Testing Strategy](.github/instructions/testing-strategy.instructions.md)

---

**Last Updated:** 2025-01-XX  
**Author:** AI Agent + User Collaboration
