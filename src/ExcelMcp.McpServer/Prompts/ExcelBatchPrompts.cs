using System.ComponentModel;
using ModelContextProtocol.Server;
using Microsoft.Extensions.AI;

namespace Sbroenne.ExcelMcp.McpServer.Prompts;

/// <summary>
/// MCP Prompts that teach LLMs about Excel batch session management.
/// These prompts are discoverable via the MCP protocol and help LLMs understand when and how to use batch sessions.
/// </summary>
[McpServerPromptType]
public static class ExcelBatchPrompts
{
    /// <summary>
    /// Comprehensive guide on Excel batch session management for multi-operation workflows.
    /// </summary>
    [McpServerPrompt(Name = "excel_batch_guide"), Description("Learn when and how to use Excel batch sessions for high-performance multi-operation workflows")]
    public static ChatMessage BatchSessionGuide()
    {
        return new ChatMessage(ChatRole.User, @"# Excel Batch Session Management Guide

## Overview
Excel batch sessions allow you to perform multiple Excel operations efficiently by keeping the workbook open across operations. This avoids 2-5 second Excel startup overhead per operation.

## When to Use Batch Sessions

✅ **USE batch sessions when:**
- Performing **multiple related operations** on the same workbook (2+ operations)
- Building complex workflows (e.g., import query → update → refresh → read results)
- Performance matters (batch sessions are ~95% faster for multi-step operations)
- You want **explicit control** over when to save changes

❌ **DON'T use batch sessions when:**
- Performing a **single operation** (use tools without batchId - they auto-handle lifecycle)
- Working with **multiple different workbooks** simultaneously (create separate batches)

## How to Use Batch Sessions

### Step 1: Begin a Batch
```typescript
const result = await begin_excel_batch({
  filePath: ""C:\\data\\sales.xlsx""
});
const batchId = result.batchId; // Save this ID for subsequent operations
```

### Step 2: Perform Operations (Pass batchId)
```typescript
// All operations use the SAME open workbook (fast!)
await excel_powerquery({
  batchId: batchId,  // ← Pass the batch ID
  action: ""import"",
  excelPath: ""C:\\data\\sales.xlsx"",
  queryName: ""SalesData"",
  mCodeFile: ""queries/sales.pq""
});

await excel_powerquery({
  batchId: batchId,  // ← Same batch ID
  action: ""refresh"",
  excelPath: ""C:\\data\\sales.xlsx"",
  queryName: ""SalesData""
});

// Read the results
const data = await excel_worksheet({
  batchId: batchId,
  action: ""read"",
  excelPath: ""C:\\data\\sales.xlsx"",
  sheetName: ""RawData""
});
```

### Step 3: Commit or Discard
```typescript
// Save and close
await commit_excel_batch({
  batchId: batchId,
  save: true
});

// Or discard changes
await commit_excel_batch({
  batchId: batchId,
  save: false
});
```

## Important Rules

### Rule 1: File Path Must Match
All operations in a batch MUST use the **same Excel file**:
```typescript
// ✅ CORRECT
begin_excel_batch({ filePath: ""sales.xlsx"" })
excel_powerquery({ batchId, excelPath: ""sales.xlsx"", ... })

// ❌ WRONG - different file
begin_excel_batch({ filePath: ""sales.xlsx"" })
excel_powerquery({ batchId, excelPath: ""other.xlsx"", ... })  // ERROR!
```

### Rule 2: Always Commit When Done
Never abandon a batch - it keeps Excel running:
```typescript
try {
  const { batchId } = await begin_excel_batch({ filePath: ""sales.xlsx"" });
  await excel_powerquery({ batchId, ... });
  await commit_excel_batch({ batchId, save: true });
} catch (error) {
  if (batchId) await commit_excel_batch({ batchId, save: false });
}
```

### Rule 3: One Batch Per Workbook
Each workbook can only have **one active batch** at a time.

## Backward Compatibility
All tools work **without batchId** for single operations (automatic batch-of-one).

## Performance Comparison

**Without batch (4 operations):** 4 × 2-5 sec = 8-20 seconds
**With batch:** ~3 seconds total (2× to 10× faster)

## Example Workflows

### Import and Validate Query
```typescript
const { batchId } = await begin_excel_batch({ filePath: ""report.xlsx"" });
try {
  await excel_powerquery({ batchId, action: ""import"", queryName: ""Sales"", mCodeFile: ""sales.pq"" });
  await excel_powerquery({ batchId, action: ""set-load-to-table"", queryName: ""Sales"", sheetName: ""Data"" });
  const result = await excel_powerquery({ batchId, action: ""refresh"", queryName: ""Sales"" });
  
  if (result.success) {
    await commit_excel_batch({ batchId, save: true });
  } else {
    await commit_excel_batch({ batchId, save: false });
  }
} catch (error) {
  await commit_excel_batch({ batchId, save: false });
}
```

### Bulk Processing
```typescript
const { batchId } = await begin_excel_batch({ filePath: ""analysis.xlsx"" });
try {
  await excel_worksheet({ batchId, action: ""create"", sheetName: ""RawData"" });
  await excel_worksheet({ batchId, action: ""create"", sheetName: ""Summary"" });
  await excel_worksheet({ batchId, action: ""write"", sheetName: ""RawData"", csvFile: ""data.csv"" });
  await excel_parameter({ batchId, action: ""set"", paramName: ""ReportDate"", value: ""2024-01-15"" });
  await commit_excel_batch({ batchId, save: true });
} catch (error) {
  await commit_excel_batch({ batchId, save: false });
}
```

## Summary
- Use batch sessions for **multi-operation workflows** (2× to 10× faster)
- **begin_excel_batch** → operations with batchId → **commit_excel_batch**
- **Always commit or discard** - don't leave batches open
- **One batch per file** - same excelPath for all operations
- **Optional for single operations** - omit batchId for convenience");
    }
    
    /// <summary>
    /// Quick reference for batch session tool parameters.
    /// </summary>
    [McpServerPrompt(Name = "excel_batch_reference"), Description("Quick reference for Excel batch session tool parameters and best practices")]
    public static ChatMessage BatchSessionReference()
    {
        return new ChatMessage(ChatRole.User, @"# Excel Batch Session Quick Reference

## Tools

### begin_excel_batch
**Purpose:** Start a new batch session  
**Parameters:**
- `filePath` (required): Path to Excel file

**Returns:** `{ batchId, filePath, message }`

### commit_excel_batch
**Purpose:** Save and close batch session  
**Parameters:**
- `batchId` (required): Batch session ID from begin_excel_batch
- `save` (optional, default=true): Save workbook before closing

**Returns:** `{ success, batchId, filePath, saved, message }`

### list_excel_batches
**Purpose:** List all active batch sessions (debugging)  
**Parameters:** None

**Returns:** `{ count, activeBatches[] }`

## All Excel Tools Support batchId

Every excel_* tool accepts optional `batchId` parameter:
- **excel_powerquery** - Power Query operations
- **excel_worksheet** - Worksheet operations
- **excel_parameter** - Named range operations
- **excel_cell** - Cell operations
- **excel_vba** - VBA script operations
- **excel_file** - File operations

## Usage Pattern

```typescript
// 1. Begin batch
const { batchId } = await begin_excel_batch({ filePath: ""file.xlsx"" });

// 2. Perform operations (pass batchId)
await excel_powerquery({ batchId, action: ""import"", ... });
await excel_worksheet({ batchId, action: ""create"", ... });

// 3. Commit (save and close)
await commit_excel_batch({ batchId, save: true });
```

## Best Practices

1. **Error Handling:** Always commit in finally block or catch
2. **File Path:** Must match across all operations in batch
3. **Resource Cleanup:** Never abandon batches
4. **Performance:** Use for 2+ operations
5. **Single Operations:** Omit batchId (automatic cleanup)");
    }
}
