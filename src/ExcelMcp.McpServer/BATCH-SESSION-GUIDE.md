# Excel Batch Session Management for LLMs

## 🎯 What Are Batch Sessions?

Batch sessions allow you to perform **multiple Excel operations efficiently** by keeping the workbook open across operations instead of reopening it every time.

### Performance Benefits
- **Without batch**: Each operation opens Excel (2-5 seconds) → operation → closes Excel
- **With batch**: Open Excel once → multiple operations (milliseconds each) → close once
- **Speed improvement**: 10-100x faster for multi-step workflows

---

## 📖 Complete Workflow Guide

### Step 1: Begin a Batch Session

```json
{
  "tool": "excel_batch_begin",
  "arguments": {
    "excelPath": "C:\\Users\\user\\data\\report.xlsx"
  }
}
```

**Returns:**
```json
{
  "batchId": "batch_abc123",
  "excelPath": "C:\\Users\\user\\data\\report.xlsx",
  "status": "active",
  "message": "Batch session started. Excel workbook is open and ready."
}
```

**Important**: Save the `batchId` - you'll need it for all subsequent operations.

---

### Step 2: Perform Operations with BatchId

All Excel tools now accept an optional `batchId` parameter:

#### Example: Update Power Query in Batch
```json
{
  "tool": "excel_powerquery",
  "arguments": {
    "batchId": "batch_abc123",
    "action": "update",
    "queryName": "SalesData",
    "mCodeFile": "query.pq"
  }
}
```

#### Example: Read Worksheet in Same Batch
```json
{
  "tool": "excel_worksheet",
  "arguments": {
    "batchId": "batch_abc123",
    "action": "read",
    "sheetName": "Summary",
    "range": "A1:D10"
  }
}
```

#### Example: Set Parameter in Same Batch
```json
{
  "tool": "excel_parameter",
  "arguments": {
    "batchId": "batch_abc123",
    "action": "set",
    "parameterName": "ReportDate",
    "value": "2025-10-26"
  }
}
```

**Key Point**: All operations use the **same open workbook** - no reopening between operations.

---

### Step 3: Save Changes (Optional)

You can save at any point during the batch:

```json
{
  "tool": "excel_batch_save",
  "arguments": {
    "batchId": "batch_abc123"
  }
}
```

**Returns:**
```json
{
  "batchId": "batch_abc123",
  "saved": true,
  "message": "Workbook saved successfully."
}
```

**When to save:**
- Before risky operations (create checkpoint)
- After completing a logical group of changes
- **Auto-saves on commit** if you don't save explicitly

---

### Step 4: End the Batch Session

**Always end your batch when done:**

```json
{
  "tool": "excel_batch_commit",
  "arguments": {
    "batchId": "batch_abc123"
  }
}
```

**Returns:**
```json
{
  "batchId": "batch_abc123",
  "status": "committed",
  "autoSaved": true,
  "message": "Batch session committed. Workbook saved and closed."
}
```

**What happens:**
1. Saves workbook (if not already saved)
2. Closes workbook
3. Releases Excel resources
4. Invalidates batchId (can't be reused)

---

## 🔄 Complete Example Workflow

**Scenario**: Update report with new data, refresh Power Query, read results

```javascript
// 1. START BATCH
const batch = await excel_batch_begin({
  excelPath: "C:\\Reports\\monthly.xlsx"
});
// Returns: { batchId: "batch_xyz789", ... }

// 2. UPDATE PARAMETER (in batch)
await excel_parameter({
  batchId: batch.batchId,
  action: "set",
  parameterName: "MonthEnd",
  value: "2025-10-31"
});

// 3. REFRESH POWER QUERY (in batch)
await excel_powerquery({
  batchId: batch.batchId,
  action: "refresh",
  queryName: "SalesData"
});

// 4. READ RESULTS (in batch)
const data = await excel_worksheet({
  batchId: batch.batchId,
  action: "read",
  sheetName: "Dashboard",
  range: "A1:F100"
});

// 5. SAVE CHECKPOINT
await excel_batch_save({
  batchId: batch.batchId
});

// 6. MORE OPERATIONS...
await excel_cell({
  batchId: batch.batchId,
  action: "set-value",
  sheetName: "Summary",
  cellAddress: "A1",
  value: "Report Updated: 2025-10-26"
});

// 7. COMMIT (auto-saves and closes)
await excel_batch_commit({
  batchId: batch.batchId
});
```

**Total time:** ~3 seconds (vs. ~20+ seconds without batching)

---

## 🆚 Batch vs. Non-Batch Operations

### Without BatchId (Legacy/Simple Operations)

Each tool call opens and closes Excel:

```json
{
  "tool": "excel_powerquery",
  "arguments": {
    "excelPath": "C:\\data\\report.xlsx",
    "action": "view",
    "queryName": "Sales"
  }
}
```

**Behavior:**
- Opens workbook
- Performs operation
- **Auto-saves if changes were made**
- Closes workbook
- ~2-5 seconds overhead per call

**Use when:**
- Single operation needed
- Quick read-only queries
- Backward compatibility

### With BatchId (Recommended for Multi-Step)

Operations reuse open workbook:

```json
{
  "tool": "excel_powerquery",
  "arguments": {
    "batchId": "batch_abc123",
    "action": "view",
    "queryName": "Sales"
  }
}
```

**Behavior:**
- Uses already-open workbook
- Performs operation
- **Workbook stays open** (no auto-save)
- ~100ms per call

**Use when:**
- Multiple related operations
- Performance matters
- Transactional workflows

---

## 🛡️ Error Handling

### Batch Errors

If an operation fails mid-batch:

```json
{
  "error": "Query 'InvalidName' not found",
  "batchId": "batch_abc123",
  "batchStatus": "active"
}
```

**Your options:**
1. **Continue**: Fix the error and try other operations
2. **Save**: Save progress so far with `excel_batch_save`
3. **Abandon**: Call `excel_batch_commit` to close without saving more

### Automatic Cleanup

If you forget to commit:
- Batch sessions auto-expire after 5 minutes of inactivity
- Workbook closes WITHOUT saving uncommitted changes
- Resources are released automatically

**Best Practice**: Always commit your batches explicitly.

---

## 📊 Batch Session Status

Check batch status anytime:

```json
{
  "tool": "excel_batch_status",
  "arguments": {
    "batchId": "batch_abc123"
  }
}
```

**Returns:**
```json
{
  "batchId": "batch_abc123",
  "status": "active",
  "excelPath": "C:\\Reports\\monthly.xlsx",
  "operationCount": 7,
  "lastOperation": "2025-10-26T10:30:45Z",
  "hasUnsavedChanges": true
}
```

---

## 🎓 Best Practices

### ✅ DO:
- **Start batch** for any workflow with 2+ operations
- **Save periodically** during long workflows (create checkpoints)
- **Commit when done** - always clean up resources
- **Check status** if uncertain about batch state
- **Use meaningful paths** - absolute paths work best

### ❌ DON'T:
- **Mix batch and non-batch** on same file (can cause conflicts)
- **Forget to commit** - resources will leak temporarily
- **Reuse batchId** after commit (it's invalidated)
- **Keep batches open indefinitely** (5-minute timeout)
- **Assume auto-save** during batch (only on commit)

---

## 🔍 Common Patterns

### Pattern 1: Read-Only Analysis
```javascript
const batch = await excel_batch_begin({ excelPath: "data.xlsx" });

// Multiple reads
const sheet1 = await excel_worksheet({ batchId: batch.batchId, action: "read", sheetName: "Sales" });
const sheet2 = await excel_worksheet({ batchId: batch.batchId, action: "read", sheetName: "Costs" });
const params = await excel_parameter({ batchId: batch.batchId, action: "list" });

// No save needed - read-only
await excel_batch_commit({ batchId: batch.batchId });
```

### Pattern 2: Transactional Updates
```javascript
const batch = await excel_batch_begin({ excelPath: "report.xlsx" });

try {
  // Update parameters
  await excel_parameter({ batchId: batch.batchId, action: "set", parameterName: "Date", value: "2025-10-26" });
  
  // Refresh data
  await excel_powerquery({ batchId: batch.batchId, action: "refresh", queryName: "Sales" });
  
  // Update summary
  await excel_cell({ batchId: batch.batchId, action: "set-value", sheetName: "Summary", cellAddress: "A1", value: "Updated" });
  
  // Commit (auto-saves)
  await excel_batch_commit({ batchId: batch.batchId });
} catch (error) {
  // On error, commit anyway to close workbook
  await excel_batch_commit({ batchId: batch.batchId });
  throw error;
}
```

### Pattern 3: Incremental Save
```javascript
const batch = await excel_batch_begin({ excelPath: "large.xlsx" });

// Phase 1
await excel_powerquery({ batchId: batch.batchId, action: "update", queryName: "Q1", mCodeFile: "q1.pq" });
await excel_batch_save({ batchId: batch.batchId }); // Checkpoint

// Phase 2
await excel_powerquery({ batchId: batch.batchId, action: "update", queryName: "Q2", mCodeFile: "q2.pq" });
await excel_batch_save({ batchId: batch.batchId }); // Checkpoint

// Done
await excel_batch_commit({ batchId: batch.batchId });
```

---

## 📋 Quick Reference

| Tool | Purpose | When to Use |
|------|---------|-------------|
| `excel_batch_begin` | Start session | Beginning of multi-operation workflow |
| `excel_batch_save` | Save changes | Create checkpoints during workflow |
| `excel_batch_commit` | End session | Always at end of workflow |
| `excel_batch_status` | Check state | Debugging or long-running workflows |
| All `excel_*` tools | Operations | Include `batchId` for batch operations |

---

## 🚀 Migration from Non-Batch Code

**Before (slow - reopens Excel 3 times):**
```javascript
await excel_parameter({ excelPath: "file.xlsx", action: "set", parameterName: "Date", value: "2025-10-26" });
await excel_powerquery({ excelPath: "file.xlsx", action: "refresh", queryName: "Sales" });
await excel_worksheet({ excelPath: "file.xlsx", action: "read", sheetName: "Summary" });
```

**After (fast - opens Excel once):**
```javascript
const batch = await excel_batch_begin({ excelPath: "file.xlsx" });
await excel_parameter({ batchId: batch.batchId, action: "set", parameterName: "Date", value: "2025-10-26" });
await excel_powerquery({ batchId: batch.batchId, action: "refresh", queryName: "Sales" });
await excel_worksheet({ batchId: batch.batchId, action: "read", sheetName: "Summary" });
await excel_batch_commit({ batchId: batch.batchId });
```

**Performance gain:** 10-20 seconds saved

---

## 💡 Key Takeaways

1. **Batches = Performance**: Use for any workflow with multiple operations
2. **LLM Controls Lifecycle**: You decide when to start, save, and commit
3. **Backward Compatible**: Non-batch operations still work (just slower)
4. **Always Commit**: Clean up resources explicitly
5. **Auto-Save on Commit**: Changes are saved unless you abandon batch

**Remember**: Batch sessions are a powerful optimization - use them whenever you have related Excel operations to perform!
