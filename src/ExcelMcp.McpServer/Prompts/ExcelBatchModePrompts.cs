using System.ComponentModel;
using Microsoft.Extensions.AI;
using ModelContextProtocol.Server;

namespace Sbroenne.ExcelMcp.McpServer.Prompts;

/// <summary>
/// Essential guide for using Excel batch mode to achieve 95% faster operations.
/// </summary>
[McpServerPromptType]
public static class ExcelBatchModePrompts
{
    /// <summary>
    /// Essential guide: When and how to use Excel batch mode for 95% faster operations.
    /// </summary>
    [McpServerPrompt(Name = "excel_batch_mode_guide")]
    [Description("Essential guide: When and how to use Excel batch mode for 95% faster operations")]
    public static ChatMessage BatchModeGuide()
    {
        return new ChatMessage(ChatRole.User, @"# Excel Batch Mode - Performance Critical Pattern

## ⚡ Performance Impact: 95% Faster

**Without Batch**: 2-5 seconds per operation (Excel startup overhead every time)
**With Batch**: <100ms per operation (workbook stays open, instance reused)

## 🎯 When to Use Batch Mode

**ALWAYS use batch mode when performing 2+ operations on the same workbook:**

✅ **USE BATCH MODE:**
- Importing/updating **2+ Power Queries**
- Creating/updating **2+ Data Model measures**
- Creating **2+ worksheets**
- Creating **2+ tables**
- Any combination of multiple operations on same file
- User says ""create a workbook with..."" or ""set up..."" or ""import these queries...""

❌ **Skip batch mode:**
- Single operation on a file (e.g., ""just import this one query"")
- Operations on different workbooks (one batch per workbook)

**Pattern Recognition Examples:**
- ""Import these 5 queries"" → **USE BATCH MODE**
- ""Create measures for TotalSales, AvgPrice, ProductCount"" → **USE BATCH MODE**
- ""Set up a dashboard with 3 sheets and 2 queries"" → **USE BATCH MODE**
- ""Just import SalesData query"" → No batch needed (single operation)

## 📋 Three-Step Pattern

### Step 1: Start Batch Session
```typescript
const batch = await begin_excel_batch({ 
  filePath: ""workbook.xlsx"" 
});
const batchId = batch.batchId;  // Save this ID!
```

### Step 2: Execute Multiple Operations (Super Fast!)
```typescript
// All these operations use the SAME open workbook
await excel_powerquery({ 
  action: ""import"", 
  excelPath: ""workbook.xlsx"",
  queryName: ""SalesData"",
  sourcePath: ""sales.pq"",
  batchId: batchId  // ← Pass batch ID here!
});

await excel_powerquery({ 
  action: ""import"", 
  excelPath: ""workbook.xlsx"",
  queryName: ""CustomersData"",
  sourcePath: ""customers.pq"",
  batchId: batchId  // ← Same batch ID!
});

await excel_datamodel({ 
  action: ""create-measure"", 
  excelPath: ""workbook.xlsx"",
  tableName: ""Sales"",
  measureName: ""TotalRevenue"",
  daxFormula: ""SUM(Sales[Amount])"",
  batchId: batchId  // ← Same batch ID!
});
```

### Step 3: Commit and Save
```typescript
await commit_excel_batch({ 
  batchId: batchId, 
  save: true  // true = save changes, false = discard
});
```

## ⚠️ Critical Rules

1. **One batch per workbook** - Each batch is tied to a specific file
2. **Always commit** - Don't leave batches hanging (60s auto-cleanup)
3. **Pass batchId to every tool** - All excel_* tools accept optional batchId parameter
4. **Match excelPath** - Must use same file path as begin_excel_batch
5. **save: true by default** - Only use save: false to intentionally discard all changes

## 🔄 Complete Workflow Example

```typescript
// User: ""Create a sales workbook with 3 queries and 2 measures""

// 1. Start batch (opens Excel once)
const batch = await begin_excel_batch({ filePath: ""sales.xlsx"" });
const bid = batch.batchId;

// 2. Import queries (all super fast, no Excel restarts!)
await excel_powerquery({ action: ""import"", batchId: bid, queryName: ""Sales"", sourcePath: ""sales.pq"", excelPath: ""sales.xlsx"" });
await excel_powerquery({ action: ""import"", batchId: bid, queryName: ""Products"", sourcePath: ""products.pq"", excelPath: ""sales.xlsx"" });
await excel_powerquery({ action: ""import"", batchId: bid, queryName: ""Customers"", sourcePath: ""customers.pq"", excelPath: ""sales.xlsx"" });

// 3. Create measures (still in same batch!)
await excel_datamodel({ action: ""create-measure"", batchId: bid, tableName: ""Sales"", measureName: ""TotalRevenue"", daxFormula: ""SUM(Sales[Amount])"", excelPath: ""sales.xlsx"" });
await excel_datamodel({ action: ""create-measure"", batchId: bid, tableName: ""Sales"", measureName: ""AvgOrderValue"", daxFormula: ""AVERAGE(Sales[Amount])"", excelPath: ""sales.xlsx"" });

// 4. Commit (saves and closes)
await commit_excel_batch({ batchId: bid, save: true });

// Result: 5 operations completed in ~6-7 seconds instead of ~15-25 seconds!
```

## 💡 API Response Hints

If you see messages like these in tool responses:

> ""For multiple imports: Use begin_excel_batch to group operations efficiently""
> ""Creating multiple measures? Use begin_excel_batch to keep Data Model open (much faster)""
> ""Creating multiple sheets? Use begin_excel_batch for complete workbook setup""

**You forgot to use batch mode!** The API is reminding you to use batching for better performance.

## 🚫 Common Mistakes

### ❌ Wrong: Not using batch mode for multiple operations
```typescript
// This is SLOW (Excel opens/closes 3 times)
await excel_powerquery({ action: ""import"", queryName: ""Q1"", ... });
await excel_powerquery({ action: ""import"", queryName: ""Q2"", ... });
await excel_powerquery({ action: ""import"", queryName: ""Q3"", ... });
```

### ✅ Correct: Use batch mode
```typescript
const batch = await begin_excel_batch({ filePath: ""workbook.xlsx"" });
await excel_powerquery({ action: ""import"", queryName: ""Q1"", batchId: batch.batchId, ... });
await excel_powerquery({ action: ""import"", queryName: ""Q2"", batchId: batch.batchId, ... });
await excel_powerquery({ action: ""import"", queryName: ""Q3"", batchId: batch.batchId, ... });
await commit_excel_batch({ batchId: batch.batchId, save: true });
```

### ❌ Wrong: Forgetting to commit
```typescript
const batch = await begin_excel_batch({ filePath: ""workbook.xlsx"" });
await excel_powerquery({ action: ""import"", batchId: batch.batchId, ... });
// ❌ Missing commit_excel_batch!
// Batch will auto-cleanup after 60s but changes may be lost!
```

### ✅ Correct: Always commit
```typescript
const batch = await begin_excel_batch({ filePath: ""workbook.xlsx"" });
await excel_powerquery({ action: ""import"", batchId: batch.batchId, ... });
await commit_excel_batch({ batchId: batch.batchId, save: true });  // ✅ Always commit!
```

## 🎓 Decision Tree

```
Are you performing 2+ operations on the same Excel file?
├─ YES → Use batch mode (begin_excel_batch → operations with batchId → commit_excel_batch)
└─ NO → Skip batch mode (just call the tool directly)
```

## 📊 Performance Comparison

| Scenario | Without Batch | With Batch | Speedup |
|----------|---------------|------------|---------|
| Import 5 queries | ~15 seconds | ~1 second | 15x faster |
| Create 10 measures | ~25 seconds | ~1.5 seconds | 17x faster |
| Setup workbook (3 sheets + 2 queries + 2 measures) | ~20 seconds | ~2 seconds | 10x faster |

**Bottom Line:** Batch mode is a game-changer for multi-operation workflows. Always use it when doing 2+ things to the same file!");
    }
}
