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
        return new ChatMessage(ChatRole.User, @"When working with Excel files, you have access to a batch mode that makes multiple operations 95% faster (2-5 seconds per operation → <100ms per operation).

# When to Use Batch Mode

Use batch mode when performing 2+ operations on the same Excel file:
- Importing multiple Power Queries
- Creating multiple Data Model measures
- Creating multiple worksheets or tables
- Any combination of multiple operations on the same file

Do NOT use batch mode for single operations or when working with different files.

# The Three-Step Pattern

1. First call begin_excel_batch with the filePath
2. Pass the returned batchId to all subsequent tool calls on that file
3. Finally call commit_excel_batch with save: true

Example workflow for importing 3 Power Queries:

First, start the batch:
begin_excel_batch(filePath: ""workbook.xlsx"")
→ Returns { batchId: ""abc123..."" }

Then perform all operations, passing the batchId each time:
excel_powerquery(action: ""import"", excelPath: ""workbook.xlsx"", queryName: ""Sales"", sourcePath: ""sales.pq"", batchId: ""abc123..."")
excel_powerquery(action: ""import"", excelPath: ""workbook.xlsx"", queryName: ""Products"", sourcePath: ""products.pq"", batchId: ""abc123..."")
excel_powerquery(action: ""import"", excelPath: ""workbook.xlsx"", queryName: ""Customers"", sourcePath: ""customers.pq"", batchId: ""abc123..."")

Finally, commit to save:
commit_excel_batch(batchId: ""abc123..."", save: true)

# Critical Rules

1. Always call commit_excel_batch when done (don't leave batches hanging)
2. Use the same batchId for all operations on the same file
3. One batch per workbook (can't share batches across different files)
4. Set save: true to save changes, save: false to discard changes

# Recognizing Multi-Operation Requests

When the user asks to:
- ""Import these 5 queries"" → Use batch mode
- ""Create measures for TotalSales, AvgPrice, ProductCount"" → Use batch mode
- ""Set up a workbook with 3 sheets and data"" → Use batch mode
- ""Just import this one query"" → Don't use batch mode (single operation)

# Response Hints

If you see messages in responses like:
- ""For multiple imports: Use begin_excel_batch to group operations efficiently""
- ""Creating multiple measures? Use begin_excel_batch to keep Data Model open""

These hints mean you should have used batch mode but didn't. On the next similar request, use the batch pattern from the start.");
    }
}
