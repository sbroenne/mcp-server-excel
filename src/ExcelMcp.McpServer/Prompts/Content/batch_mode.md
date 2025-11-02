# Batch Mode Workflow

**Tools**: begin_excel_batch, commit_excel_batch

**When to use batch mode**:
- 2+ operations on same Excel file
- Keywords in request: numbers ("import 4 queries"), plurals ("create measures"), lists
- Performance critical: 75-95% faster than individual operations

**Server-specific behavior**:
- begin_excel_batch: Returns batchId, opens Excel once
- commit_excel_batch: Saves or discards, closes Excel
- All operations between begin/commit share same Excel instance
- One batch per workbook (cannot share batches across files)

**Batch workflow pattern**:
1. begin_excel_batch(excelPath) → Returns { batchId: "abc123" }
2. Pass batchId to ALL subsequent operations
3. commit_excel_batch(batchId, save: true) → Persists changes

**Keyword triggers (AUTO-DETECT)**:
- Numbers: "import 4 queries", "create 5 parameters", "3 measures"
- Plurals: "queries", "parameters", "measures", "worksheets"
- Lists: User provides enumerated items
- Repetitive: "each", "all", "every"

**Performance comparison**:
- Without batch: 4 operations = 8-12 seconds (2-3s per operation)
- With batch: 4 operations = 1-2 seconds (one Excel session)
- Savings: 75-90% faster

**Common mistakes**:
- Not detecting batch opportunities → Look for keywords upfront
- Forgetting to commit batch → Changes not saved
- Using batch for single operation → Unnecessary overhead
- Not passing batchId to all operations → Breaks batch session

**Workflow optimization**:
- ALWAYS detect batch keywords BEFORE first operation
- Use commit_excel_batch with save: true to persist
- Use commit_excel_batch with save: false to discard
