# begin_excel_batch Tool

**Purpose**: Start a batch session for multi-operation workflows (75-95% faster)

**When to use begin_excel_batch**:
- 2+ operations on same Excel file
- Detect keywords: numbers, plurals, lists in user request
- Performance critical workflows
- Returns batchId for subsequent operations

**Server-specific behavior**:
- Opens workbook once, keeps in memory
- Returns unique batchId string
- One batch per file (cannot create duplicate batches)
- Must call commit_excel_batch when done to release resources
- Normalized file paths prevent duplicate sessions

**Usage pattern**:
1. begin_excel_batch(excelPath) → { batchId: "abc123..." }
2. Pass batchId to ALL subsequent operations
3. commit_excel_batch(batchId, save: true) when done

**Common mistakes**:
- Not calling commit_excel_batch → Resource leak
- Creating batch for single operation → Unnecessary overhead
- Creating multiple batches for same file → Error (only one allowed)
- Losing batchId → Cannot complete session

**Keyword detection (when to use)**:
- "import 4 queries" → 4 = batch mode
- "create measures for Sales, Revenue, Profit" → plural + list = batch mode
- "add parameters: StartDate, EndDate, Region" → list = batch mode
- Single operation → NO batch mode

**Workflow optimization**:
- Always detect batch opportunity BEFORE first operation
- Save batchId immediately for all subsequent calls
- Use commit_excel_batch(save: false) to discard changes
