# commit_excel_batch Tool

**Purpose**: End batch session, save/discard changes, release resources

**When to use commit_excel_batch**:
- ALWAYS after begin_excel_batch (REQUIRED)
- When all batch operations complete
- To save changes: commit_excel_batch(batchId, save: true)
- To discard changes: commit_excel_batch(batchId, save: false)

**Server-specific behavior**:
- Saves workbook if save=true (default)
- Closes workbook and releases Excel instance
- Removes batch from active sessions
- Prevents resource leaks
- Cannot commit same batch twice

**Parameters**:
- batchId: Required (from begin_excel_batch)
- save: Optional, default true (false = discard all changes)

**Common mistakes**:
- Forgetting to commit batch → Resource leak, Excel stays open
- Committing with wrong batchId → Error
- Not using save=false when testing → Unwanted changes saved
- Committing batch twice → Error

**Error scenarios**:
- If save fails: Batch still disposed to prevent resource leak
- If batchId not found: Already committed or invalid

**Workflow optimization**:
- Always pair with begin_excel_batch (begin → operations → commit)
- Use save=false for read-only workflows or testing
- Check with list_excel_batches if unsure about active sessions
