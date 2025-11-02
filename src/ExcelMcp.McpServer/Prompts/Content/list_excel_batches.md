# list_excel_batches Tool

**Actions**: None (single purpose tool)

**When to use list_excel_batches**:
- Debugging batch session issues
- Check which files have uncommitted batches
- Verify no resource leaks from forgotten commits
- Use commit_excel_batch to close active sessions

**Server-specific behavior**:
- Lists all active batch sessions with batchId and filePath
- Active batches hold Excel instances open (resource consumption)
- Always commit batches when done to release resources
- Each batchId is unique per session

**Common use cases**:
- Forgotten to commit batch → Check with this tool
- Multiple batches accidentally created → List and commit all
- Resource leak investigation → Find unclosed sessions

**Common mistakes**:
- Not committing batches after use → Resource leaks
- Creating multiple batches for same file → Only one batch per file allowed

**Workflow optimization**:
- Use periodically during development to check for leaks
- Always commit batches in production workflows
- Pattern: begin → operations → commit (never forget commit)
