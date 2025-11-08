# excel_batch - Server Quirks

**Auto-detect batch opportunities** (keywords in user requests):
- Numbers: "create 4 queries" → BATCH
- Plurals: "create measures" → BATCH
- Lists: "add Sales, Revenue, Profit" → BATCH
- Multiple: "change several things" → BATCH

**Server-specific quirks**:
- One batch per file (no duplicates allowed)
- MUST commit to prevent Excel resource leaks
- save=false: Discards changes (testing/read-only workflows)
- list action: Debug forgotten batches (should be empty in production)
