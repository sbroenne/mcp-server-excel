# excel_batch Tool

**Actions**: begin, commit, list

**When to use excel_batch**:
- 🚨 **ALWAYS CHECK FIRST**: 2+ operations on same Excel file (75-90% faster)
- 🔍 **AUTO-DETECT keywords**: numbers, plurals, lists in user requests
- ⚡ **Performance-critical**: Must use for any multi-operation workflow
- 📝 **Examples that REQUIRE batch mode**:
  - "create 4 queries" → 4 operations = BATCH
  - "create measures" → plural = BATCH  
  - "add Sales, Revenue, Profit" → list = BATCH
  - "change several things" → multiple = BATCH

**Server-specific behavior**:
- begin: Opens workbook, returns batchId, keeps Excel instance alive
- commit: Saves (or discards) + closes workbook, releases resources
- list: Shows all uncommitted batches (debugging)
- One batch per file (cannot create duplicates)
- MUST commit batches to prevent resource leaks

**Action disambiguation**:
- begin: Returns batchId string - save this for all subsequent operations
- commit with save=true: Persists all changes made during batch
- commit with save=false: Discards all changes (useful for testing/read-only)
- list: Check for forgotten batches (should always be empty in production)

**Common mistakes I make**:
- Forgetting to commit batch → Resource leak (Excel stays open)
- Not saving batchId from begin → Cannot complete session
- Using batch for single operation → Unnecessary overhead
- Creating multiple batches for same file → Error (only one allowed)

**Keyword detection (when to use batch)**:
- "create 4 queries" → number = batch
- "create measures for Sales, Revenue, Profit" → plural + list = batch
- "add parameters: StartDate, EndDate, Region" → list = batch
- Single operation → NO batch

**Workflow pattern**:
```
1. excel_batch(action: 'begin', filePath: 'file.xlsx') 
   → { batchId: "abc123..." }
   
2. excel_powerquery(action: 'create', batchId: 'abc123...')
3. excel_powerquery(action: 'create', batchId: 'abc123...')
4. excel_datamodel(action: 'create-measure', batchId: 'abc123...')

5. excel_batch(action: 'commit', batchId: 'abc123...', save: true)
   → Saves and closes
```

**Performance**:
- Without batch: 4 operations = 8-12 seconds
- With batch: 4 operations = 1-2 seconds
- Savings: 75-90% faster

**Workflow optimization**:
- ALWAYS detect batch opportunity BEFORE first operation
- Use list action if unsure about active batches
- Use save=false for read-only or testing workflows
