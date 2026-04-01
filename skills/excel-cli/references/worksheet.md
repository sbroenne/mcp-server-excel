# worksheet - Cross-File Operations

## Atomic Operations (Recommended)

**copy-to-file** and **move-to-file** are the simplest way to transfer sheets between files.

| Action | Description | Key Parameters |
|--------|-------------|----------------|
| `copy-to-file` | Copy sheet to another file | sourceFile, sheetName, targetFile |
| `move-to-file` | Move sheet to another file | sourceFile, sheetName, targetFile |

**Benefits:**
- No session management required
- Files are opened, modified, saved, and closed automatically
- Single atomic operation - no cleanup needed

**Example - Copy sheet to another file:**
```
action: copy-to-file
sourceFile: C:\Reports\Q1.xlsx
sheetName: Summary
targetFile: C:\Reports\Annual.xlsx
targetName: Q1 Summary  # Optional: rename during copy
```

**Example - Move sheet to another file:**
```
action: move-to-file
sourceFile: C:\Drafts\Data.xlsx
sheetName: FinalData
targetFile: C:\Published\Report.xlsx
beforeSheet: Sheet1  # Optional: position in target
```

## Positioning Parameters

Use `beforeSheet` OR `afterSheet` (not both) to control where the sheet appears in the target file:

- `beforeSheet: "Sheet1"` - Insert before Sheet1
- `afterSheet: "Sheet1"` - Insert after Sheet1
- Neither specified - Append to end

## When to Use Session-Based Operations

For same-file operations (copy within same workbook, rename, delete, tab colors), use session-based actions with `sessionId`.

## Common Errors

| Error | Cause | Solution |
|-------|-------|----------|
| "Source and target files must be different" | Same file for both | Use `copy` action instead |
| "Source file not found" | File doesn't exist | Verify file path |
| "Sheet not found" | Typo in sheet name | Use `list` action to see available sheets |
