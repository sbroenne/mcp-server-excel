# worksheet - Worksheet Operations

## Same-File Session Operations

Use session-based actions for worksheet lifecycle within the same workbook:

| Action | Parameters |
|--------|------------|
| `create` | `sheet_name` |
| `rename` | `old_name`, `new_name` |
| `delete` | `sheet_name` |
| `move` | `sheet_name`, `before_sheet`/`after_sheet` |
| `copy` | `source_name`, `target_name` |

**Rename example:**
```
action: rename
old_name: Sheet1
new_name: Summary
```

Rename requires `old_name` + `new_name`.

## Atomic Cross-File Operations

**copy-to-file** and **move-to-file** are the simplest way to transfer sheets between files.

| Action | Description | Key Parameters |
|--------|-------------|----------------|
| `copy-to-file` | Copy sheet to another file | `source_file`, `source_sheet`, `target_file` |
| `move-to-file` | Move sheet to another file | `source_file`, `source_sheet`, `target_file` |

**Benefits:**
- No session management required
- Files are opened, modified, saved, and closed automatically
- Single atomic operation - no cleanup needed

**Example - Copy sheet to another file:**
```
action: copy-to-file
source_file: C:\Reports\Q1.xlsx
source_sheet: Summary
target_file: C:\Reports\Annual.xlsx
target_sheet_name: Q1 Summary  # Optional: rename during copy
```

**Example - Move sheet to another file:**
```
action: move-to-file
source_file: C:\Drafts\Data.xlsx
source_sheet: FinalData
target_file: C:\Published\Report.xlsx
before_sheet: Sheet1  # Optional: position in target
```

## Positioning Parameters

Use `before_sheet` OR `after_sheet` (not both) to control where the sheet appears in the target file:

- `before_sheet: "Sheet1"` - Insert before Sheet1
- `after_sheet: "Sheet1"` - Insert after Sheet1
- Neither specified - Append to end

## When to Use Session-Based Operations

For same-file operations (copy within same workbook, rename, delete, tab colors), use session-based actions with `session_id`.

## Rename Parameters

For `rename`, use `old_name` and `new_name`.

- MCP rename requires `old_name` + `new_name`
- CLI uses `--old-name` + `--new-name`
- Copy and cross-file parameters such as `sheet_name`, `source_name`, `source_sheet`, `target_name`, and `target_sheet_name` are not rename aliases

## Common Errors

| Error | Cause | Solution |
|-------|-------|----------|
| "Source and target files must be different" | Same file for both | Use `copy` action instead |
| "Source file not found" | File doesn't exist | Verify file path |
| "Sheet not found" | Typo in sheet name | Use `list` action to see available sheets |
