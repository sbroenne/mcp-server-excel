# Data Model: Rename Queries & Data Model Tables

## Entities

### RenameRequest (Power Query)
- `objectType`: `power-query`
- `oldName`: string (required)
- `newName`: string (required)

### RenameRequest (Data Model Table)
- `objectType`: `data-model-table`
- `oldName`: string (required)
- `newName`: string (required)

### RenameResult
- `success`: boolean
- `errorMessage`: string | null
- `objectType`: `power-query` | `data-model-table`
- `oldName`: string
- `newName`: string
- `normalizedOldName`: string (trimmed)
- `normalizedNewName`: string (trimmed)

## Validation Rules

### Name normalization
- Trim leading/trailing whitespace for comparison.
- Case-insensitive uniqueness checks within the same scope.

### No-op
- If `Trim(oldName)` equals `Trim(newName)`, treat as no-op success (do not call COM).

### Case-only rename
- Allowed; still attempt COM rename (Excel decides).

### Conflicts
- `newName` must not already exist in the target scope (queries vs model tables), excluding the renamed object.

## Relationships
- Power Query names and Data Model table names are related when a query is loaded to the Data Model.
- If Data Model tables cannot be renamed directly via COM, the rename operation may be achievable only for PQ-backed tables by renaming the underlying query.

## State Transitions
- `requested` → `validated` → `attempted` → `succeeded | failed`
- On failure, workbook state must remain consistent (Excel COM must reject invalid renames; no partial state mutations beyond Excel’s atomic rename).
