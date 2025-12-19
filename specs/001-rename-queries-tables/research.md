# Research: Rename Queries & Data Model Tables

## Goal
Support renaming:
- Power Query queries (workbook queries)
- Excel Data Model tables (Power Pivot model tables)

This repo uses late-bound Excel COM (`dynamic`) via `ExcelMcp.ComInterop` and must follow the Constitution rules (COM release in `finally`, no Core-level try/catch that suppresses exceptions, MCP returns JSON for business errors).

## Decision 1: Power Query rename approach

**Decision**: Rename a Power Query by setting the query object’s `Name` property.

**Rationale**:
- The Core already enumerates queries via `workbook.Queries` and retrieves `query.Name` (see `ComUtilities.FindQuery`).
- Excel query objects are treated as named resources that can be deleted and enumerated; rename via `Name` is the natural COM path.

**Implementation notes**:
- Find query by exact name (existing helper is case-sensitive). The feature requires trim + case-insensitive uniqueness checks, so the rename command should:
  - Trim both old/new names.
  - Perform case-insensitive collision checks against all query names excluding the target.
  - Treat “trim-equal” rename as a no-op success.
  - Attempt COM rename even for case-only differences.
- COM objects involved: `Queries` collection + `Query` object. Release both.

**Alternatives considered**:
- Delete + recreate query: rejected (would change lineage, potentially load settings).
- UI automation: rejected (non-goal; brittle, interactive).

## Decision 2: Data Model table rename feasibility

**Decision**: Treat Data Model table rename as **not directly supported** by the documented Excel COM PowerPivot `ModelTable` object model. Implement rename as:
1) Attempt a direct rename via COM (best-effort: `modelTable.Name = newName`) to catch any Excel versions that allow it through late binding.
2) If that fails because the object is read-only, fall back to **source-based renaming** when possible:
   - If the `ModelTable` is backed by a Power Query connection (common: “Query - {QueryName} …”), rename the underlying Power Query and verify the model table list reflects the new name.
   - Otherwise, return a clear “not supported for this table source” error.

**Rationale**:
- Microsoft’s PowerPivot model object model documentation explicitly describes `ModelTable` as read-only (cannot be created/edited) and lists `Name` as read-only.
- The repo already relies on the relationship between Power Query naming and ModelTable naming (e.g., `IsQueryLoadedToDataModel` checks model table names containing the query name).

**Alternatives considered**:
- Delete + recreate Data Model table under new name: rejected (would almost certainly break relationships/measures; violates “safe rename” intent).
- Renaming `SourceWorkbookConnection.Name` only: uncertain; may not rename the model table display name.

## Risk register

### R1: Data Model rename may be impossible for non-PQ sources
- Impact: User Story 2 cannot be fully satisfied for all table origins.
- Mitigation:
  - Scope rename to PQ-backed model tables where rename can be achieved by renaming the query.
  - Provide explicit error messaging when the table is not renameable.
  - Consider updating the feature spec to reflect this limitation if confirmed by testing.

### R2: Case-only rename behavior varies
- Excel may reject case-only renames or treat them as no-op.
- Mitigation: always attempt COM rename; treat Excel outcome as authoritative.

### R3: Name normalization in existing helpers is case-sensitive
- `ComUtilities.FindQuery` uses `==` match.
- Mitigation: implement a feature-specific lookup that matches the spec’s trim + case-insensitive semantics.
