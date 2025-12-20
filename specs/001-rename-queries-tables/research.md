# Research: Rename Queries & Data Model Tables

## Goal
Support renaming:
- Power Query queries (workbook queries)
- Excel Data Model tables (Power Pivot model tables)

This repo uses late-bound Excel COM (`dynamic`) via `ExcelMcp.ComInterop` and must follow the Constitution rules (COM release in `finally`, no Core-level try/catch that suppresses exceptions, MCP returns JSON for business errors).

## Decision 1: Power Query rename approach

**Decision**: Rename a Power Query by setting the query object's `Name` property.

**Rationale**:
- The Core already enumerates queries via `workbook.Queries` and retrieves `query.Name` (see `ComUtilities.FindQuery`).
- Excel query objects are treated as named resources that can be deleted and enumerated; rename via `Name` is the natural COM path.

**Implementation notes**:
- Find query by exact name (existing helper is case-sensitive). The feature requires trim + case-insensitive uniqueness checks, so the rename command should:
  - Trim both old/new names.
  - Perform case-insensitive collision checks against all query names excluding the target.
  - Treat "trim-equal" rename as a no-op success.
  - Attempt COM rename even for case-only differences.
- COM objects involved: `Queries` collection + `Query` object. Release both.

**Alternatives considered**:
- Delete + recreate query: rejected (would change lineage, potentially load settings).
- UI automation: rejected (non-goal; brittle, interactive).

## Decision 2: Data Model table rename feasibility

**Decision**: Data Model table rename is **NOT POSSIBLE** via the Excel COM API.

**VERIFIED via COM API testing (2025-12-20)**:

The following approaches were tested and **ALL FAILED** to change `ModelTable.Name`:

| Approach | Result |
|----------|--------|
| Direct `table.Name = newName` | Throws `TargetParameterCountException` (HResult: 0x8002000E) |
| Rename underlying Power Query | ModelTable.Name unchanged |
| Rename Connection | ModelTable.Name unchanged |
| Update connection string `Location=` parameter | ModelTable.Name unchanged |
| Call `model.Refresh()` after PQ rename | ModelTable.Name unchanged |
| Save workbook and reopen | ModelTable.Name unchanged |

**Conclusion**: `ModelTable.Name` is immutable. It is cached at table creation time from the connection/source name and cannot be changed afterwards through any COM API operation.

**Implementation**:
1. Attempt a direct rename via COM (`modelTable.Name = newName`) for future-proofing.
2. When it fails (expected), return clear error explaining the limitation.
3. User workaround: Delete the table and recreate it with the new name (this also deletes associated measures).

**Rationale**:
- Microsoft's PowerPivot model object model documentation explicitly describes `ModelTable` as read-only (cannot be created/edited) and lists `Name` as read-only.
- Empirical testing confirms no workaround exists via COM API.

**Alternatives rejected**:
- Delete + recreate Data Model table under new name: Would break relationships/measures; violates "safe rename" intent.
- Renaming Power Query as fallback: Verified does NOT change ModelTable.Name.
- UI automation: Non-goal; brittle, interactive.

## Risk register

### ~~R1: Data Model rename may be impossible for non-PQ sources~~ RESOLVED
- **Resolution**: Data Model rename is impossible for ALL sources, not just non-PQ.
- ModelTable.Name is immutable regardless of table origin.
- Feature spec User Story 2 cannot be satisfied; documented as a known Excel limitation.

### R2: Case-only rename behavior varies
- Excel may reject case-only renames or treat them as no-op.
- Mitigation: always attempt COM rename; treat Excel outcome as authoritative.

### R3: Name normalization in existing helpers is case-sensitive
- `ComUtilities.FindQuery` uses `==` match.
- Mitigation: implement a feature-specific lookup that matches the spec's trim + case-insensitive semantics.
