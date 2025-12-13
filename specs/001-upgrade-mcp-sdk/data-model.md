# Data Model: Upgrade MCP SDK to 0.5.0-preview.1

**Date**: 2025-12-13  
**Spec**: `specs/001-upgrade-mcp-sdk/spec.md`

---

## Overview

This upgrade is primarily dependency and code migration work; it does not introduce new domain entities stored on disk. The "entities" below are conceptual and document the artifacts used during the upgrade process.

---

## Entity: DependencySet

| Field           | Type     | Description                                       |
|-----------------|----------|---------------------------------------------------|
| PackageId       | string   | NuGet package identifier (e.g., `ModelContextProtocol`) |
| OldVersion      | string   | Currently pinned version                          |
| NewVersion      | string   | Target version (`0.5.0-preview.1`)                |
| BreakingChanges | string[] | List of breaking change identifiers from changelog |

### Validation Rules

- `NewVersion` must be a valid SemVer string.
- `BreakingChanges` populated from SDK changelog.

### Relationships

- One-to-many: DependencySet → ImpactReport (each breaking change → one or more impacted locations).

---

## Entity: ImpactReport

| Field         | Type   | Description                                   |
|---------------|--------|-----------------------------------------------|
| BreakingChange| string | Identifier of a specific breaking change      |
| FilePath      | string | Relative path to impacted file                |
| LineNumber    | int?   | Line number if applicable                     |
| FixApplied    | bool   | Whether the fix has been implemented          |

### Validation Rules

- `FilePath` must exist in repository.
- `FixApplied` = true only after build succeeds post-fix.

### State Transitions

1. Identified (compiler error) → FixApplied = false
2. Fixed → FixApplied = true
3. Verified (build succeeds, tests pass)

---

## Entity: ValidationMatrix

| Field        | Type     | Description                                |
|--------------|----------|--------------------------------------------|
| Criterion    | string   | Success-criterion identifier (SC-*)        |
| Command      | string   | Shell command or test filter               |
| ExpectedOutcome | string | Expected result (e.g., "0 warnings", "exit 0") |
| Passed       | bool     | Whether criterion passed last run          |

### Relationships

- One-to-many: ValidationMatrix → ImpactReport (criterion may verify multiple fixes).

---

## Entity: RollbackPlan

| Field        | Type   | Description                                    |
|--------------|--------|------------------------------------------------|
| Trigger      | string | Condition that triggers rollback               |
| Action       | string | Rollback action (e.g., revert package bump)    |
| Owner        | string | Person responsible for rollback decision       |
| Tested       | bool   | Whether rollback action has been dry-run       |

### Validation Rules

- Each critical path must have a rollback plan entry.

---

## Diagram (Conceptual)

```text
DependencySet ─────┐
                   │   1..*
                   ▼
              ImpactReport
                   ▲
                   │   verified by
                   │
           ValidationMatrix
                   │
                   │   rollback via
                   ▼
              RollbackPlan
```

---

## Notes

- These entities exist as mental/document models; no database or runtime persistence is required.
- Use this model when populating `tasks.md` to track fixes and validation checkpoints.
