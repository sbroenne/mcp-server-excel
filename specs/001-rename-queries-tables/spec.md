# Feature Specification: Rename Queries & Tables

**Feature Branch**: `001-rename-queries-tables`  
**Created**: 2025-12-19  
**Status**: Draft  
**Input**: User description: "Add ability to rename existing Power Query queries and Excel Data Model tables"

## Clarifications

### Session 2025-12-19

- Q: When renaming a Data Model table, what should we do if there are dependencies (relationships, measures, calculated fields)? → A: Attempt the rename via COM; if Excel succeeds, assume dependencies are preserved; if Excel rejects/throws, fail with the Excel error.
- Q: For rename conflict detection (Power Query names and Data Model table names), what should the uniqueness rule be? → A: Case-insensitive uniqueness + trim whitespace.
- Q: Where do you want the rename capabilities exposed? → A: MCP tool + Core + CLI commands.
- Q: After a successful rename, should the system automatically save the workbook? → A: Never auto-save; caller decides when to save/commit.
- Q: Should case-only renames (e.g., "Sales" → "sales") be allowed? → A: Yes; attempt the COM rename even if only casing changes.

## User Scenarios & Testing *(mandatory)*

<!--
  IMPORTANT: User stories should be PRIORITIZED as user journeys ordered by importance.
  Each user story/journey must be INDEPENDENTLY TESTABLE - meaning if you implement just ONE of them,
  you should still have a viable MVP (Minimum Viable Product) that delivers value.
  
  Assign priorities (P1, P2, P3, etc.) to each story, where P1 is the most critical.
  Think of each story as a standalone slice of functionality that can be:
  - Developed independently
  - Tested independently
  - Deployed independently
  - Demonstrated to users independently
-->

### User Story 1 - Rename an existing Power Query (Priority: P1)

As a user automating workbook setup and maintenance, I want to rename an existing Power Query so that the query name matches my preferred naming convention without needing to recreate or re-import the query.

**Why this priority**: Renaming is a common, low-risk maintenance task that removes manual UI work and enables consistent naming across workbooks and teams.

**Independent Test**: Can be fully tested by creating/importing a query, renaming it, and verifying the new name appears in the workbook while the query content remains unchanged.

**Acceptance Scenarios**:

1. **Given** a workbook with a Power Query named "OldQuery", **When** I request a rename to "NewQuery", **Then** the workbook lists "NewQuery" and no longer lists "OldQuery".
2. **Given** a workbook with a Power Query named "OldQuery" with defined query content, **When** I rename it to "NewQuery", **Then** the query content remains the same after the rename.
3. **Given** a workbook that already contains a Power Query named "NewQuery", **When** I rename "OldQuery" to "NewQuery", **Then** the rename fails with a clear message explaining the name conflict.

---

### User Story 2 - Rename an existing Data Model table (Priority: P2)

As a user maintaining an analytical workbook, I want to rename an existing table in the workbook’s data model so that downstream reporting (field lists, visuals, documentation) uses consistent and understandable table names.

**Why this priority**: Data model naming conventions significantly impact usability, but renaming is currently a frequent manual step when iterating on model design.

**Independent Test**: Can be fully tested by creating a workbook with a data model table, renaming the table, and verifying the table list reflects the new name while the workbook remains usable.

**Acceptance Scenarios**:

1. **Given** a workbook with a data model table named "OldTable", **When** I request a rename to "NewTable", **Then** the data model lists "NewTable" and no longer lists "OldTable".
2. **Given** a workbook where renaming "OldTable" would create a duplicate table name, **When** I request the rename, **Then** the rename fails with a clear message explaining the name conflict.
3. **Given** a workbook where "OldTable" has dependent objects (for example relationships or measures), **When** I request the rename, **Then** the system attempts the rename via Excel COM and (a) succeeds if Excel preserves dependencies, or (b) fails if Excel rejects the rename (with the Excel error message).

---

### User Story 3 - Rename safely within an automated workflow (Priority: P3)

As a user running multi-step workbook automation, I want rename operations to behave predictably and report meaningful results so that my automation can decide what to do next (continue, retry, or stop).

**Why this priority**: Rename operations are often part of a larger workflow (import → load → rename → refresh → validate). Reliable outcomes and clear errors prevent fragile automation.

**Independent Test**: Can be fully tested by running a rename operation against (a) a valid object, (b) a missing object, and (c) a conflicting name, and validating that each outcome is deterministic.

**Acceptance Scenarios**:

1. **Given** a workbook where the target query/table does not exist, **When** I request a rename, **Then** the operation fails with a clear message stating the object was not found.
2. **Given** a workbook where the new name is invalid, **When** I request a rename, **Then** the operation fails with a clear message stating the naming rule violation.
3. **Given** a rename request where the new name is exactly the existing name (after trimming), **When** I request the rename, **Then** the operation completes without changing the workbook state.
4. **Given** a rename request where the new name differs only by casing (for example "Sales" → "sales"), **When** I request the rename, **Then** the system attempts the rename via COM and succeeds or fails based on Excel behavior.

---

[Add more user stories as needed, each with an assigned priority]

### Edge Cases

- Rename target does not exist (Power Query or data model table).
- New name already exists (name collision; compare case-insensitively and after trimming whitespace).
- New name contains invalid characters or violates naming rules.
- Rename requested while workbook is read-only or protected.
- Rename requested for an object that is currently refreshing or otherwise busy.
- Rename requested for a data model table that has dependencies that cannot be safely preserved.
- Rename requested where the new name differs only by casing (case-only rename).

## Requirements *(mandatory)*

<!--
  ACTION REQUIRED: The content in this section represents placeholders.
  Fill them out with the right functional requirements.
-->

### Functional Requirements

- **FR-001**: System MUST allow renaming an existing Power Query by specifying the current name and the desired new name.
- **FR-002**: System MUST allow renaming an existing data model table by specifying the current name and the desired new name.
- **FR-003**: System MUST prevent renames that would result in duplicate names within the same scope (queries vs model tables), using case-insensitive comparison after trimming leading/trailing whitespace. The uniqueness check MUST exclude the target object being renamed (so case-only renames of the target are allowed).
- **FR-004**: System MUST validate the requested new name and fail with a clear message when the name is invalid.
- **FR-005**: System MUST fail with a clear message when the target query/table does not exist.
- **FR-006**: Rename operations MUST be safe: they must not leave the workbook in a broken or partially-updated state.
- **FR-007**: Rename operations MUST be usable in automated workflows (no interactive prompts required to complete the operation).
- **FR-008**: The system MUST return a structured result that includes the object type (query or model table), old name, new name, and success/failure.
- **FR-009**: For data model table rename, the system MUST attempt the rename via Excel COM; if Excel rejects/throws, the system MUST fail with that error rather than completing a rename that would break the workbook.
- **FR-010**: The system MUST expose rename operations via (1) MCP tools and (2) CLI commands, both backed by Core commands.
- **FR-011**: Rename operations MUST NOT automatically save the workbook; saving/committing is controlled by the caller.
- **FR-012**: If the requested new name is exactly the existing name after trimming, the operation MUST be a no-op success (no COM rename attempt).

### Assumptions

- Users are operating on a workbook that can be opened for editing.
- Queries and model tables have unique names within their respective collections.
- The workbook contains a supported data model feature set (if a workbook has no data model, rename requests fail clearly).

### Non-Goals

- Editing query content, steps, or load settings as part of rename.
- Renaming worksheet tables (Excel tables) in this feature.
- Renaming objects via Excel UI automation (send keys/clicks) instead of COM APIs.
- Automatically repairing or rewriting user-authored formulas or references outside the supported rename scope.

### Key Entities *(include if feature involves data)*

- **Workbook**: The file being automated; contains queries and optionally a data model.
- **Power Query**: A named query inside a workbook; has a name and query content.
- **Data Model Table**: A named table inside the workbook’s data model.
- **Rename Request**: The user-provided old name and new name, plus the target object type.
- **Operation Result**: The outcome returned to the user (success/failure, messages, and the final name).

## Success Criteria *(mandatory)*

<!--
  ACTION REQUIRED: Define measurable success criteria.
  These must be technology-agnostic and measurable.
-->

### Measurable Outcomes

- **SC-001**: Users can rename an existing Power Query in under 10 seconds for a typical workbook.
- **SC-002**: Users can rename an existing data model table in under 15 seconds for a typical workbook.
- **SC-003**: In a set of valid rename requests, at least 95% complete successfully on the first attempt.
- **SC-004**: In failure cases (missing object, invalid name, name conflict), users receive a clear error message that identifies the cause.
- **SC-005**: After a successful rename, saving and reopening the workbook retains the renamed query/table.
