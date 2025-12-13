# Feature Specification: Upgrade MCP SDK to 0.5.0-preview.1

**Feature Branch**: `001-upgrade-mcp-sdk`  
**Created**: 2025-12-09  
**Status**: Draft  
**Input**: User description: "Upgrade to ModelContextProtocol 0.5.0-preview.1, analyze changelog for new/changed features, and plan impact on ExcelMcp."

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

### User Story 1 - Verify SDK upgrade compatibility (Priority: P1)

Engineering maintains ExcelMcp so that it builds and runs on ModelContextProtocol 0.5.0-preview.1 without regressions in existing tools or CLI flows.

**Why this priority**: The project must stay compatible with upstream MCP protocol changes; broken builds block all contributors and release pipelines.

**Independent Test**: Pull branch, bump the dependency to 0.5.0-preview.1, run targeted build and feature-scoped tests (no code changes needed beyond the bump) and confirm no failures.

**Acceptance Scenarios**:

1. **Given** the dependency is upgraded, **When** `dotnet build` runs, **Then** it succeeds with zero warnings or errors.
2. **Given** the dependency is upgraded, **When** feature-scoped test filters run for MCP server and core layers, **Then** all pass without new failures.

---

### User Story 2 - Capture changelog-to-impact mapping (Priority: P2)

As a maintainer, I can see a concise mapping of 0.5.0-preview.1 release notes to affected ExcelMcp components (MCP server tools, prompt files, batch/session flows) so I know what to modify or validate.

**Why this priority**: Without an explicit impact map, important protocol changes (notifications, capabilities, schema adjustments) could be missed, leading to runtime errors or non-compliant responses.

**Independent Test**: Generate an impact report document referencing release note items and listing affected code areas and required actions; reviewers can validate it without executing code.

**Acceptance Scenarios**:

1. **Given** release notes and API diffs are reviewed, **When** the impact report is produced, **Then** each noted change is tied to specific ExcelMcp areas (tools, prompts, transports, tests).
2. **Given** the impact report exists, **When** a reviewer inspects it, **Then** any missing or ambiguous items are called out with [NEEDS CLARIFICATION] markers or assumptions.

---

### User Story 3 - Define validation and rollback plan (Priority: P3)

As release engineering, I have a validation and rollback checklist for the MCP SDK bump so we can cut or revert the upgrade safely if issues surface.

**Why this priority**: A controlled rollback path reduces downtime and risk if the preview package introduces breaking changes.

**Independent Test**: Review the plan to ensure it lists validation steps, decision gates, and rollback commands; this can be approved without executing the upgrade.

**Acceptance Scenarios**:

1. **Given** the validation plan is drafted, **When** a reviewer reads it, **Then** it lists build/test coverage, targeted scenarios (tools, prompts, CLI), and decision criteria for release vs. rollback.
2. **Given** rollback steps are defined, **When** simulated failure scenarios are reviewed, **Then** the steps include dependency revert, branch reset, and communication steps.

---

[Add more user stories as needed, each with an assigned priority]

### Edge Cases

- Protocol version negotiation fails or the client/server advertises older schema; plan must specify fallback or minimum supported version behavior.
- SDK introduces new notification capabilities (e.g., tool list or roots list changes); verify ExcelMcp either declares or omits them explicitly to avoid misleading clients.
- Structured content or multi-content tool results (e.g., multiple content blocks per result) appear in 0.5.0; ensure serialization/deserialization and MCP error responses remain JSON and not exceptions.
- New or obsoleted enums/diagnostics in the SDK produce build warnings; confirm suppression policy or code updates are applied instead of ignoring warnings.
- Preview package introduces experimental APIs; ensure no accidental opt-in without explicit decision.

## Requirements *(mandatory)*

<!--
  ACTION REQUIRED: The content in this section represents placeholders.
  Fill them out with the right functional requirements.
-->

### Functional Requirements

- **FR-001**: Dependency version for ModelContextProtocol packages MUST be bumped to 0.5.0-preview.1 and build without new warnings or errors.
- **FR-002**: A written changelog impact report MUST enumerate new/changed/removed SDK features and map each to affected ExcelMcp components (Core commands, MCP server tools, prompts, CLI, tests).
- **FR-003**: Compatibility assessment MUST identify any SDK API obsoletions/experimental flags and document whether ExcelMcp uses them; remediation steps MUST be listed for each usage.
- **FR-004**: Validation plan MUST define the minimal test matrix (build + targeted `Feature` filters for MCP server/core) required before merging the upgrade.
- **FR-005**: Rollback plan MUST specify how to revert to the previous SDK version and how to gate release if blocking regressions are found.
- **FR-006**: Tool response contract MUST be re-verified for 0.5.0 changes (e.g., structured content, notifications) and deviations MUST be addressed with code or schema updates.
- **FR-007**: Documentation updates (developer guidance and prompts, if impacted) MUST be listed with owners and locations to edit.
- **FR-008**: Release notes sources MUST be archived/linked for future audits; if official notes are missing, alternative sources (NuGet metadata, repo releases) MUST be captured.
- **FR-009**: Decision on adopting new capabilities (e.g., listChanged notifications) MUST be recorded with rationale — choose to keep capabilities unchanged (no listChanged opt-in) unless a future release explicitly requires dynamic notifications.
- **FR-010**: All enum schema usages MUST be migrated from `EnumSchema`/`LegacyTitledEnumSchema` to the new schema types introduced in SDK 0.5.0; suppression of MCP9001 warnings is NOT allowed.
- **FR-011**: Migrate ALL call sites that previously passed individual request parameters (e.g., `JsonSerializerOptions`, progress tokens) to use the unified `RequestOptions` bag across MCP Server, Core, CLI, and tests in a single pass.

#### Derived Changes & New Capabilities (from 0.5.0-preview.1)

- **FR-012 — Factories removal**: Replace any remaining usage of `McpServerFactory`/`McpClientFactory` with `McpServer.CreateAsync` / `McpClient.CreateAsync` and refactor initialization code accordingly.
- **FR-013 — Interface type updates**: Audit for references to `IMcpEndpoint`, `IMcpClient`, `IMcpServer`; migrate to concrete `McpClient`, `McpServer`, and `McpSession` abstractions where applicable.
- **FR-014 — Enumeration API alignment**: Ensure all list operations use the SDK’s `List*Async` (or synchronous list where applicable); remove/deprecate any usage of `Enumerate*Async` helpers.
- **FR-015 — Protocol exception data**: Enhance server-side error handling to optionally include structured `Data` on `McpProtocolException` for protocol errors (parameter validation, preconditions) while continuing to return JSON results for business errors per our MCP server guide.
- **FR-016 — New error code handling**: Recognize and correctly surface `ResourceNotFound` (-32002) in MCP tool responses and CLI, mapping to clear, actionable error messages without throwing for business errors.
- **FR-017 — Method signature updates**: If used, update `SetLoggingLevel` → `SetLoggingLevelAsync` and `UnsubscribeFromResourceAsync` to the `UnsubscribeRequestParams` signature.
- **FR-018 — Cancellation token arg rename**: Remove named-argument references to `token`; align to `cancellationToken` to avoid compile breaks.
 - **FR-019 — URL mode elicitation (deferred)**: Not required for this upgrade. Defer URL-mode elicitation changes to a future release unless a concrete tool requires URL inputs and demonstrates benefit.
 - **FR-020 — Client tool metadata**: Where we expose client tools, adopt `WithMeta` to attach meaningful metadata (e.g., feature tags, version info) to improve discoverability for consuming clients.
- **FR-021 — Schema migration completion**: Migrate all enum and related schema declarations in MCP Server tool definitions to the new SDK schema types; verify descriptions match behavior per our MCP Server Guide.
 - **FR-022 — New MCP attributes adoption**: Evaluate and adopt any new or expanded MCP SDK attributes for server registration and schema enrichment (e.g., enhanced `[McpServerTool]`, `[McpServerPrompt]`, or attribute metadata fields). Ensure attribute usage centralizes tool description guidance (Rule 18) and keeps schemas aligned with server behavior.

#### .NET Console Application Best Practices (MCP Server)

- **FR-023 — Stdout protocol purity**: MCP Server MUST not write non-protocol output to stdout. All diagnostics, logs, and human-readable messages MUST go to stderr to avoid corrupting MCP transports.
- **FR-024 — Deterministic exit codes**: MCP Server MUST return deterministic process exit codes: `0` for normal shutdown, and `1` for any fatal startup/runtime failure (without relying on unhandled exception process termination).
- **FR-025 — Graceful shutdown with time budget**: MCP Server MUST shut down gracefully on cancellation (Ctrl+C / termination) within 5 seconds, ensuring in-flight work is stopped safely and telemetry/log buffers are flushed without hanging.
- **FR-026 — Configuration-first behavior**: MCP Server MUST be configurable via standard console configuration sources (environment variables, config files where applicable) for log level, telemetry enablement, and other runtime options; it MUST not require interactive prompts.
- **FR-027 — Structured logging & diagnostics**: MCP Server MUST expose structured logging suitable for console/daemon use (timestamped levels, category names), with a clear way to increase verbosity for troubleshooting.
- **FR-028 — Startup validation**: MCP Server MUST validate critical startup prerequisites (e.g., required environment/config presence where mandatory) and fail fast with a clear error message on stderr and a non-zero exit code.

### Key Entities *(include if feature involves data)*

- **Dependency Set**: ModelContextProtocol package versions and transitive dependencies tracked in Directory.Packages.props.
- **Impact Report**: Mapping of SDK changes to ExcelMcp areas (Core, MCP Server tools/prompts, CLI, tests) plus remediation actions.
- **Validation Matrix**: Required build and test filters to declare the upgrade safe (feature-scoped integration tests only, per critical rules).
- **Rollback Plan**: Steps to revert package versions and branch changes if regressions are detected.

## Assumptions & Changelog Analysis

### SDK 0.5.0-preview.1 Changelog Findings

| Change | PR | Impact | Action |
|--------|----|----|--------|
| `RequestOptions` bag replaces individual params on high-level requests | #970 | Call sites passing `JsonSerializerOptions` or `ProgressToken` separately break | Audit call sites; migrate to `RequestOptions` |
| Removed `McpServerFactory` / `McpClientFactory` | #985 | Factory classes deleted | Use `McpClient.CreateAsync` / `McpServer.CreateAsync` |
| Removed `IMcpEndpoint`, `IMcpClient`, `IMcpServer` interfaces | #985 | Interface types deleted | Use `McpClient`, `McpServer`, `McpSession` abstract classes |
| Removed `Enumerate*Async` methods | #1060 | Enumeration helpers gone | Use `List*Async` (likely no change needed if already using List) |
| `McpProtocolException.Data` property | #1028 | Protocol exceptions can carry extra data | Optional: enrich error responses |
| `ResourceNotFound` error code (-32002) | #1062 | New error code for missing resources | Validate error-handling paths |
| `SetLoggingLevel` → `SetLoggingLevelAsync` | #1063 | Method renamed | Adjust if used |
| `UnsubscribeFromResourceAsync` signature change | #1063 | Uses `UnsubscribeRequestParams` | Adjust if used |
| Argument rename: `token` → `cancellationToken` | #1063 | Named arguments break | Search for named arg usages |
| URL mode elicitation | #1021 | New elicitation flow | Optional adoption for future prompts |
| `WithMeta` for `McpClientTool` | #1027 | Attach metadata to client tools | Optional enhancement |
| `EnumSchema` / `LegacyTitledEnumSchema` obsolete (MCP9001) | #985 | Produces build warning | Suppress or migrate to new schema types |

### Assumptions

- ExcelMcp does not currently use `McpServerFactory`/`McpClientFactory` (uses attribute-based registration).
- ExcelMcp does not call `Enumerate*Async` methods.
- ExcelMcp does not directly call `SetLoggingLevel` or `UnsubscribeFromResourceAsync`.
- Named arguments for `CancellationToken token` are not used in ExcelMcp call sites.
- These assumptions must be verified via codebase search before marking upgrade complete.

## Success Criteria *(mandatory)*

<!--
  ACTION REQUIRED: Define measurable success criteria.
  These must be technology-agnostic and measurable.
-->

### Measurable Outcomes

- **SC-001**: `dotnet build` succeeds with ModelContextProtocol 0.5.0-preview.1 and zero warnings/errors across all projects.
- **SC-002**: Targeted integration tests for MCP server and core layers (feature filters) complete with zero new failures after the upgrade.
- **SC-003**: Changelog impact report is completed and reviewed, covering 100% of identified SDK changes with mapped actions or explicit “no action” rationale.
- **SC-004**: Validation and rollback plan is documented and approved, with clear go/no-go criteria and revert steps validated by reviewer sign-off.
- **SC-005**: No MCP9001 (obsolete enum schema) warnings remain; all enum schema references use the new types.
- **SC-006**: No usage of deprecated standalone parameters remains at compile time; all affected calls use `RequestOptions` across MCP Server, Core, CLI, and tests.
- **SC-007**: No references remain to removed factories/interfaces (`McpServerFactory`, `McpClientFactory`, `IMcp*` types); builds and tests pass after migration.
- **SC-008**: All list operations rely on supported list APIs; no `Enumerate*Async` usages present.
- **SC-009**: Protocol error handling demonstrates inclusion of structured `Data` when relevant, and business errors continue to return JSON with `success: false` (no thrown exceptions).
- **SC-010**: `ResourceNotFound` (-32002) is surfaced consistently in MCP responses/CLI output with clear context.
- **SC-011**: Any touched signatures (e.g., `SetLoggingLevelAsync`, `UnsubscribeRequestParams`, `cancellationToken`) compile cleanly across all layers.
 - **SC-012**: `WithMeta` usage is documented and verified end-to-end in at least one tool/prompt. URL-mode elicitation is out of scope for this upgrade.
 - **SC-013**: At least one tool/prompt demonstrates attribute-driven metadata and schema enrichment in the generated tool schema, and descriptions remain accurate per the MCP Server Guide.
- **SC-014**: No non-protocol writes to stdout are observed during MCP Server startup and runtime (validated by test harness or reviewable automation).
- **SC-015**: MCP Server returns exit code `0` on normal shutdown and a non-zero exit code on fatal errors (validated via automated invocation scenarios).
- **SC-015a**: Fatal startup/runtime failures return exit code `1`.
- **SC-016**: Cancellation-triggered shutdown completes within a defined time budget (e.g., under 5 seconds) while preserving protocol integrity and flushing telemetry/logs.
- **SC-017**: Logging verbosity can be raised via configuration without code changes and without introducing stdout noise.

## Clarifications

### Session 2025-12-13

- Q: What is our policy for obsolete enum schema types (MCP9001)? → A: Migrate now to new enum schema types.
- Q: What is our migration scope for the new `RequestOptions` bag replacing individual parameters? → A: Migrate all call sites across MCP Server, Core, CLI, and tests in one pass.
- Q: What exit code should MCP Server use for fatal errors? → A: Use exit code `1` for all fatal errors.

### Applied Decisions

- Functional Requirements: Add explicit requirement to migrate all enum schema usages away from `EnumSchema`/`LegacyTitledEnumSchema` to the new schema types during the SDK upgrade (no warning suppression).
- Success Criteria: Add measurable outcome ensuring no MCP9001 (obsolete enum schema) warnings remain after migration.
- Functional Requirements: Add requirement to adopt `RequestOptions` universally across MCP Server, Core, CLI, and tests.
- Success Criteria: Add measurable outcome confirming no deprecated standalone parameter usage remains and `RequestOptions` is applied across all layers.

