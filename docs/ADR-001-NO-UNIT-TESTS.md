# ADR-001: Testing Excel behavior with real Excel

**Status:** Superseded by the current [testing strategy](../.github/instructions/testing-strategy.instructions.md)

**Original decision date:** 2025-11-02

## Original decision

The original decision required real Excel integration tests instead of mocked
COM tests. It also stated that the repository had no independently testable
logic and therefore should have no unit tests.

The real-Excel requirement remains valid for COM behavior. The blanket ban on
unit tests no longer describes the project and must not guide new work.

## Current policy

- Test Excel operations, object lifetime, refresh, and workbook persistence
  against desktop Excel. Mocks do not establish that COM behavior works.
- Test parsing, mapping, serialization, generation, and other Excel-independent
  behavior with focused tests that do not require Excel.
- Test the CLI and MCP entry points when a changed contract crosses them.
- Use the smallest relevant test group; session/batch changes also require
  targeted ComInterop OnDemand tests.
- Save and reopen only when persistence is the behavior being tested.

The repository already contains non-COM tests, including
`tests/ExcelMcp.Core.Tests/Unit/ServiceRegistryJsonParsingTests.cs` and
`tests/ExcelMcp.Core.Tests/Unit/PowerQueryIdentityParsingTests.cs`.
These test our own parsing and dispatch behavior, not Excel or .NET itself.

## Rationale

Excel's object model has runtime behavior that cannot be established by a mock:
COM marshaling, application state, refresh completion, and saved workbook
contents require the real application. That does not make a pure parser or
generated argument conversion dependent on Excel. Both kinds of tests are
needed, at the layer that owns the behavior.

The [repository instructions](../.github/copilot-instructions.md#build-and-validation)
define build, local Excel E2E, and CI requirements. GitHub-hosted runners do not
have Excel; Excel-free checks are not a substitute for local COM coverage.

This file retains the historical decision's location so existing links resolve.
Use the current testing strategy for commands and detailed test design rules.
