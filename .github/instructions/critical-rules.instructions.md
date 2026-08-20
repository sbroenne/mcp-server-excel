---
applyTo: "**"
excludeAgent: "code-review"
---

# Non-negotiable repository rules

## Change discipline

- Make the smallest complete change that fixes the root cause. Do not modify
  unrelated code or discard changes already present in the worktree.
- For behavioral changes, use a focused red-green cycle: reproduce the missing
  behavior in a test, observe the failure, implement, and rerun the same test.
  Do not manufacture tests for documentation-only or configuration-only edits.
- Search sibling commands and all generated/parallel surfaces for the same
  pattern before concluding a fix is complete.
- Do not leave placeholders, broad silent fallbacks, commented-out code, or
  unresolved TODO/FIXME/HACK markers in production code.

## Validation

- Build with zero warnings and run the smallest existing tests that exercise the
  changed behavior. Use an explicit timeout for every Excel-dependent test run.
- Core integration tests require real Excel. Never replace COM-dependent
  coverage with mocks. Pure parsing, mapping, and generator logic may use fast
  non-COM tests when that is the behavior under test.
- Verify resulting workbook state, returned fields, and persistence where
  applicable; `Success == true` alone is not sufficient.
- Do not call `batch.Save()` unless the test specifically verifies persistence
  across close/reopen.
- Session and batch infrastructure changes require targeted OnDemand tests.

## Results and exceptions

- `Success == true` requires an empty or null `ErrorMessage`.
- Core command operations run through `IExcelBatch.Execute`. Do not add a broad
  catch that converts exceptions into a second result object. Let the batch and
  service layers preserve the failure context.
- Catch only when handling a specific recoverable condition or adding required
  context at the owning boundary. Never swallow exceptions or return a
  success-shaped fallback.
- MCP tools return structured JSON for tool execution failures. Reserve thrown
  argument/protocol exceptions for invalid tool input or an unknown action.
- MCP stdio stdout is JSON-RPC only. Send diagnostics and logs to stderr.

## Excel COM safety

- Prefer strongly typed `Microsoft.Office.Interop.Excel` APIs. Use `dynamic`
  only for a documented PIA/runtime gap.
- Every acquired COM object must be released in reverse order from a `finally`
  block with `ComUtilities.Release`.
- Excel collections are one-based. Convert COM numeric values with `Convert.*`
  instead of direct casts when Excel can marshal a different numeric type.
- Do not call `RefreshAll()` for operations that require synchronous completion;
  use the established connection or `QueryTable.Refresh(false)` path.
- `ExcelWriteGuard` owns `ScreenUpdating`. Do not suppress it in commands.
  Suppress calculation only in the established bulk value/formula write paths,
  and do not globally suppress events.
- Use `ExcelShutdownService` for workbook close and Excel quit paths. Do not kill
  processes by name.

## Surface parity

- MCP Server and CLI are equal entry points. Changes to actions, parameters,
  defaults, validation, results, timeout behavior, or naming must agree through
  Core, generated Service dispatch, CLI, MCP, tests, and guidance.
- C# interface parameters use camelCase. The MCP generator derives snake_case;
  use the supported name attribute only when normal conversion cannot express
  the required external name.
- Keep enum/action mappings exhaustive and reject unknown values explicitly.
- Update the source template or generator, not only generated output.
- Run the repository coverage, naming, success-flag, COM-leak, dynamic-cast, and
  documentation-count scripts that apply to the changed surface.

## Documentation and release

- Update user documentation only when behavior or workflows changed. There is
  no minimum file or test count; coverage must match the actual risk.
- `skills/shared` is the source of truth for guidance shared by CLI skills and
  generated MCP prompts.
- Use PowerShell examples for Windows workflows. Do not hand-edit generated
  skills, generated prompts, version strings, or `CHANGELOG.md`.
- Add a changeset for user-visible features, fixes, and breaking changes;
  otherwise use the `skip-changelog` PR label.

## Git and privacy

- Never commit to `main`, bypass hooks, force-push, or use a hook-disabling
  environment/configuration override.
- A coding-agent assignment that requests repository changes authorizes its
  delivery branch commits and pull request. Otherwise, ask before commit, push,
  or posting public content. Never merge without a separate explicit request.
- Sanitize customer names, workbook names, local paths, credentials, connection
  strings, and other confidential context from public artifacts.
