---
name: "plugin-build-smoke-regression"
description: "Catch packaging-script regressions by asserting the real script exit path and current overlay surface."
domain: "build-tests"
confidence: "high"
source: "observed"
---

## Context
Use this when a packaging or publish script has a real smoke-path failure that synthetic file assertions missed.

## Patterns
- Add at least one test that runs the real script entry point and asserts **exit code 0** plus the final human summary/output block.
- Keep synthetic template fixtures aligned with the **current shipped overlay names**, not historical file names.
- Assert current assets are present **and** legacy bootstrap files are absent, so stale template content cannot make tests pass accidentally.
- Prefer ASCII summary assertions for PowerShell smoke tests when prior failures involved parser/host issues around fancy status output.

## Examples
- `tests/ExcelMcp.SkillGeneration.Tests/PluginBootstrapBuildTests.cs`: `BuildPlugins_SmokeRun_ExitsZeroAndPrintsAsciiSummary`
- `tests/ExcelMcp.SkillGeneration.Tests/PluginBootstrapBuildTests.cs`: current bootstrap assets use `bin\download.ps1` and reject legacy `download-mcp.ps1` / `download-cli.ps1` / packaged `bootstrap-state.json`

## Anti-Patterns
- Do not validate only copied files when the real failure happened in script shutdown/status output.
- Do not let synthetic fixtures carry obsolete helper names that the shipped overlay no longer uses.
