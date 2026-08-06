# ExcelMcp landscape — 2026-08-05

## Scope and reading guide

This is a beginner-friendly snapshot of `sbroenne/mcp-server-excel` as examined on 2026-08-05. It is based on the checked-out `main` revision `60d25c08caefd3726b8e24419b1569ed99e41731` (2026-07-27), the project's first-party documentation and source, and the repository's live GitHub issues and pull requests. “Verified” below means confirmed by maintainers/source evidence, not merely that a user reported it.

## Executive summary

ExcelMcp is a Windows desktop automation bridge: an AI assistant asks for an Excel action, and the bridge drives the **installed Excel application** through Microsoft's COM interface. It is not an online spreadsheet service and not a general-purpose `.xlsx` parser. That gives it unusually broad fidelity—Power Query, DAX, PivotTables, charts, VBA, formulas, formatting, and live Excel calculation—but also makes a local Windows machine with Excel a hard requirement and makes version-, locale-, policy-, and workbook-specific behavior important. [Project overview and requirements](https://github.com/sbroenne/mcp-server-excel/blob/60d25c08caefd3726b8e24419b1569ed99e41731/README.md#L12-L28)

For a non-technical user, the useful mental model is: **your AI is the planner; ExcelMcp is the local pair of hands; Excel remains the source of truth.** Tell the AI what you want, have it show Excel for review when the change matters, and save/inspect the workbook normally. Do not treat it as a safe unattended batch system for sensitive or irreplaceable files.

## What a user can do today

The public feature reference groups 234 operations across 26 tools, including file/session management; worksheets, ranges and formatting; Excel tables; Power Query and connections; Data Model/DAX; PivotTables and slicers; charts; VBA; named ranges; calculation; conditional formatting; screenshots; and window management. The tool can render worksheet/range screenshots for an AI to visually check its work. [Feature inventory](https://github.com/sbroenne/mcp-server-excel/blob/60d25c08caefd3726b8e24419b1569ed99e41731/README.md#L31-L64) [Detailed feature reference](https://github.com/sbroenne/mcp-server-excel/blob/60d25c08caefd3726b8e24419b1569ed99e41731/FEATURES.md)

One recently delivered example of its advanced coverage is visual conditional formatting: version 1.10.2 added creation plus detailed inspection/round-trip output for color scales, data bars, icon sets, top/bottom, above/below, time-period, unique-value and blank-cell rules. [v1.10.2 changelog](https://github.com/sbroenne/mcp-server-excel/blob/60d25c08caefd3726b8e24419b1569ed99e41731/CHANGELOG.md#L14-L20)

There are two ways in:

| Route | Best fit | What happens |
| --- | --- | --- |
| MCP server | Conversational assistants such as Claude Desktop or VS Code chat | The assistant discovers purpose-built Excel tools and calls them during a conversation. |
| `excelcli` command line | Coding agents and scripts | Commands go to a background local service so a workbook can stay open across commands. |

The project presents these as equal entry points with shared behavior. The CLI is more token-efficient for coding agents; the MCP server is more natural for exploratory chat. They are separate processes and **do not share a live Excel session** with one another. [Route comparison](https://github.com/sbroenne/mcp-server-excel/blob/60d25c08caefd3726b8e24419b1569ed99e41731/README.md#L112-L146) [Session separation](https://github.com/sbroenne/mcp-server-excel/blob/60d25c08caefd3726b8e24419b1569ed99e41731/README.md#L190-L237)

### Practical boundaries

- It requires Windows, desktop Excel 2016 or later, and a desktop environment. It is explicitly not designed for Linux/macOS, server-side Excel processing, or high-volume batch work. [Requirements and non-goals](https://github.com/sbroenne/mcp-server-excel/blob/60d25c08caefd3726b8e24419b1569ed99e41731/README.md#L24-L28) [Suitability guidance](https://github.com/sbroenne/mcp-server-excel/blob/60d25c08caefd3726b8e24419b1569ed99e41731/README.md#L98-L108)
- ExcelMcp asks users to close Excel files before automation because it needs exclusive workbook access. [Quick-start note](https://github.com/sbroenne/mcp-server-excel/blob/60d25c08caefd3726b8e24419b1569ed99e41731/README.md#L110-L111)
- Macro/VBA access additionally requires the user to enable Excel's “Trust access to the VBA project object model”; the project deliberately does not enable that setting itself. [VBA feature guidance](https://github.com/sbroenne/mcp-server-excel/blob/60d25c08caefd3726b8e24419b1569ed99e41731/FEATURES.md)

## Architecture and data flow

```text
AI assistant / script
        |
        | MCP tool call                 | CLI command
        v                               v
MCP server (in process)          CLI client -> local named-pipe daemon
        \                               /
         \                             /
          v                           v
             ExcelMcp Service (sessions and routing)
                              |
                              v
                Core commands (the Excel feature logic)
                              |
                              v
         COM interop / single-threaded Excel automation plumbing
                              |
                              v
                   Local Microsoft Excel application
                              |
                              v
                     Workbook, formulas, queries, etc.
```

The MCP server invokes the service directly in its own process. The CLI instead talks to a background daemon through a named pipe, preserving CLI sessions across separate commands. Both ultimately use the same Core command layer and Excel's `Excel.Application` COM API. [Architecture overview](https://github.com/sbroenne/mcp-server-excel/blob/60d25c08caefd3726b8e24419b1569ed99e41731/README.md#L190-L237) [Service composition](https://github.com/sbroenne/mcp-server-excel/blob/60d25c08caefd3726b8e24419b1569ed99e41731/src/ExcelMcp.Service/ExcelMcpService.cs#L23-L63)

The main building blocks are:

- **ComInterop**: manages Excel's COM objects, STA threading, batches, and session lifetime—the compatibility-sensitive layer closest to Excel. [Contributing architecture](https://github.com/sbroenne/mcp-server-excel/blob/60d25c08caefd3726b8e24419b1569ed99e41731/docs/CONTRIBUTING.md#L69-L81)
- **Core**: the actual Excel feature operations, shared by MCP and CLI.
- **Service**: session management and routing; its pipe server permits up to ten concurrent connections. [Service pipe handling](https://github.com/sbroenne/mcp-server-excel/blob/60d25c08caefd3726b8e24419b1569ed99e41731/src/ExcelMcp.Service/ExcelMcpService.cs#L103-L141)
- **Source generators**: derive MCP tools and CLI commands from Core interfaces. This is the key guard against the two routes drifting apart, although it does make generated-surface testing important. [Generator implementation](https://github.com/sbroenne/mcp-server-excel/blob/60d25c08caefd3726b8e24419b1569ed99e41731/src/ExcelMcp.Generators.Mcp/McpToolGenerator.cs#L8-L36)
- **Delivery surfaces**: standalone executables/ZIPs, NuGet tools, VS Code extension, MCPB package, skills and plugins. [Release strategy](https://github.com/sbroenne/mcp-server-excel/blob/60d25c08caefd3726b8e24419b1569ed99e41731/docs/RELEASE-STRATEGY.md#L7-L37)

## Operational, security, and privacy picture

### What the security model does—and does not—mean

ExcelMcp runs as the signed-in Windows user and can automate the files and Excel instances that user can access. It does not ask for administrator rights. That is convenient, but a prompt can make consequential local changes: overwrite values, refresh external connections, run a VBA procedure, or save a workbook. Review AI-proposed actions, work on copies for high-stakes files, and only connect trusted AI clients. [Current-user model and user guidance](https://github.com/sbroenne/mcp-server-excel/blob/60d25c08caefd3726b8e24419b1569ed99e41731/SECURITY.md#L67-L101)

For CLI, the named pipe includes the Windows user SID and uses an ACL for that same user; it is local-only and cannot be accessed by a different Windows user or over the network. However, **another process running as the same Windows user can connect**. The project documents that limitation as intentional and comparable to other local developer daemons. [Named-pipe security model](https://github.com/sbroenne/mcp-server-excel/blob/60d25c08caefd3726b8e24419b1569ed99e41731/SECURITY.md#L37-L64) [ACL implementation](https://github.com/sbroenne/mcp-server-excel/blob/60d25c08caefd3726b8e24419b1569ed99e41731/src/ExcelMcp.Service/ServiceSecurity.cs#L7-L81)

The policy claims full-path normalization, permitted workbook extensions, file-size and path-length limits, hidden Excel windows, resource cleanup, analyzers, and CodeQL. [Security features](https://github.com/sbroenne/mcp-server-excel/blob/60d25c08caefd3726b8e24419b1569ed99e41731/SECURITY.md#L18-L35) The checked source clearly normalizes paths and checks `.xlsx`/`.xlsm` when opening files, but this review did **not** find the stated 1 GB file-size limit or a general 32,767-character path limit in the production code; new-file creation uses a much lower practical path check. Treat that as a documentation/implementation discrepancy to resolve rather than relying on the policy claim. [Open-session validation](https://github.com/sbroenne/mcp-server-excel/blob/60d25c08caefd3726b8e24419b1569ed99e41731/src/ExcelMcp.ComInterop/Session/ExcelSession.cs#L73-L95) [New-file path check](https://github.com/sbroenne/mcp-server-excel/blob/60d25c08caefd3726b8e24419b1569ed99e41731/src/ExcelMcp.ComInterop/Session/ExcelSession.cs#L129-L137)

### Privacy

According to the privacy policy, all workbook work stays local. The MCP server (not the CLI/daemon) sends limited anonymous usage telemetry to Azure Application Insights: tool/action, duration, outcome, a session ID, hashed machine identifier, version, and exception type; it says it does not send workbook contents, names, paths, or personal information. [Privacy policy](https://github.com/sbroenne/mcp-server-excel/blob/60d25c08caefd3726b8e24419b1569ed99e41731/PRIVACY.md#L9-L66)

The code includes a redactor for paths, secrets and emails. [Redactor](https://github.com/sbroenne/mcp-server-excel/blob/60d25c08caefd3726b8e24419b1569ed99e41731/src/ExcelMcp.McpServer/Telemetry/SensitiveDataRedactor.cs#L8-L84) One item merits a security/privacy review: the unhandled-exception path obtains redacted values, but then creates Application Insights `ExceptionTelemetry` from the original exception object. That code path may preserve raw message/stack content depending on the SDK; it should be verified against the emitted telemetry and either sanitized before construction or documented accurately. This is a **source-evidenced risk to investigate**, not a confirmed data exposure. [Unhandled-exception path](https://github.com/sbroenne/mcp-server-excel/blob/60d25c08caefd3726b8e24419b1569ed99e41731/src/ExcelMcp.McpServer/Telemetry/ExcelMcpTelemetry.cs#L181-L194)

## Quality, testing, and release constraints

The project's central testing decision is to favour integration tests that launch real Excel over mocked COM unit tests. That is a sensible response to COM-specific failure modes such as threading, cleanup, conversion, and refresh behaviour; it also means tests require Windows and Excel and take longer. The project describes roughly 200+ integration tests, several on-demand/diagnostic categories, and separate manual LLM tests. [Test rationale and categories](https://github.com/sbroenne/mcp-server-excel/blob/60d25c08caefd3726b8e24419b1569ed99e41731/tests/README.md#L1-L75) [ADR](https://github.com/sbroenne/mcp-server-excel/blob/60d25c08caefd3726b8e24419b1569ed99e41731/docs/ADR-001-NO-UNIT-TESTS.md#L80-L107)

There is a strong MCP smoke test that exercises the SDK/protocol, dependency injection, tool discovery, sessions and telemetry with real Excel. [MCP smoke test](https://github.com/sbroenne/mcp-server-excel/blob/60d25c08caefd3726b8e24419b1569ed99e41731/tests/ExcelMcp.McpServer.Tests/Integration/Tools/McpServerSmokeTests.cs#L15-L145) The key release risk is coverage availability: standard CI deliberately runs build and Excel-free checks only; the full integration suite is on a self-hosted Windows/Excel runner and has a 300-minute allowance. Some environment/version combinations therefore cannot be exercised on every public PR. [Standard CI boundary](https://github.com/sbroenne/mcp-server-excel/blob/60d25c08caefd3726b8e24419b1569ed99e41731/.github/workflows/ci.yml#L97-L123) [Dedicated integration workflow](https://github.com/sbroenne/mcp-server-excel/blob/60d25c08caefd3726b8e24419b1569ed99e41731/.github/workflows/integration-tests.yml#L71-L86)

Releases are unified: the automation builds and publishes the CLI and MCP server, packages/registries/extensions/skills, creates a GitHub release, and uses changeset fragments to produce the changelog. This broad delivery matrix is valuable for users but increases release verification burden. [Release workflow](https://github.com/sbroenne/mcp-server-excel/blob/60d25c08caefd3726b8e24419b1569ed99e41731/.github/workflows/release.yml#L707-L877) [Changeset policy](https://github.com/sbroenne/mcp-server-excel/blob/60d25c08caefd3726b8e24419b1569ed99e41731/docs/RELEASE-STRATEGY.md#L51-L102)

## Live tracker: bugs, requests, and active work

Checked against the live GitHub tracker on 2026-08-05. There were three open issues and four open PRs. One open issue is stale as a tracking record: its requested work was already merged and released.

| Status | Item | Why it is classified this way | User impact / next move |
| --- | --- | --- | --- |
| **Verified bug, active fix** | [#750: formulas fail with `0x800A03EC` on Excel 2019 de-DE 32-bit](https://github.com/sbroenne/mcp-server-excel/issues/750) | A detailed reproduction reports both formula reads and writes failing. The maintainer requested changes on open [PR #751](https://github.com/sbroenne/mcp-server-excel/pull/751), explicitly confirming the compatibility root cause and requiring a capability probe, documentation, a changeset, and target-environment validation. | Formula-heavy workbooks may be unusable on older/perpetual Excel. Finish and validate the capability-based fallback; preserve modern dynamic-array semantics where supported. |
| **Confirmed product gap / error-message improvement; not independently reproduced here** | [#753: Python in Excel reports raw `#NAME?`](https://github.com/sbroenne/mcp-server-excel/issues/753) | It was opened by the maintainer and the source handles transient `#BUSY!`, `#CONNECT!`, and `#BLOCKED!` states but has no matching `#NAME?` availability diagnosis. [Current markers](https://github.com/sbroenne/mcp-server-excel/blob/60d25c08caefd3726b8e24419b1569ed99e41731/src/ExcelMcp.Core/Commands/PythonInExcel/PythonInExcelCommands.cs#L14-L17) | Give an actionable “Python in Excel unavailable” response while retaining retry behaviour for transient cloud states. Positive-path testing needs a licensed/entitled Microsoft 365 environment. |
| **Stale/open tracking issue — inference from release evidence** | [#743: inspect detailed visual conditional-format settings](https://github.com/sbroenne/mcp-server-excel/issues/743) | The issue remains open, but its requested creation and type-specific list/round-trip support shipped in merged [PR #745](https://github.com/sbroenne/mcp-server-excel/pull/745) and is recorded in v1.10.2. This classification is an inference from those primary records, not an explicit “fixed” comment on the issue. | Close #743 with a release/PR reference, unless a remaining acceptance criterion is identified. |
| **Active PR, not ready to merge** | [#751: `Formula2` to `Formula` compatibility fallback](https://github.com/sbroenne/mcp-server-excel/pull/751) | Open, mergeable and non-draft, but the maintainer left a changes-requested review: cache capability per session, probe by read rather than swallowing write errors, add release/docs work, and validate on the reporter's Excel. | This is the immediate delivery risk. Do not describe #750 as fixed until this PR is revised, verified and merged. |
| **Active security-dependency PR** | [#752: undici 7.28.0 → 7.29.0](https://github.com/sbroenne/mcp-server-excel/pull/752) | Dependabot proposes an update in the VS Code extension; its PR cites upstream security fixes. | Review compatibility and merge promptly if checks pass; it affects extension tooling, not Excel COM behaviour. |
| **Active security-dependency PR** | [#754: cryptography 48.0.1 → 50.0.0](https://github.com/sbroenne/mcp-server-excel/pull/754) | Dependabot proposes a major-version update in the LLM-test environment and cites CVE-2026-69247. | Prioritize security review but test the LLM-test environment because a major dependency update can have compatibility effects. |
| **Active security-dependency PR** | [#755: fast-uri 3.1.4 → 3.1.5](https://github.com/sbroenne/mcp-server-excel/pull/755) | Dependabot proposes a VS Code extension development-dependency update whose PR cites an upstream security advisory. | Low implementation scope, but review and merge after the extension checks pass. |

## Gaps and opportunities

1. **Compatibility is a product surface, not an edge case.** Real Excel versions, bitness, locale, licences and tenant settings vary more than a normal server runtime. Build an explicit capability model and show it to users, rather than discovering limitations mid-workflow. #750 and #753 are concrete examples.
2. **Safety controls should match the power of the tool.** The current model gives the AI the user's effective Excel/file capability. Dry runs, an allowlisted workspace folder, an operation preview, and a human-readable change log would make the tool feel safer without weakening its Excel fidelity.
3. **Privacy assurance needs executable evidence.** The telemetry implementation should be tested/inspected against its written promises, particularly for exception messages and stacks. The file-size/path-policy mismatch should likewise be fixed in either code or documentation.
4. **Quality depends on scarce environments.** Real-Excel integration is the right signal, but the self-hosted dependency makes a compatibility matrix and reproducible test images especially important.
5. **Round-trip inspection should remain the standard for advanced objects.** Conditional formatting is a positive recent example—#743's requested support is released even though the issue is still open. More generally, “create, read back, compare, and safely update” is more trustworthy than create-only automation.

## Prioritized feature brainstorm

| Priority | Idea | Why it earns the priority | Main risk / guardrail |
| --- | --- | --- | --- |
| P0 | **Excel capability check and preflight report**: version, bitness, dynamic arrays/`Formula2`, VBA trust, Python-in-Excel availability, workbook protection and known constraints | Prevents avoidable failures and gives beginners a plain-English “what will work here?” answer. Directly addresses #750/#753. | Capabilities can be account/policy/network dependent; distinguish “supported,” “unavailable,” and “not yet determined,” and cache only stable probes. |
| P0 | **Finish formula compatibility safely** | It fixes a reported blocker for the advertised Excel 2016+ floor. | A fallback must not mask write errors or silently change dynamic-array/table semantics; validate on the reported Excel 2019 locale/bitness. |
| P1 | **Review-before-write mode**: planned operations, affected workbook/sheet/range, save destination, and confirmation for destructive actions | Turns powerful automation into a comprehensible assistant for new users and limits prompt-driven accidents. | The plan can diverge from Excel's actual effects; label it a preview and capture actual post-operation results. |
| P1 | **Change journal and reversible checkpoints**: record operations; optionally save a timestamped copy before writes | Makes AI work auditable and recoverable for business spreadsheets. | Copying large/protected workbooks is slow and may alter external-link/macros behaviour; make it opt-in and report failures clearly. |
| P2 | **Close the loop on released features**: automatically link/close shipped issues and surface the installed version's capability set | Prevents tracker confusion like #743 and makes it clear to users whether a feature is already available in their installed build. | Release metadata must be authoritative; avoid auto-closing issues that retain unresolved acceptance criteria. |
| P2 | **Telemetry privacy hardening + a local telemetry viewer/opt-out setting** | Aligns implementation, privacy statement and user expectations; improves trust for sensitive-workbook users. | Do not promise anonymity without end-to-end tests of what the SDK emits; keep the CLI's no-telemetry contract intact. |
| P2 | **Compatibility test matrix and contributor diagnostics bundle** | Turns a difficult support problem into structured evidence: Excel version, bitness, locale, capability report and sanitized operation trace. | Diagnostic data can be sensitive; make collection local-by-default and redact/share only with explicit consent. |
| P3 | **Guided “build and verify” workflows** for common tasks (import CSV → table → PivotTable → chart → screenshot review) | Helps beginners succeed without learning 26 tools or the current 239 operations. | Avoid concealing too much: each step should expose its effect and let users stop or edit. |

## Recommended near-term plan

1. Land #751 only after its requested capability probe, target-machine validation, documentation and release checks are complete.
2. Implement #753 as a precise, tested availability message rather than treating every non-result as a retry.
3. Review and merge the three active security-dependency updates (#752, #754 and #755), with focused extension/LLM-test validation as appropriate.
4. Reconcile `SECURITY.md` with actual enforcement, audit exception telemetry end to end, and close/reconcile the stale #743 tracker item.
5. Add a beginner-facing preflight/review experience before expanding into still more write operations; use conditional formatting as the model for future round-trip features.

### Reconciliation note (current working tree)

The historical observations above predate the safety automation work. The current shared validator enforces a 1 GiB existing-workbook limit and a 32,767-character general path limit, with a practical 218-character Excel SaveAs creation limit. `test` accepts existing `.xlsx`/`.xlsm`; `open` accepts `.xlsx`, `.xlsm`, and legacy `.xls`; `create` writes `.xlsx`/`.xlsm`. Exception telemetry is now constructed from sanitized details, `EXCELMCP_TELEMETRY_OPTOUT=true` disables MCP telemetry initialization, and the CLI remains telemetry-free. Safety is opt-in; enabling a safety control defaults omitted shutdown policy to discard-with-evidence. Checkpoints are local unencrypted full-workbook copies, SHA-256 is not malicious-tamper proof, and recovery cannot restore unsaved memory. Lifecycle file/session mutations remain outside the universal generated-command review handshake, while atomic cross-file worksheet mutations fail closed when safety flags are supplied.
