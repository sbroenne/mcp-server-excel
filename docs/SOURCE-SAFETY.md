# Build-time source safety

Ordinary production project and solution builds enforce these policies in
Debug and Release. No PowerShell prerequisite or second scanner exists.
`Directory.Build.targets` attaches the internal, nonshipped `netstandard2.0`
compiler analyzer and Microsoft's `Microsoft.CodeAnalysis.BannedApiAnalyzers`.
The existing SDK/NetAnalyzers continue to enforce general .NET rules.
Generators, the analyzer itself and test projects do not receive production
policy analyzers. The analyzer requires only Roslyn, not Office/VBE runtime
assemblies; Excel semantic checks use the compilation's PIA symbols.

## Diagnostics and deliberate scope

| ID | Compiler/build evidence | Guidance |
| --- | --- | --- |
| `EXCEL001` | Executable `dynamic` type syntax, including nullable types, casts, parameters, returns and generic arguments in completed areas | Use the supported Excel PIA. Document each verified remaining gap immediately above the containing declaration/statement with `// PIA gap:` and member/runtime evidence. |
| `EXCEL002` | Explicit numeric/enum conversion whose operand is dynamic in a completed area | Use `Convert.*`; a PIA-gap comment does not exempt numeric conversion. |
| `EXCEL003` | Chained or discarded reference-type Excel acquisitions, resolved by namespace/type rather than variable names | Capture each acquisition and release in `finally`. `ExcelContext.App`/`Book` and batch `GetWorkbook` are borrowed session references. Managed collection counts are not COM acquisitions. |
| `EXCEL004` | A directly returned/exposed `ResultBase` initializer, or a local object at return after assignments and simple branches (including aliases), provably has `Success=true` and nonempty constant `ErrorMessage` | Clear the error or return failure. Local initializers may be repaired before return; different objects, cleared errors and unknown mutations are not conflated. |
| `EXCEL005` | Evaluated package references, assembly references and resolved transitive dependency identities, including imported items | Production cannot depend on DocumentFormat.OpenXml, System.IO.Packaging or WindowsBase package APIs. |
| `EXCEL006` | Executable workbook-part string literals, corresponding project/import XML values, or XML load/create/parse calls fed by known workbook parts/archives | Access workbook contents with Excel COM. Workbook ZIP streams/text are traced through local initializers, managed readers, helper returns and helper parameters within the compilation. |
| `RS0030` | Microsoft's banned-API analyzer resolves forbidden OpenXML/package namespace symbols | The production banned-symbol list is `config\BannedSymbols.txt`. |

COM checking deliberately covers typed Excel acquisitions. It does not prove
arbitrary interprocedural COM ownership or inspect unavailable external
Office-core/VBE/ADO object models. Reviewed late-bound gaps still require
manual lifetime review. Success checking shares object state across simple
local alias declarations, assignments and reference casts in the enclosing
function block. Simple `if` branches propagate state to branch-local returns
and merge only facts shared by the continuing paths (up to eight nested
branches). Reassigning a local detaches that binding, not its existing aliases;
independent objects stay separate. Conditional alias relationships, calls and
unmodeled uses of tracked references invalidate that object's proof; complex
control flow discards the current proof. This is not complete control-flow or
interprocedural alias analysis.
ZIP provenance requires bounded workbook evidence: a known workbook-part
name (including constant concatenation), or an archive opened from a known
`.xlsx`, `.xlsm`, `.xlsb`, `.xltx` or `.xltm` path. It follows source available
in the compilation, not arbitrary external dependencies or unknown filenames.

Source comments and ordinary string text are not executable dynamic code.
Unrelated ZIP settings XML and GZip XML processing remain valid; mentioning
both compression and XML APIs is not by itself a workbook-content violation.
Test-only fixture OOXML is exempt. Approved binary IRM/AIP metadata inspection
must remain separate from workbook-part parsing.

## Completed-area activation and gaps

Set `excelmcp_typed_excel = true` in `.editorconfig` for completed source
subtrees/files. The analyzer receives the option per syntax tree; it does not
infer completion from filenames. For example:

```ini
[src/ExcelMcp.Core/Commands/Range/*.cs]
excelmcp_typed_excel = true
```

Core and ComInterop are activated in their entirety. The unfinished mode below
is retained for explicitly staged future areas, not as an exemption for these
completed projects.

Unfinished files retain the old documented double-parenthesized dynamic-cast
requirement until activation, but no infrastructure filename is blanket-exempt.
Generic `TODO`/`Reason` comments are accepted only in this unfinished mode.

Completed-mode evidence is restricted to reviewed Office-core/VBE/ADO
dependency gaps and named missing/runtime-mismatched PIA members.
When the cast receiver is a typed Excel object, the analyzer also restricts
the accessed member to the reviewed native surface. An arbitrary comment
about Office-core cannot exempt `Range.Rows` or other supported native members.
The inventory and actual verification evidence live in
[PIA coverage](PIA-COVERAGE.md); adding a gap requires reviewing that evidence,
the narrow compiler policy and both positive/negative tests. Comments record
reviewed evidence; they are not themselves proof of a vendor/runtime limitation.

Style value-type getters and `Font.ThemeFont` are not approved late-bound gaps.
The tested mixed `Font.ThemeFont` native value was integer `0`, not null.
Style and differential-Font probes that fail during setup do not establish a
runtime gap; use the typed PIA getters.

## Validation and compatibility commands

The compiler-testing library is `Microsoft.CodeAnalysis.CSharp.Analyzer.Testing`
1.1.4 with `DefaultVerifier` and existing xUnit. Roslyn Workspaces is aligned
with the repository's Roslyn version; no obsolete framework-specific testing
adapter is used. Tests assert exact diagnostic IDs, source spans and guidance,
and cover legal code to guard against semantic false positives.

```powershell
dotnet test tests\ExcelMcp.ScriptSafety.Tests\ExcelMcp.ScriptSafety.Tests.csproj -c Release --filter "FullyQualifiedName~TypedSourceGuardTests|FullyQualifiedName~WorkbookPackageAccessGuardTests"
```

The isolated build regressions mutate an initially successful SDK project,
assert that every policy fails a real compilation, repair it and prove a
successful incremental build in Debug/Release. They also exercise imported
and evaluated transitive dependencies and test-only OOXML exceptions. Fixtures
and retained results belong under gitignored `artifacts\`, not tracked source.

Hooks and CI rely on their selected SDK builds, retaining native compiler
failure output rather than executing four scanner passes. Source-safety policy
changes select the production compiler consumers plus owning ScriptSafety
regressions. Unknown inputs and empty compatibility project discovery fail.

Existing `scripts\check-com-leaks.ps1`, `check-success-flag.ps1`,
`check-dynamic-casts.ps1` and `check-workbook-package-access.ps1` remain build
adapters for existing callers. `build.ps1 check-source --rule <name>` validates
the requested legacy rule name, then compiles all four policies together.
`com-leaks --inputs-json` maps valid source paths to owning production projects;
it no longer applies a filename/member-name regex to snippets. These commands
are intentionally redundant if a selected production build already passed.
