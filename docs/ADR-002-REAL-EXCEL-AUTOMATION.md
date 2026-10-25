# ADR-002: Use desktop Excel as the workbook engine

**Status:** Current

## Context and decision

ExcelMcp automates the installed Windows Excel application through COM rather
than implementing workbook operations by reading or rewriting file packages.
Typed Excel interop definitions are embedded at build time and are the
intended way to reach Excel. The build rejects dynamic declarations and
late-bound operations in production and test code. Office Core and VBA editor
definitions are also allowed as build-only inputs, provided the used types are
embedded and deployed applications require no extra interop DLLs. They come from
the pinned `ExcelDna.Interop` package download, not an installed Office directory.
Actual Excel tests and output checks must establish that boundary. This does not
authorize a runtime Office/VBE dependency, a machine-specific reference, or an
assembly resolver.

ADO definitions may also be generated at build time from Windows' signed ADO
2.8 type library using Microsoft's `TlbImp.exe`. The approved build prerequisite
is the .NET Framework 4.8.1 Developer Pack, not the Visual Studio IDE. Normal
`dotnet` builds invoke the tool explicitly, without full-framework MSBuild or
`COMReference`; generated definitions remain in `obj` and used types are
embedded. This must not introduce a runtime ADO definition DLL or replace
Excel's authenticated live connection.

Document properties have one user-approved, isolated exception. Actual Excel
rejects the matching Office `DocumentProperties` and `DocumentProperty`
interfaces, including definitions freshly imported from Microsoft's signed
Office library. `ExcelDocumentProperties` therefore uses Microsoft's documented
native `IDispatch` interface for this feature alone. It resolves member
identifiers from Excel, permits only its fixed collection/property operations,
checks returned item identities against the Office definitions, and retains
native error codes and causes. It owns and releases arguments and results.
No dynamic declaration, reflection fallback, general dispatch API, or extra
runtime interop DLL is allowed. Native regressions must establish reads,
writes, types, enumeration, saved values, errors, and unchanged controls.
See [the native evidence](PIA-COVERAGE.md#document-property-interface-limitation).

Other APIs missing from the referenced definitions are not exceptions;
unsupported capabilities must not bypass the compiler with dynamic access.

The code has not fully reached that state: several command areas still hold
objects the typed definitions support as `dynamic`. That is conversion debt
tracked in [#1140](https://github.com/sbroenne/mcp-server-excel/issues/1140),
not an accepted exception.

The pre-open protection check reads binary container metadata, not workbook
contents. That exception supports opening protected files; it is not another
workbook engine.

## Reasons and tradeoffs

File libraries can work without desktop Excel and are easier to host across
platforms. They do not provide Excel's full calculation, refresh, Data Model,
macro, and authentication behavior. Using Excel also avoids reconstructing
workbook features the automation did not intend to change.

Typed interop exposes compile-time mistakes that broad dynamic access would
defer until runtime, where they surface as binder exceptions. The compiler
checks member availability; real-Excel tests separately establish runtime
behavior. A missing typed member does not justify bypassing the compiler with
late binding; the document-property exception above was separately approved
after native tests ruled out the supplied and freshly imported interfaces.

The cost is a Windows/Excel dependency, COM lifetime management, and features
whose availability depends on the installed Excel version and desktop.

## Implementation and guidance

- [Excel batch](../src/ExcelMcp.ComInterop/Session/ExcelBatch.cs)
- [Protection metadata reader](../src/ExcelMcp.ComInterop/OleCompoundFileReader.cs)
- [Interop project](../src/ExcelMcp.ComInterop/ExcelMcp.ComInterop.csproj)
- [COM instructions](../src/ExcelMcp.ComInterop/AGENTS.md) and [PIA coverage](PIA-COVERAGE.md)
