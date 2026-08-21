# Legacy Excel COM API Coverage

This inventory records how ExcelMcp treats older and administrative Excel COM
surfaces. Features are included only when they are deterministic, non-interactive,
safe for an automation server, and testable against a real Excel instance.

## Implemented

| COM surface | ExcelMcp surface | Decision |
|-------------|------------------|----------|
| `Workbook.XmlMaps`, `XmlMaps.Add`, `XmlMap.Delete` | `xmlmap` list/add/delete | Typed PIA APIs provide deterministic workbook-scoped lifecycle operations. |
| `Range.XPath.SetValue` | `xmlmap map-range` | Maps a cell or single-column range without opening Excel UI. |
| `Workbook.XmlImportXml`, `XmlMap.ImportXml` | `xmlmap import-xml` | Imports XML already in memory. Existing maps and automatic destination mapping are supported without file pickers. |
| `XmlMap.ExportXml` | `xmlmap export-xml` | Returns XML in memory and avoids server-side output path ambiguity. |
| Built-in and custom document properties | `workbook` list/get/set/delete properties | Supports deterministic workbook metadata without opening property dialogs. |
| Workbook scenarios | `analysis` scenario lifecycle and summary actions | Manages named changing-cell sets and creates Excel summary reports. |
| Workbook links | `workbook` list/update/break external links | Exposes explicit link inspection and mutation without update prompts. |
| Worksheet page setup | `sheetstyle` get/set page setup | Covers orientation, fit-to-page, and centering metadata without invoking print UI. |

XML content is parsed with DTD processing disabled. XML schemas and import data
must each use exactly one inline value or readable file alias (`schemaFile` /
`xmlDataFile` in batch JSON, snake_case in MCP, kebab-case in the CLI), and
XSD `import`/`include`/`redefine` dependencies are rejected so Excel cannot implicitly fetch
external schemas. XML imports also reject `xsi:schemaLocation` and
`xsi:noNamespaceSchemaLocation` before invoking Excel, preventing HTTP, UNC, or
local-file schema resolution during automatic mapping.

## Explicit Exclusions

| COM surface | Status | Reason |
|-------------|--------|--------|
| Application/workbook/worksheet events | Excluded | Callback timing depends on user activity, add-ins, calculation, and Excel message pumping. Persistent subscriptions do not fit request/response command semantics and cannot be tested deterministically. |
| `Application.Dialogs`, `Dialog.Show`, file pickers | Excluded | Modal and interactive. These calls can block a headless server indefinitely and require a foreground desktop/user response. |
| `CommandBars`, controls, Ribbon/UI customization | Excluded | UI-only, add-in-dependent, and largely superseded by Ribbon extensibility. State varies by Excel version and installed add-ins. |
| `SendMail`, `SendForReview`, routing slips, mail envelopes | Excluded | Sends external communication, depends on a configured mail client/account, may display security prompts, and creates privacy-sensitive side effects. |
| Smart tags and smart-tag actions | Excluded | Deprecated/removed from modern Office workflows and not reliably available across supported Excel versions. |
| DDE, `ExecuteExcel4Macro`, XLM registration APIs | Excluded | Security-sensitive legacy code execution and process communication surfaces are not appropriate for a general automation server. |
| `XmlMap.Import` / `XmlMap.Export` URL and file variants | Excluded | In-memory import/export provides the same automation capability without implicit network access, overwrite ambiguity, or server-side path side effects. |
| Protected View trust bypass, macro-security changes, add-in installation | Excluded | Security-policy changes must remain user/admin controlled and cannot be safely automated as workbook operations. |

## Assessed for Future Work

The following typed APIs are deterministic enough to reconsider when a concrete
automation workflow and real Excel tests justify them:

- Workbook custom views.

These are not placeholders or promised actions. They remain unimplemented until
their behavior, safety defaults, and cross-version test coverage are defined.
