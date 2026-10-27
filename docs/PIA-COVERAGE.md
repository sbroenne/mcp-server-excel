# PIA Coverage Guide

This document tracks `Microsoft.Office.Interop.Excel` coverage, known runtime
behavior, and remaining verification limits. The compiler enforces typed
access in both production and test projects.

---

## TL;DR — Current Status

| Area | Status |
|------|--------|
| Core Excel types (`Workbook`, `Worksheet`, `Range`, etc.) | ✅ Typed in the PIA |
| PivotTable and Chart command code | Typed; full solution compilation restored |
| Table/slicer, Range, ConditionalFormat, Screenshot, and Sheet command code | Typed; full solution compilation restored |
| Drawing geometry/text and Forms controls | ✅ Typed |
| Shape creation/type/fill/line and layout members | Typed using embedded Excel and Office Core definitions; selected real-Excel Service cases passed |
| Sparklines (`SparklineGroups`, `SparklineGroup`, colors/points) | ✅ Typed |
| What-If Analysis (`GoalSeek`, `Scenarios`, `Table`) | ✅ Typed |
| QueryTables and threaded comments | ✅ Typed via the 16.x Excel PIA |
| XML maps (`XmlMaps`, `XmlMap`, `XmlImportXml`, `XmlMap.ExportXml`) | ✅ Typed; external schema resolution is rejected before COM |
| PivotCache options, grouping, drill-through, and chart plotting/formatting | ✅ Typed |
| Workbook metadata, properties, Save As/copy, fixed-format export, and links | Typed; document properties use the approved isolated adapter described below |
| Workbook themes and Document Inspector modules | Typed with embedded Office definitions; native and selected Service checks passed |
| Collections (`Sheets`, `Names`, `ListObjects`, etc.) | ✅ Fully typed |
| Power Query (`Workbook.Queries`, `WorkbookQuery`) | ✅ Typed via the 16.x Excel PIA |
| DataModel (`ModelTables`, `ModelRelationships`, etc.) | ✅ Typed |
| `ModelTableColumn.DataType` | ✅ Typed; list/read responses expose raw OLE DB type codes and observed names |
| DataModel measures/formats (`ModelMeasures`, `ModelMeasure`, `ModelFormat*`) | ✅ Typed via the 16.x Excel PIA |
| Connection sub-types (`OLEDBConnection`, `ODBCConnection`, `TextConnection`) | Typed; text/web imports use linked QueryTables for their available settings |
| `ModelTableColumn.IsCalculatedColumn` / formula / expression | ❌ Not exposed by the 16.x Excel PIA; runtime COM also has no reliable calculated-column metadata API |
| `ModelTable.LastRefresh` | ❌ Missing from the 16.x Excel PIA and unavailable at runtime; no reliable table refresh timestamp |
| Data Model source/model connection metadata | ✅ Typed (`SourceWorkbookConnection`, `DataModelConnection`, `ModelConnection`, `ModelTables`, command metadata) |
| Data Model live refresh status | ❌ No `Model.Refreshing` or `ModelTable.Refreshing` PIA/COM surface; refresh calls are synchronous |
| `ModelMeasure.FormatInformation` | Typed as `object`; readers select the supported typed `ModelFormat*` interface |
| `WorkbookConnection.Refreshing` / `CancelRefresh` | Missing on the base type; existing refresh controls use typed OLEDB/ODBC sub-connections |
| VBA (`VBProject`, `VBComponents`) | Typed using embedded VBE definitions; selected Service cases and published CLI/MCP status calls passed |
| `AutomationSecurity` | Typed Office Core enum; startup assignment verified in real Excel without an external interop DLL |
| WebConnection | Absent from the compiler definitions and actual workbook-connection interface; real-Excel tests establish typed QueryTable access for text/web refresh settings |
| ADO connection, recordset, fields, schema, and errors | Typed using build-only Microsoft imports; native, selected Service, and published CLI/MCP queries passed; repeated published MCP model-query/save/reopen checks passed |

---

## Typed access and runtime limits

EXL011 rejects all dynamic declarations and late-bound operations.
Write the desired operation against the typed PIA: the normal build checks
whether that API is available. A missing member is a capability limitation,
not permission to bypass the compiler with dynamic access.

### Compilation and verification boundary

- **Status**: The full solution builds with zero warnings and errors, including production, ordinary tests, and diagnostic tests. Connection helpers and test verification helpers now use typed interfaces.
- **Build enforcement**: EXL011 rejects dynamic declarations and late-bound operations in production and test projects. EXL005 rejects casts to `dynamic`; neither check has a legacy-file exemption.
- **Full-suite evidence**: The Windows x64 run on 2026-10-25 passed all normal selections: 2,843 cases without Excel and 2,621 Excel cases, including VBA and screenshots. Another 48 ComInterop infrastructure cases passed. Complete acceptance added one CLI case, one MCP case, and 20 external-cube cases: 5,534 unique cases in total, with no failures or skipped cases. Japanese-locale execution, ARM64 execution, configured IRM workbooks, and separate research-only diagnostics remain outside this evidence.
- **Entry points**: CLI/MCP package checks and output inspection passed. Freshly published CLI calls passed VBA status, an ADO query, save/reopen, and exact reopened model totals. After the save fix below, the exact freshly published MCP workflow passed ten fresh-process repetitions: create `.xlsm`, inspect VBA status, load two Power Query rows to the model, evaluate DAX, close with save, reopen, and verify both the marker cell and the model's actual total `18` and row count `2`. Both freshly published entry points also passed text/web relocation, save/reopen, settings/description checks, and changed-source refresh while retaining source and unrelated cells. The complete `Test-E2E.ps1 -SkipBuild` command passed all three stages: CLI, MCP, and external OLAP. The external stage used the local synthetic Atoti cube; it does not establish compatibility with every external server.
- **Save failure and fix**: The failure predates the conversion: the exact workflow failed twice against freshly compiled, separately published main at `a92b11b94ab6fa04e7de129c4d22504cd873f82d`; the converted build passed once and failed once in the paired comparison. Native investigation later observed `Save` return with `Saved == true`, followed by disabling `ScreenUpdating` changing it to false without a workbook edit. `ExcelWriteGuard` now leaves screen updating unchanged when any open workbook in that Excel application has a Data Model. It still suppresses repainting for ordinary workbooks, leaves events/calculation alone, and restores only a setting it changed. The model connection lifetime is unchanged. No forced saved flag, extra save, delay, or relaxed error check is used; save failures still retain the session. Twelve public Service cases passed after this fix, covering VBA inspection on/off, typed/untyped model columns, three repetitions, and exact reopened cells/model results.
- **Tracking**: [#1140](https://github.com/sbroenne/mcp-server-excel/issues/1140) tracks checked Excel access and real-Excel verification.

### Power Query — `Workbook.Queries` / `WorkbookQuery`

- **Status**: Typed via `Microsoft.Office.Interop.Excel` 16.x.
- **Affected files**: `ComUtilities.cs` and Power Query command files now use `Excel.Queries` and `Excel.WorkbookQuery`.

### DataModel measures/formats — `ModelMeasures` / `ModelMeasure` / `ModelFormat*`

- **Status**: Typed via `Microsoft.Office.Interop.Excel` 16.x.
- **Affected files**: Data Model measure commands now use `Excel.ModelMeasures`, `Excel.ModelMeasure`, and typed `ModelFormat*` properties.

### WorkbookConnection refresh status/cancel

- **Base-type limitation**: `WorkbookConnection.Refreshing` and `WorkbookConnection.CancelRefresh` do not compile against `Microsoft.Office.Interop.Excel` 16.0.18925.20022.
- **Current implementation**: `ConnectionCommands.RefreshControl.cs` reads and cancels through typed `OLEDBConnection` and `ODBCConnection` members. Other connection types report that status/cancellation is unsupported; the missing base members are not needed.

### Data Model calculated columns

- **Compile probe result**: `ModelTableColumn.IsCalculatedColumn`, `Formula`, and `Expression` all fail with CS1061 against `Microsoft.Office.Interop.Excel` 16.0.18925.20022.
- **Runtime result**: Excel reliably exposes only column name, data type, and parent through `ModelTableColumn`; calculated-column formulas and mutation are not practical automation surfaces.
- **Current behavior**: `list-columns` and `read-table` return explicit `isCalculated: null` (unknown), not an inferred false value. They preserve the raw `ModelTableColumn.DataType` code and report its observed OLE DB name: `5` = `DOUBLE`, `7` = `DATE`, `11` = `BOOLEAN`, `20` = `BIGINT`, and `130` = `WCHAR` (Unicode text). Other codes are reported as `Unknown (code)`. Use Power Query for computed columns or a DAX measure for query-time calculations.
- **Runtime quirk**: Although the Excel documentation describes `XlParameterDataType`, the model column property returns OLE DB type codes at runtime. Interpreting `5` through `XlParameterDataType` incorrectly labels doubles as `SMALLINT`. These labels describe the loaded model column, not the source Power Query type or a DAX expression's result type; a `Currency.Type` Power Query column can load as `DOUBLE`. Labeling does not convert or round values.
- **Catalog limitation**: A native ADO probe loaded one worksheet table into the model and verified its data by querying a sum of 18. `TMSCHEMA_COLUMNS` exposed the requested `TableID`, `ExplicitName`, and `Type` fields but returned zero rows for that loaded model. Absence from that catalog cannot establish that a column is not calculated. The user-approved unknown result is represented by a nullable `IsCalculated` property whose null is retained in JSON even when other null fields are omitted.

### ModelTable refresh metadata and status

- **Compile probe result**: `ModelTable.LastRefresh`, `ModelTable.Refreshing`, and `Model.Refreshing` fail with CS1061 against `Microsoft.Office.Interop.Excel` 16.0.18925.20022.
- **Runtime result**: After a successful explicit `ModelTable.Refresh()`, late-bound `LastRefresh` access still fails, so the timestamp is not exposed.
- **Current behavior**: `refresh` uses the typed synchronous `Model.Refresh()` / `ModelTable.Refresh()` APIs.
- **Hard limitation**: There is no reliable refresh timestamp, live refresh-status, or cancellation API for the Data Model itself.

### Data Model connection metadata

- **Compile probe result**: `Model.DataModelConnection`, `ModelTable.SourceWorkbookConnection`, `WorkbookConnection.ModelConnection`, `WorkbookConnection.ModelTables`, `WorkbookConnection.InModel`, and `ModelConnection.CommandText` / `CommandType` / `ADOConnection` compile against `Microsoft.Office.Interop.Excel` 16.0.18925.20022.
- **Runtime finding**: `Model.DataModelConnection.ModelTables` throws `0x800A03EC` for Excel's embedded model connection. `read-connection` therefore enumerates the reliable typed `Model.ModelTables` collection instead.
- **Runtime finding**: A worksheet-backed `ModelTable.SourceWorkbookConnection` exposes source connection metadata, but its `ModelConnection` property throws `0x800A03EC`; command metadata is therefore exposed only by `read-connection` for the actual model-type connection.
- **Current behavior**: `read-connection` exposes non-sensitive model connection metadata and table names. The raw ADO connection is intentionally not serialized because it may contain provider/session details.

### VBProject / VBComponents / VBComponent

- **Definitions**: `Microsoft.Vbe.Interop.dll` from the pinned `ExcelDna.Interop` package download. The build embeds used types; this definition assembly must not be deployed or loaded at runtime.
- **Evidence**: Compiler tests emit VBA project/module access without an external VBE assembly reference. A real-Excel ComInterop test creates a standard module, writes and reads exact source lines, checks project protection/mode, and confirms no separate interop assembly was loaded.
- **Current implementation**: VBA status, references, source reads/searches/edits, module lifecycle, and macro execution use typed Excel/VBE objects. Selected real-Excel Service cases, including argument-boundary/security-restoration checks, passed. Published CLI and MCP status calls also returned the expected project access, protection, and mode.

### AutomationSecurity

- **Definition**: `Microsoft.Office.Core.MsoAutomationSecurity` from the embedded Office Core definitions.
- **Implementation**: `ExcelBatch.cs` assigns the checked `AutomationSecurity` property, without reflection. Its normal force-disable setting is checked in real Excel. An explicitly approved macro reopen temporarily changes and restores that setting. `VbaCommands.Run` does not lower it and restores the prior cultures.
- **Deployment rule**: No separate `office.dll`, Excel PIA, or VBE interop DLL may be copied into published product output or listed as a runtime dependency. No assembly resolver or machine-specific hint path is allowed.

### Drawing shapes — Office.Core enum and formatting signatures

- **Compile probe result**: `Excel.Shapes`, `Excel.Shape`, geometry, `TextFrame`, `Characters`, `Font`, `ControlFormat`, and `AddFormControl` compile through the Excel PIA.
- **Office Core definitions**: `Shapes.AddPicture`, `AddShape`, `AddTextbox`, and `AddConnector` require `MsoTriState`, `MsoAutoShapeType`, `MsoTextOrientation`, and `MsoConnectorType`. Reading `Shape.Type` also requires those definitions. A compiler test retains Excel-only input to prove the missing-reference failure; another supplies and embeds all three definition assemblies.
- **Current implementation**: Drawing creation, formatting, visibility, grouping layout, and ordering use typed Excel objects and Office enums. `FillFormat`, `LineFormat`, `ColorFormat`, and `ConnectorFormat` are exposed by the Excel PIA itself; their enum signatures require Office Core.
- **Runtime evidence**: A real-Excel ComInterop test creates a rectangle and checks exact name, dimensions, type, fill type, and color without loading external interop assemblies. Selected Service drawing cases also passed; neither selection establishes every drawing operation.
- **Sparklines stay typed**: `Range.SparklineGroups`, `SparklineGroups.Add`, `SparklineGroup.ModifySourceData`, `Type`, `SeriesColor`, `Points`, and `Delete` are all available through the Excel PIA.
- **Affected files**: `DrawingCommands.Objects.cs`.

### Workbook themes and document inspection

- **Definitions**: The embedded Office definitions expose `OfficeTheme`, `ThemeColorScheme`, `ThemeFontScheme`, `ThemeFonts`, `ThemeFont`, `DocumentInspectors`, and `DocumentInspector`.
- **Font access**: `ThemeFontScheme.MajorFont` and `MinorFont` each return an owned `ThemeFonts` collection. Its typed `Item(MsoFontLanguageIndex)` method selects the Latin, complex-script, or East Asian font. Both the collection and each acquired font are released.
- **Native evidence**: A real-Excel ComInterop case sets and reads an exact accent color and six distinct major/minor font names, calls each available inspector with its typed status output, and checks that inspection retains the changed accent. It confirms no external interop definition assembly was loaded.
- **Public behavior**: Selected Service cases passed checks for exact saved theme definitions, persistence, invalid inputs, inspector findings, and selective fixes. Their verification helpers use typed interfaces.

### Document-property interface limitation

- **Compile result**: The pinned Office definitions expose `DocumentProperties` and its property collection methods; the definitions embed successfully without an external Office assembly reference.
- **Native failure**: Both `Workbook.BuiltinDocumentProperties` and `Workbook.CustomDocumentProperties` return objects that reject a cast to `Office.DocumentProperties`, interface `{2DF8D04D-5BFA-101B-BDE5-00AA0044DE52}`, with `InvalidCastException` / `0x80004002` (`E_NOINTERFACE`). Separate real-Excel cases retain and release each raw returned object and confirm this exact limitation.
- **Reported interfaces**: Both collections report Office type library `{2DF8D04C-5BFA-101B-BDE5-00AA0044DE52}`, version 2.8, and the same `DocumentProperties` interface identity as the compiler definitions. Their single reported base is standard `IDispatch`. Standard .NET enumeration can acquire an individual built-in property, but that item also rejects its reported `DocumentProperty` interface `{2DF8D04E-5BFA-101B-BDE5-00AA0044DE52}` with `0x80004002`. Enumeration alone therefore does not supply typed property reads, writes, or custom-property creation.
- **Fresh Microsoft import**: An isolated investigation imported the installed, Microsoft-signed Office 16.0.20430.20182 type library with Microsoft `TlbImp.exe` 4.8.9037.0 and an explicit import of signed Windows `stdole2.tlb`. Fresh definitions retain the relevant identities and signatures; `/transform:dispret` does not change them. Actual Excel calls using freshly imported, embedded definitions still fail with `0x80004002` for both collections and an individual built-in item. The isolated output contains no additional Office/stdole interop DLL, emitted assembly reference, or runtime dependency. No installed-Office build prerequisite or product reference change was made.
- **Approved isolated adapter**: After those failures were established, the user approved native COM access for document properties alone. `ExcelDocumentProperties` opens only the workbook's built-in/custom collections, keeps `IDispatch` private, obtains member identifiers through Microsoft's documented `GetIDsOfNames`, and permits only `Count`, `Item`, `Add`, `Name`, `Type`, `LinkToContent`, `Value`, and `Delete`. Returned items must report the Office `DocumentProperty` identity. It preserves native error codes/causes and releases argument variants and owned results, including failed operations. No dynamic, reflection fallback, public general-dispatch helper, extra definition DLL, or EXL011/EXL005 exemption was added.
- **Regression evidence**: Three focused actual-Excel cases first failed against the matching typed Office interfaces, then passed with the adapter. The expanded selection passed nine cases, including misuse/disposal checks and the supporting interface-limit tests. They check built-in reads/writes, all five custom-property types, indexed enumeration, replacement/deletion, retained neighboring properties, native errors including the unreadable built-in date, and exact values after save/reopen. A missing item preserves Excel's `0x80070057` cause and the original `0x80020009` dispatch result. The published ComInterop library matches the tested bytes and has no external interop assembly references or deployed definition DLLs; actual calls also check that none were loaded. Production metadata commands and the Service verification helper use the adapter; selected public Service cases passed too.
- **Native contract**: [Microsoft's `IDispatch::GetIDsOfNames`](https://learn.microsoft.com/en-us/windows/win32/api/oaidl/nf-oaidl-idispatch-getidsofnames) and [`IDispatch::Invoke`](https://learn.microsoft.com/en-us/windows/win32/api/oaidl/nf-oaidl-idispatch-invoke) define the interface layout, reversed argument order, property-assignment named argument, and error transport used here. The exception does not authorize native dispatch for another feature.

### WebConnection

- **Simply not in the Excel PIA**: `WorkbookConnection.WebConnection` is not exposed in `Microsoft.Office.Interop.Excel`. Only `OLEDBConnection`, `ODBCConnection`, and `TextConnection` have typed sub-connection properties.
- **Current implementation**: `ConnectionCommands.cs` uses typed provider interfaces or linked QueryTables for available source and refresh settings. Native QueryTable probes additionally establish text delimiters/column selection and selected web tables. These are extraction settings, not SQL `commandText`, and do not add public connection options. OLEDB/ODBC SQL or table command text remains available through their typed provider interfaces. Native and selected Service cases check source selection and actual imported cells.
- **Legacy load-to**: Existing text/web imports move through Excel's typed `QueryTables.Add(sourceQueryTable, destination)` overload, which copies their native import settings. After a successful synchronous refresh, the original QueryTable is deleted without clearing its cells, and the copied connection takes its original public name and description. Restoring the description uses the existing linked-import property path: assigning only `WorkbookConnection.Description` on the copy returned the correct in-memory value but lost it after reopening. The standard property path preserves it. Exactly one linked import is required. Loading again at the existing A1 refreshes it directly rather than copying and shifting its cells. Fifteen real-Excel Service cases cover delimited/fixed-width text, selected web tables, extraction/format/refresh settings, changed-source refresh, unrelated cells, same-sheet moves, protected source/target sheets and refresh failures. If the copied import cannot refresh, both imports and connections remain; the error identifies their locations and names. Cancellation and timeout keep their original meaning, and no rollback or automatic copy deletion is performed. Four additional public Service cases pass save/reopen and changed-source refresh checks for text/web, with moved and unmoved controls, including exact settings, description, values, and retained source/unrelated cells. Freshly published CLI and MCP calls also pass both text and web relocation, reopened settings/description, changed-source values, and retained source/unrelated cells.

### ADO compiler definitions

Microsoft's `TlbImp.exe` 4.8.9037.0, installed by the .NET Framework 4.8.1
Developer Pack, imports Windows' Microsoft-signed `msado28.tlb`. The fixed
arguments are `/namespace:ADODB /machine:Agnostic /asmversion:2.8.0.0
/strictref:nopia /silence:3001`; `/primary` is deliberately absent because this repository
does not publish Microsoft's type library. Both input signatures are checked;
tool version is pinned and tool/type-library file versions and SHA256 hashes
are recorded in the build log. Windows servicing can change the signed input
file; reproducing a specific import requires that recorded input version too.
`/strictref:nopia` explicitly ignores registered primary interop assemblies.
The narrowly selected `/silence:3001` removes only the notice that such an
ignored assembly is registered; other importer diagnostics remain enabled.
The repeated-import check rejects warning output and verifies identical
embedded application bytes from two fresh imports.

`Directory.Build.targets` owns generation under each project's `obj` directory,
incremental input/output checks, clean tracking, and compiler-only embedding.
Builds fail explicitly when the tool or Windows definitions are missing.
The importer is run directly from ordinary `dotnet` builds, not through
`COMReference` or full-framework MSBuild. The generated DLL is never committed
or a runtime package asset. A compiler test emits real connection, recordset,
field, schema, and error access with no `ADODB` or Excel assembly reference.

A real-Excel test loads a small worksheet table into the model and uses its
existing ADO session. It checks exact query columns/values, end-of-recordset,
closed recordset state, cube schema, and a second successful query after a
failed one. The failure code and provider error records are compared with
native VBA access to the same connection object. VBA and .NET wrapper error
wording differs; provider null strings retain the existing empty-string
presentation. The test does not require a provider to supply error records
when native access supplies none.

A clean ComInterop build removes and regenerates the definition file.
Reimporting leaves the deterministic product assembly unchanged; an unchanged
build reuses the generated input. Freshly published ComInterop output has no
external Excel/Office/VBE/ADO DLL, assembly reference, or runtime DLL asset in
its dependency file. CLI/MCP distribution inspection and selected actual
Service/entry-point calls have since passed, including repeated published MCP
model-query/save/reopen checks after the screen-update guard fix described above.
Definition-dependency checks alone do not establish save behavior.

Repeated imports are not byte-identical: the Microsoft importer generates
assembly identity bookkeeping such as a fresh module identifier. Do not claim
byte-for-byte reproducibility of that intermediate DLL. The consumed contract
is ADO 2.8 with stable interface identities; deterministic application output
and actual Excel execution must be checked separately. Neither a successful
import nor clean published output establishes every workbook behavior.

---

## Migrated Typed Surfaces

These were incorrectly left as `dynamic` during earlier PIA migration work due
to false negatives from binary string search on the assembly. They are now
typed; the build rule prevents adding the same kind of dynamic fallback again.

### DataModel typed surfaces — `Excel.Model`, `Excel.ModelTables`, etc.

- **Types confirmed in PIA**:
  - `Excel.Model` (`workbook.Model`)
  - `Excel.ModelTables`, `Excel.ModelTable`
  - `Excel.ModelRelationships`, `Excel.ModelRelationship`
  - `Excel.ModelTableColumns`, `Excel.ModelTableColumn`
- **Affected files**: `DataModelCommands.Helpers.cs` and all DataModel command files
- **Migration status**: Complete for table/relationship surfaces; measure-specific APIs are typed through the 16.x PIA.

### Connection Sub-Types — `Excel.OLEDBConnection`, `Excel.ODBCConnection`, `Excel.TextConnection`

- **Compiler definitions**: All three sub-connection interfaces exist. The collection is `Excel.Connections`, not `WorkbookConnections`.
- **Actual Excel interfaces**: Text and web QueryTables expose typed `BackgroundQuery`, `RefreshOnFileOpen`, `SavePassword`, and `RefreshPeriod` getters/setters. Two actual-Excel cases set and read those fields, identify their workbook connections, and confirm no external interop assembly was loaded.
- **Missing members**: Native type information confirms that `WorkbookConnection` has neither `WebConnection` nor `ConnectionString`. `TextConnection` exposes `Connection` and text-import options, but none of the four refresh settings above or `CommandText`. Its linked QueryTable exposes the settings instead; the text case verifies that `WorkbookConnection.Ranges` leads back to the same QueryTable and connection.
- **Corrected connection labels**: The official definitions give `TEXT=4`, `WEB=5`, and `XMLMAP=3`; actual text/web connections confirm the first two values. `ConnectionHelpers.GetConnectionTypeName` and provider selection now use those definitions rather than the previous incorrect mapping of `4` to `WEB`, `5` to `XMLMAP`, and `3` to `TEXT`.
- **Command text finding, not an extraction-definition limitation**: Reading `QueryTable.CommandText` for text/web imports fails with `OutOfMemoryException`, HRESULT `0x8007000E`, in the installed Excel build. Two additional native cases explicitly test both an inactive owning sheet and the active owning sheet, as required by [Microsoft's CommandText documentation](https://learn.microsoft.com/en-us/office/vba/api/excel.querytable.commandtext), and repeat the read after refresh through the linked QueryTable. Activation does not change that result. This establishes a failing typed getter, not absence of the import definition; production must not catch arbitrary memory failures as missing data.
- **Readable and changeable extraction settings**: [QueryTable.Connection](https://learn.microsoft.com/en-us/office/vba/api/excel.querytable.connection) holds the text-file path or web URL. [WebTables](https://learn.microsoft.com/en-us/office/vba/api/excel.querytable.webtables) and `WebSelectionType` control which HTML tables are imported. `TextFileStartRow`, delimiter/quoting properties, and [TextFileColumnDataTypes](https://learn.microsoft.com/en-us/office/vba/api/excel.querytable.textfilecolumndatatypes) control text parsing and skipped columns. These compile against the pinned definitions and work in actual Excel.
- **Extraction outcome evidence**: `Execute_LegacyExtractionDefinitions_PreserveSelectionAndRefresh` passes for text and web. The text case skips a preamble and middle source column, retains identifiers `0007`/`0008` as text, and checks amounts `17`/`29`; changing the column settings then refreshes the previously skipped column instead. The web case imports only the second of two local HTML tables, checks its exact cells, switches to the first table, and verifies replacement. Both confirm the source file is unchanged, discover the linked QueryTable through `WorkbookConnection.Ranges`, and load no external interop definition assembly.
- **Public contract distinction**: `connection.create` documents OLEDB/ODBC SQL query text or a table name and rejects direct text/web sources. Existing `querytable.create-text` / `create-web` commands own direct imports. `querytable.view` returns source location, text delimiter/encoding, and web selection/formatting settings; it does not expose every native text parsing option, including start row and column types. Connection readers do not call the failing text/web SQL-command getter; they retain available source and refresh settings instead. Selected Service cases check these paths without adding public import options.

### Conditional-formatting rule interfaces

- **Implementation**: Creation, reads, and editing use the seven supported interfaces: `FormatCondition`, `ColorScale`, `Databar`, `IconSetCondition`, `Top10`, `AboveAverage`, and `UniqueValues`. Shared formatting helpers select typed getters, not object/property-name dispatch.
- **Actual Excel evidence**: A native case creates all seven types and verifies exact type numbers, priorities, and affected ranges through those interfaces. Three additional cases establish unset `Font.Bold`/`Italic` as `DBNull`, distinct from explicit false and true. The read helpers retain null for unset values and catch only known unavailable-member errors. Full solution compilation is restored; the native selection is not a substitute for the complete conditional-formatting Service selection.

### Chart formatting and forecast periods

- **Checked interfaces**: Chart strategies, lookup, source-range resolution, geometry, axes, trendlines, and material formatting now name their Excel interfaces. `ChartFormat`, `FillFormat`, `LineFormat`, and `ColorFormat` are supplied by the Excel definitions; line visibility uses the embedded Office enum. Existing Service material-verification helpers use those same interfaces.
- **Actual Excel evidence**: A focused case verifies exact chart geometry, a regular chart's null `PivotLayout`, the `ChartArea.Parent` chart identity, a two-series count and the first series' values/categories, fill/line colors, transparency, and line weight. It confirms that no external interop assembly was loaded. Selected public Service chart-formatting cases passed separately.
- **Column-point transparency**: Three public Service cases read back `0`, `0.25`, and `1` after writing those values, retain the expected fill color, omit an error when the value is available, and leave a neighboring point unchanged. The old test's expected unreadable getter does not apply to these column-chart points; marker transparency limits remain separate.
- **Preserved forecast rounding**: The existing public forecast inputs are doubles, but the old `Trendline.Forward` / `Backward` properties accept integers. Actual Excel type information confirms their `VT_I4` getters and hidden-member flags. Native checks in both directions establish that `0.5` rounds to `0`, `1.5` to `2`, and `2.5` to `2`. The four setters now use `Convert.ToInt32` to retain that whole-period rounding explicitly. Microsoft's documented `Forward2` / `Backward2` properties use doubles, but switching to them would change the existing numeric behavior and is not part of this conversion. Six Service cases passed for creating/updating fractional forecasts and unchanged series data.
- **Separate native restriction**: The native column-chart case rejects backward forecasts of `1` through the old property and `1.5` through the newer property with `0x80004005`. That error is not classified as a missing member or replaced with zero. The separate existing scatter-chart Service forecast case passed.

### PivotTable source names and visible members

- **Checked paths**: Creation, lifecycle, field strategies, calculated fields/members, selection snapshots, and slicers now use explicit Excel interfaces. The ADO hierarchy lookup uses the approved embedded ADO definitions. Unused helpers that returned either `CubeField` or `PivotField` were removed; each strategy has its own typed field getter.
- **Values marker**: Excel's synthetic multi-value field exposes the `#N/A` error marker as numeric text through the PIA's string-valued `SourceName`. Checking only the old integer marker would include that synthetic field; checking only the text would incorrectly omit a real column with the same name. The reader confirms both the marker and the owning PivotTable's `DataPivotField.Name`, and does not request that field when fewer than two value fields exist.
- **Visible members**: `PivotField.VisibleItems` is an indexed property, not a method. `field.VisibleItems[Type.Missing]` returns the typed `PivotItems` collection; `field.VisibleItems[index]` returns an individual member. No native-dispatch exception is required.
- **Actual Excel evidence**: Four cases check zero, one, and two value fields alongside real columns named `-2146826246` and `Values`, and verify a two-member filter becomes exactly one visible `North` member while `South` is hidden. They retain the source data and confirm no external interop assembly is loaded. The combined selection with two text/web settings cases and two extraction-selection cases passes eight tests.
- **Public verification**: Selected multi-value and filter snapshot Service cases passed after compilation was restored. Native interface checks do not replace the wider public workbook checks.

### ComUtilities Helpers — Return Types

- `FindQuery(Excel.Workbook workbook, ...)` returns `Excel.WorkbookQuery?`
- `FindConnection(Excel.Workbook workbook, ...)` returns `Excel.WorkbookConnection?`
- `FindSheet(Excel.Workbook workbook, ...)` returns `Excel.Worksheet?`
- `FindName(Excel.Workbook workbook, ...)` returns `Excel.Name?`
- `ResolveRange()` in `RangeHelpers.cs` returns `Excel.Range?`

### Style indentation and mixed formatting

`Style.IndentLevel` is declared as `int`, but Excel returns a native null for a
new style, including after first setting its indentation to zero. A native VBA reference
read using an object returns `DBNull` through COM; the typed C# getter instead
throws `InvalidCastException` with `0x80004002`. Setting indentation to two
returns two. Alignment/protection inclusion flags do not change these cases.
Resetting two to zero returns integer zero, which the helper preserves.

`ComUtilities.ReadStyleIndentLevel` owns only this getter and converts only
that exact exception type/code to null. Core continues to omit the nullable
field and include `indentLevel` in `mixedFields`; it does not invent zero.
Other COM exceptions, cancellation, state errors, and different conversion
codes propagate unchanged. This is not a general interpretation of
`0x80004002` or permission to treat arbitrary failed reads as absent values.
Mixed range formatting remains variant-valued: compiler-checked reads retain
`DBNull` rather than coercing it to false or zero.

### Workbook open/save and regional formats

`Workbooks.Open` and `Workbook.Save` now use the Excel PIA directly. Both
declarations carry `LCIDConversion` attributes, so matching method names alone
does not prove that formats survive the change from dynamic calls.
The ComInterop tests check saved table-column date and number formats, edited
values, and a second save/reopen cycle. The separate
`OpenAndSave_JapaneseTableDateFormat_PreservesTableColumnDxf` regression also
forces an edit before saving and checks the reopened value. It requires a
Japanese Windows and Excel installation; passing the general tests elsewhere
does not replace that check.

---

## Compiler enforcement

The [compiler analyzer](../src/ExcelMcp.Analyzers/Rules.cs) rejects dynamic
declarations, inferred dynamic locals, and late-bound operations in production
and test projects. EXL005 also rejects explicit casts to `dynamic`; comments do
not exempt them. Use typed Excel interop directly: the normal build then checks
whether the referenced PIA exposes the member. If it does not, the compile
error is the evidence of a missing typed API; do not replace it with dynamic
access. Keep runtime behavior and deployment constraints in this guide as
separate checks because compile-time member availability does not prove Excel
runtime behavior.

Embedding the Excel PIA makes some ordinary property results appear as
`dynamic`, including `Range.Formula`, `Range.Formula2`, and `Range.Value2`.
Reading these properties through an `Excel.Range` is allowed when the result
is captured as `object`, for example `object value = range.Value2;`.
Using `var` to retain a dynamic result or accessing further members on that
result is rejected. The analyzer tests use the actual embedded Excel PIA to
check this distinction.

### Build-only Office and VBA definitions

`Directory.Packages.props` pins `ExcelDna.Interop` 16.0.0 as a `PackageDownload`.
Only its Microsoft-signed `office.dll` and `Microsoft.Vbe.Interop.dll` definitions
are used. The newer `Microsoft.Office.Interop.Excel` package remains the Excel
definition source. `ForceEmbedExcelInteropTypes` adds the Office/VBE compiler
inputs only to projects already resolving the Excel PIA, embeds used types, and
keeps the definition DLLs out of product output.

Check three different things after changing this boundary: compiler tests must
emit actual Office/VBE calls without external interop assembly references;
real Excel must execute the affected calls without loading those assemblies;
published files and dependency manifests must contain no runtime interop
dependency. The analyzer tests copy definition DLLs under `AnalyzerFixtures`
solely as compiler inputs, not as product runtime files.

ComInterop and CLI/MCP distribution output have been checked for definition
DLLs, emitted assembly references, and dependency-manifest runtime entries.
None were found for Excel, Office Core, VBE, ADO, or stdole. Existing package
checks passed for local tool installation, native launchers, npm packages, and
standalone archives. ARM64 output was inspected but not executed on the x64
host. Actual x64 CLI/MCP calls exercised VBA, ADO, reopened model results,
and text/web relocation with retained import settings. The earlier
fresh-model save failure was reproduced on main before being fixed as described
above. The separate complete acceptance results are recorded above.

---

## How This Happened (Root Cause)

During the original PIA migration, coverage was checked using binary string search with `[System.Text.Encoding]::Unicode` on the PIA DLL. This method reads the binary assembly as UTF-16LE text, which produces garbage. All string searches return false negatives — every searched type name was reported as "NOT FOUND" even when present.

This caused 8 DataModel and 3 Connection sub-type APIs to be incorrectly classified as "PIA gaps" when they are in fact available in the referenced Excel PIA.

**Prevention**: Use typed PIA calls in product code and tests. The ordinary
build is the compile probe and rejects late-bound fallbacks; real-Excel tests
remain necessary to establish runtime behavior.
