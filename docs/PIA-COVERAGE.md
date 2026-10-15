# PIA Coverage Guide

This document tracks the status of `Microsoft.Office.Interop.Excel` type coverage across ExcelMcp, explains why some casts remain `dynamic`, and documents the verified methodology for checking PIA coverage.

---

## TL;DR — Current Status

| Area | Status |
|------|--------|
| Core Excel types (`Workbook`, `Worksheet`, `Range`, etc.) | ✅ Typed in the PIA |
| PivotTable, Table/slicer, Chart, Range, NamedRange, ConditionalFormat, Screenshot, and Sheet command code | ✅ Supported Excel objects use typed PIA access; narrow external-model and verified runtime exceptions remain |
| Drawing geometry/text and Forms controls | ✅ Typed |
| Shape creation/type/fill/line members | ⚠️ Office.Core signatures; isolated late binding |
| Sparklines (`SparklineGroups`, `SparklineGroup`, colors/points) | ✅ Typed |
| What-If Analysis (`GoalSeek`, `Scenarios`, `Table`) | ✅ Typed |
| QueryTables and threaded comments | ✅ Typed via the 16.x Excel PIA |
| XML maps (`XmlMaps`, `XmlMap`, `XmlImportXml`, `XmlMap.ExportXml`) | ✅ Typed; external schema resolution is rejected before COM |
| PivotCache options, grouping, drill-through, and chart plotting/formatting | ✅ Typed |
| Workbook metadata, properties, Save As/copy, fixed-format export, and links | ✅ Typed except Office document-property collections |
| Collections (`Sheets`, `Names`, `ListObjects`, etc.) | ✅ Fully typed |
| Power Query (`Workbook.Queries`, `WorkbookQuery`) | ✅ Typed via the 16.x Excel PIA |
| DataModel (`ModelTables`, `ModelRelationships`, etc.) | ✅ Typed |
| `ModelTableColumn.DataType` | ✅ Typed; list/read responses expose raw OLE DB type codes and observed names |
| DataModel measures/formats (`ModelMeasures`, `ModelMeasure`, `ModelFormat*`) | ✅ Typed via the 16.x Excel PIA |
| Connection sub-types (`OLEDBConnection`, `ODBCConnection`, `TextConnection`) | ✅ Migrated — callers use typed WorkbookConnection |
| `ModelTableColumn.IsCalculatedColumn` / formula / expression | ❌ Not exposed by the 16.x Excel PIA; runtime COM also has no reliable calculated-column metadata API |
| `ModelTable.LastRefresh` | ❌ Missing from the 16.x Excel PIA and unavailable at runtime; no reliable table refresh timestamp |
| Data Model source/model connection metadata | ✅ Typed (`SourceWorkbookConnection`, `DataModelConnection`, `ModelConnection`, `ModelTables`, command metadata) |
| Data Model live refresh status | ❌ No `Model.Refreshing` or `ModelTable.Refreshing` PIA/COM surface; refresh calls are synchronous |
| `ModelMeasure.FormatInformation` | ✅ Typed as `object`; dispatch uses typed `ModelFormat*` patterns |
| Connection refresh status / cancellation | ✅ Typed `OLEDBConnection` and `ODBCConnection`; unsupported types are reported explicitly |
| VBA (`VBProject`, `VBComponents`) | ⚠️ External VBE object model; narrowly documented late binding |
| `AutomationSecurity` | ⚠️ Office.Core enum (office.dll); accessed late-bound via `((dynamic)(object))` because Office.Core is not referenced. The embedded Excel PIA carries no office.dll dependency, so no resolver is needed |
| WebConnection and selected TextConnection properties | ⚠️ Missing from the 16.x Excel PIA; isolated late binding |
| ADO types (ADODB.Connection, Recordset, Fields) | ⚠️ External ADODB object model; narrowly documented late binding |
| Native `Workbooks.Open` / `Workbook.Save` dispatch | ⚠️ Intentional runtime workaround: typed calls inject a locale argument that can rewrite localized table formats |

---

## Remaining Dynamic Debt

ExcelMcp is PIA-first. Any `dynamic` usage is technical debt unless a compile probe proves the referenced Excel PIA lacks the API, or the API belongs to an external COM object model that is not currently referenced. Prefer adding typed interop coverage over spreading new `dynamic` calls.

### Converted command code

- **Supported surfaces**: Pivot tables/fields, list objects, slicers, charts, worksheets, ranges, named ranges, and conditional-format rule interfaces now use typed PIA variables and helpers.
- **Heterogeneous objects**: Conditional-format rules and connection subtypes use concrete typed dispatch instead of a general dynamic proxy. Shared lookup/range helpers accept typed workbooks and return typed Excel objects.
- **Original enforcement gap**: The retired regex recognized only `((dynamic)` casts and missed declarations, parameters, return types, nullable spellings, and alternative casts.
- **Tracking**: [#1140](https://github.com/sbroenne/mcp-server-excel/issues/1140) combines the staged conversions with compiler-loaded source rules and real-Excel regression coverage. Compilation alone does not establish Excel behavior.

### Power Query — `Workbook.Queries` / `WorkbookQuery`

- **Status**: Typed via `Microsoft.Office.Interop.Excel` 16.x.
- **Affected files**: `ComUtilities.cs` and Power Query command files now use `Excel.Queries` and `Excel.WorkbookQuery`.

### DataModel measures/formats — `ModelMeasures` / `ModelMeasure` / `ModelFormat*`

- **Status**: Typed via `Microsoft.Office.Interop.Excel` 16.x.
- **Affected files**: Data Model measure commands now use `Excel.ModelMeasures`, `Excel.ModelMeasure`, and typed `ModelFormat*` properties.

### Connection refresh status/cancel

- **Root API gap**: `WorkbookConnection.Refreshing` and `WorkbookConnection.CancelRefresh` do not compile against `Microsoft.Office.Interop.Excel` 16.0.18925.20022.
- **Typed alternative**: `ConnectionCommands.RefreshControl.cs` uses `OLEDBConnection.Refreshing` / `CancelRefresh` and their ODBC counterparts.
- **Other connection types**: Returned support fields distinguish unsupported inspection/cancellation from an idle supported connection.
- **Refresh-date runtime gap**: An OLAP `OLEDBConnection.RefreshDate` can be an empty native VARIANT, although the PIA declares a non-nullable `DateTime`. The existing late-bound read preserves this distinction.

### Data Model calculated columns

- **Compile probe result**: `ModelTableColumn.IsCalculatedColumn`, `Formula`, and `Expression` all fail with CS1061 against `Microsoft.Office.Interop.Excel` 16.0.18925.20022.
- **Runtime result**: Excel reliably exposes only column name, data type, and parent through `ModelTableColumn`; calculated-column formulas and mutation are not practical automation surfaces.
- **Current behavior**: `list-columns` and `read-table` preserve the raw `ModelTableColumn.DataType` code and report its observed OLE DB name: `5` = `DOUBLE`, `7` = `DATE`, `11` = `BOOLEAN`, `20` = `BIGINT`, and `130` = `WCHAR` (Unicode text). Other codes are reported as `Unknown (code)`. Use Power Query for computed columns or a DAX measure for query-time calculations.
- **Runtime quirk**: Although the Excel documentation describes `XlParameterDataType`, the model column property returns OLE DB type codes at runtime. Interpreting `5` through `XlParameterDataType` incorrectly labels doubles as `SMALLINT`. These labels describe the loaded model column, not the source Power Query type or a DAX expression's result type; a `Currency.Type` Power Query column can load as `DOUBLE`. Labeling does not convert or round values.

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

- **Location in COM**: `Microsoft.Vbe.Interop` (separate DLL — `vbe7.dll` / `VBA7.dll`)
- **Why not available**: No .NET 5+ compatible NuGet package exists. `ThammimTech.Microsoft.Vbe.Interop` targets .NET Framework only. The official `Microsoft.Vbe.Interop` NuGet package is unsigned and unmaintained.
- **Workaround**: `((dynamic)ctx.Book).VBProject` — COM late binding to the VBE object model
- **Affected files**: `VbaCommands.Lifecycle.cs`, `VbaCommands.Operations.cs`

### MsoAutomationSecurity (= 3)

- **Location in COM**: `Microsoft.Office.Core` (office.dll), a shared Office type library separate from the Excel PIA.
- **Why not typed**: `MsoAutomationSecurity` is an enum in `Microsoft.Office.Core`. We do not reference or embed the Office.Core PIA, so the typed enum constant is not available in our build.
- **Workaround**: Use the literal integer values (`1` = `msoAutomationSecurityLow`, `3` = `msoAutomationSecurityForceDisable`) via `((dynamic)(object)tempExcel).AutomationSecurity = ...;` — late-bound IDispatch access that exchanges a plain `int` with Excel and never touches an Office.Core type.
- **The `(object)` cast keeps the call late-bound**: Casting a typed `Excel.Application` directly to `dynamic` would let the DLR resolve `AutomationSecurity` against the typed `MsoAutomationSecurity` signature, which would require an Office.Core reference. Casting to `(object)` first erases the static type and forces pure IDispatch binding, so no Office.Core type is needed.
- **office.dll is NOT a runtime dependency**: The Excel PIA is embedded (`EmbedInteropTypes` via the repo-root `Directory.Build.targets`; the PackageReference is compile-only via `<ExcludeAssets>runtime</ExcludeAssets>`). Embedding bakes only the Excel interop types we actually use into our assemblies — none of which are Office.Core types — so the built assemblies carry **no** reference to `office.dll`. This is why no assembly resolver is required. (Previously the PIA was referenced but not embedded, so it dragged in a transitive `office v16.0.0.0` dependency that forced a runtime `office.dll` load; that is now gone.)
- **Do NOT add a `<Reference>` to office.dll or a hand-rolled assembly resolver**: A GAC hint path is version- and machine-specific (15.0 vs 16.0 mismatch). Proper PIA embedding removes the office.dll dependency entirely, so neither is needed.
- **Affected files**: `ExcelBatch.cs`

### Drawing shapes — Office.Core enum and formatting signatures

- **Compile probe result**: `Excel.Shapes`, `Excel.Shape`, geometry, `TextFrame`, `Characters`, `Font`, `ControlFormat`, and `AddFormControl` compile through the Excel PIA.
- **Office.Core gap**: `Shapes.AddPicture`, `AddShape`, `AddTextbox`, and `AddConnector` require `MsoTriState`, `MsoAutoShapeType`, `MsoTextOrientation`, and `MsoConnectorType` from `office.dll`. `Shape.Type`, `AutoShapeType`, `ConnectorFormat`, `Fill`, `Line`, `Visible`, and `LockAspectRatio` have the same dependency.
- **Workaround**: `DrawingCommands` uses typed Excel objects everywhere else and isolates pure IDispatch late binding to those Office.Core-dependent members, passing documented integer enum values.
- **Why not add office.dll**: The project intentionally embeds only Excel PIA types to avoid a runtime `office.dll` dependency. Adding the Office.Core PIA would regress self-contained deployment.
- **Sparklines stay typed**: `Range.SparklineGroups`, `SparklineGroups.Add`, `SparklineGroup.ModifySourceData`, `Type`, `SeriesColor`, `Points`, and `Delete` are all available through the Excel PIA.
- **Affected files**: `DrawingCommands.Objects.cs`.

### WebConnection

- **Compile probe result**: `WorkbookConnection.WebConnection` and root `ConnectionString` fail with CS1061 against `Microsoft.Office.Interop.Excel` 16.0.18925.20022. `OLEDBConnection`, `ODBCConnection`, and `TextConnection` have typed sub-connection properties.
- **Affected files**: `ConnectionCommands.cs` (`GetTypedSubConnection` method)

### TextConnection property gaps

- **Typed surface**: `Excel.TextConnection` and its `Connection` and text-parsing properties exist in the PIA.
- **Compile probe result**: `BackgroundQuery`, `RefreshOnFileOpen`, `SavePassword`, `RefreshPeriod`, and `CommandText` fail with CS1061 against `Microsoft.Office.Interop.Excel` 16.0.18925.20022.
- **Current behavior**: Only those unsupported property reads/writes remain late-bound. A missing typed declaration does not establish that Excel supports the property at runtime; existing error handling still applies.
- **Affected files**: `ConnectionCommands.cs` and `ConnectionCommands.Lifecycle.cs`.

### Native Open/Save locale handling

- **Typed surface**: `Workbooks.Open` and `Workbook.Save` exist in the PIA.
- **Runtime mismatch**: The typed interop signatures inject an LCID. That can rewrite locale-specific structured-reference formats during open/save.
- **Current behavior**: The session preserves native dispatch for these calls in `ExcelBatch.cs` and `ExcelShutdownService.cs`. This is a documented runtime-signature exception, not a missing Excel API.

### PivotField source-name error values

- **Typed signature**: `PivotField.SourceName` is declared as `string`.
- **Verified native behavior**: The regular pivot's Values placeholder exposes an integer Excel `#N/A` error instead. `PersistentServicePivotTableTests.Snapshot.cs` verifies the native value and its error mapping while checking the complete multi-value-field Service snapshot.
- **Current behavior**: Only this source-name read stays late-bound so the error sentinel is not mistaken for a valid source field.
- **No VisibleItems gap**: The typed `VisibleItems[Type.Missing]` getter returns the collection; the bounded-selection regression verifies its exact counts and member names.

### Style and Font value-type getters

- **Verified member-specific gap — `Style.IndentLevel`**: The 16.0.18925.20022 PIA declares this getter as non-nullable `int` (DISPID 201). Built-in `Normal` and a registered custom style captured from a cleared single cell return native `VT_NULL`, exposed as `System.DBNull` by IDispatch. The typed getter throws `InvalidCastException`, HRESULT `0x80004002`, with message `OleAut reported a type mismatch.` The custom-state probe also preserves explicit `IncludeAlignment == false`, `IncludeProtection == false`, and an included dashed left border.
- **Typed native binding**: The referenced PIA provides no object/nullable-returning Style getter for this member. `IStyleIndentLevelVariant` declares exactly one getter on the existing Style IDispatch interface (IID `00020852-0000-0000-C000-000000000046`, DISPID 201), with `object?` and `MarshalAs(UnmanagedType.Struct)` representing the native VARIANT. The helper uses a borrowed view of the same Style RCW, not a separately acquired COM object. It uses no C# `dynamic`, reflection, generic dispatch helper, or additional dependency.
- **Preserved behavior**: The existing formatting reader preserves `IndentLevel == null` and the exact `indentLevel` entry in `mixedFields`; Service JSON continues omitting null-valued properties, and initialized indentation values remain numeric. All other Style getters and `Font.ThemeFont` remain PIA-typed. Inclusion flags do not identify whether indentation is unset, and a cleared cell's typed `Range.IndentLevel` is `0`, not the definition's native null; reading applied cell formatting cannot replace the Style-definition getter.
- **Observed mixed-font behavior**: Native `Font.ThemeFont` returns integer `0`, matching typed `xlThemeFontNone`. Individual major/minor fonts retain their exact values `1` and `2`.
- **Supported style behavior**: Excluded alignment/protection flags do not make the stored Style getters inherently nullable. Typed and native reads agree for the supported seeded state, including the `XlOrientation` enum rather than a Range rotation angle.
- **Regression evidence**: `NativeNullableFormattingCapabilityTests.UnseededStyle_NativeGettersAndServicePreserveExactState` independently probes all ten Style scalars and `Font.ThemeFont`, asserting the exact native null and PIA failure only for `IndentLevel`, typed/native equivalence for the remaining getters, and exact Service null/mixed fields, inclusion flags, and border values. Both states failed the Service read before the fix. The exact typed native binding returns `DBNull` in both states and integer `2` in the initialized-style case; the tests verify interface IID agreement and identical RCW identity. The existing seeded, mixed-font, and differential-formatting capability cases retain their exact-value assertions.

---

## Migrated Typed Surfaces

Earlier migration work sometimes mistook a binary string-search miss for a
missing PIA API. Compile probes and typed callers establish the supported
surfaces below; these must remain typed.

### DataModel typed surfaces — `Excel.Model`, `Excel.ModelTables`, etc.

- **Types confirmed in PIA**:
  - `Excel.Model` (`workbook.Model`)
  - `Excel.ModelTables`, `Excel.ModelTable`
  - `Excel.ModelRelationships`, `Excel.ModelRelationship`
  - `Excel.ModelTableColumns`, `Excel.ModelTableColumn`
- **Affected files**: `DataModelCommands.Helpers.cs` and all DataModel command files
- **Migration status**: Complete for table/relationship surfaces; measure-specific APIs are typed through the 16.x PIA.

### Connection Sub-Types — `Excel.OLEDBConnection`, `Excel.ODBCConnection`, `Excel.TextConnection`

- **Compile test result**: All three sub-connection types compile cleanly
- **Note**: The collection is `Excel.Connections` (NOT `WorkbookConnections`)
- **Affected files**: `ConnectionCommands.cs` (`GetTypedSubConnection` helper)
- **Helper result**: `GetTypedSubConnection(Excel.WorkbookConnection)` returns `object?` because supported connection subtypes have different interfaces. Callers dispatch to the concrete PIA type rather than using a general dynamic proxy.

### ComUtilities Helpers — Return Types

- `FindQuery(Excel.Workbook workbook, ...)` returns `Excel.WorkbookQuery?`
- `FindConnection(Excel.Workbook workbook, ...)` returns `Excel.WorkbookConnection?`
- `FindSheet(Excel.Workbook workbook, ...)` returns `Excel.Worksheet?`
- `FindName(Excel.Workbook workbook, ...)` returns `Excel.Name?`
- `SafeGetString(object? comObject, string propertyName)` and `SafeGetInt(object? comObject, string propertyName)` dispatch to their supported typed interfaces; unsupported property/type pairs throw `ArgumentOutOfRangeException`.

---

## Verified Methodology for Checking PIA Coverage

### THE CORRECT METHOD: Compile test

Create a probe project **outside the repo** (to avoid Central Package Management interference):

```powershell
New-Item -ItemType Directory D:\temp\pia-probe
Set-Location D:\temp\pia-probe
dotnet new console
dotnet add package Microsoft.Office.Interop.Excel --version 16.0.18925.20022
```

Edit `Program.cs`:
```csharp
using Excel = Microsoft.Office.Interop.Excel;

// Test: is this type in the PIA?
Excel.Queries q = default!;         // Compiles → in PIA
Excel.WorkbookQuery wq = default!;  // Compiles → in PIA
Excel.OLEDBConnection oledb = default!;  // Compiles → in PIA
// etc.
```

Then:
```powershell
dotnet build
```

Use the exact version selected by `Directory.Packages.props`. A successful
compile establishes that the tested typed API exists; CS0246 identifies a
missing type and CS1061 a missing member. A compile probe does not establish
runtime support, ownership, or behavior; those require real-Excel tests.

### WRONG: Binary string search

```powershell
# ❌ This is WRONG — produces false negatives for ALL type names
[System.IO.File]::ReadAllText($dll, [System.Text.Encoding]::Unicode)
# .NET assemblies are binary; reading as UTF-16LE produces garbage
```

### OK (but less reliable): Reflection

```powershell
$asm = [System.Reflection.Assembly]::LoadFrom($dllPath)
try { $types = $asm.GetTypes() } catch [System.Reflection.ReflectionTypeLoadException] { $types = $_.Exception.Types | Where-Object { $_ -ne $null } }
$types | Where-Object { $_.Name -like "*Query*" }
```

Reflection works but may miss embedded types or forwarded types. Compile test is always definitive.

---

## Build-Time Enforcement

Normal compilation enforces source safety through `ExcelSafetyAnalyzer`.
Core and ComInterop are completed typed areas, activated by
`excelmcp_typed_excel = true` in `.editorconfig`.

Each remaining executable `dynamic` occurrence in a completed area requires an
adjacent `// PIA gap:` comment on its containing declaration/statement, recording
reviewed, exact member-specific PIA/runtime or external-model dependency evidence.
Generic `TODO`/`Reason` comments do not exempt completed code; a gap comment does
not exempt dynamic numeric conversions. See [source safety](SOURCE-SAFETY.md)
for the bounded compiler policy and explicitly unfinished-area behavior.

Existing scripts, including `scripts\check-dynamic-casts.ps1`, remain
compiler-build compatibility adapters, not independent scanners.

---

## How This Happened (Root Cause)

During the original PIA migration, coverage was checked using binary string search with `[System.Text.Encoding]::Unicode` on the PIA DLL. This method reads the binary assembly as UTF-16LE text, which produces garbage. All string searches return false negatives — every searched type name was reported as "NOT FOUND" even when present.

This caused 8 DataModel and 3 Connection sub-type APIs to be incorrectly classified as "PIA gaps" when they are in fact available in the referenced Excel PIA.

**Prevention**: Always use a compile test (see above). This inventory records verified coverage and gaps; normal compilation enforces the [source-safety policy](SOURCE-SAFETY.md).
