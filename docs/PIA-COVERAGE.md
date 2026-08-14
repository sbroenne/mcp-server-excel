# PIA Coverage Guide

This document tracks the status of `Microsoft.Office.Interop.Excel` type coverage across ExcelMcp, explains why some casts remain `dynamic`, and documents the verified methodology for checking PIA coverage.

---

## TL;DR — Current Status

| Area | Status |
|------|--------|
| Core Excel types (`Workbook`, `Worksheet`, `Range`, etc.) | ✅ Fully typed |
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
| `ModelTableColumn.DataType` / `XlParameterDataType` | ✅ Typed; list/read responses expose raw and named values |
| DataModel measures/formats (`ModelMeasures`, `ModelMeasure`, `ModelFormat*`) | ✅ Typed via the 16.x Excel PIA |
| Connection sub-types (`OLEDBConnection`, `ODBCConnection`, `TextConnection`) | ✅ Migrated — callers use typed WorkbookConnection |
| `ModelTableColumn.IsCalculatedColumn` / formula / expression | ❌ Not exposed by the 16.x Excel PIA; runtime COM also has no reliable calculated-column metadata API |
| `ModelTable.LastRefresh` | ❌ Missing from the 16.x Excel PIA and unavailable at runtime; no reliable table refresh timestamp |
| Data Model source/model connection metadata | ✅ Typed (`SourceWorkbookConnection`, `DataModelConnection`, `ModelConnection`, `ModelTables`, command metadata) |
| Data Model live refresh status | ❌ No `Model.Refreshing` or `ModelTable.Refreshing` PIA/COM surface; refresh calls are synchronous |
| `ModelMeasure.FormatInformation` | ⚠️ Typed as `object`; dynamic property probing remains for polymorphic `ModelFormat*` objects |
| `WorkbookConnection.Refreshing` / `CancelRefresh` | ⚠️ Still missing from the 16.x Excel PIA — dynamic debt |
| VBA (`VBProject`, `VBComponents`) | ⚠️ External VBE object model, not Excel PIA — dynamic debt until typed interop is added |
| `AutomationSecurity` | ⚠️ Office.Core enum (office.dll); accessed late-bound via `((dynamic)(object))` because Office.Core is not referenced. The embedded Excel PIA carries no office.dll dependency, so no resolver is needed |
| WebConnection | ⚠️ Still missing from the 16.x Excel PIA — dynamic debt |
| ADO types (ADODB.Connection, Recordset, Fields) | ⚠️ External ADODB object model, not Excel PIA — dynamic debt until typed interop is added |

---

## Remaining Dynamic Debt

ExcelMcp is PIA-first. Any `dynamic` usage is technical debt unless a compile probe proves the referenced Excel PIA lacks the API, or the API belongs to an external COM object model that is not currently referenced. Prefer adding typed interop coverage over spreading new `dynamic` calls.

### Power Query — `Workbook.Queries` / `WorkbookQuery`

- **Status**: Typed via `Microsoft.Office.Interop.Excel` 16.x.
- **Affected files**: `ComUtilities.cs` and Power Query command files now use `Excel.Queries` and `Excel.WorkbookQuery`.

### DataModel measures/formats — `ModelMeasures` / `ModelMeasure` / `ModelFormat*`

- **Status**: Typed via `Microsoft.Office.Interop.Excel` 16.x.
- **Affected files**: Data Model measure commands now use `Excel.ModelMeasures`, `Excel.ModelMeasure`, and typed `ModelFormat*` properties.

### WorkbookConnection refresh status/cancel

- **Why not typed yet**: `WorkbookConnection.Refreshing` and `WorkbookConnection.CancelRefresh` do not compile against `Microsoft.Office.Interop.Excel` 16.0.18925.20022.
- **Current workaround**: Dynamic calls are isolated to connection refresh waiting/cancellation.
- **Preferred future fix**: Add a typed interop surface or replace the behavior with a typed Excel API path.

### Data Model calculated columns

- **Compile probe result**: `ModelTableColumn.IsCalculatedColumn`, `Formula`, and `Expression` all fail with CS1061 against `Microsoft.Office.Interop.Excel` 16.0.18925.20022.
- **Runtime result**: Excel reliably exposes only column name, data type, and parent through `ModelTableColumn`; calculated-column formulas and mutation are not practical automation surfaces.
- **Current behavior**: `list-columns` returns reliable names plus the raw and named `XlParameterDataType` values. Use Power Query for computed columns or a DAX measure for query-time calculations.
- **Runtime quirk**: Excel returns value `20` (`BIGINT`) for worksheet integer columns loaded into the model even though `20` is outside the documented `XlParameterDataType` values. ExcelMcp preserves the raw value and supplies the observed readable name.

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

- **Simply not in the Excel PIA**: `WorkbookConnection.WebConnection` is not exposed in `Microsoft.Office.Interop.Excel`. Only `OLEDBConnection`, `ODBCConnection`, and `TextConnection` have typed sub-connection properties.
- **Affected files**: `ConnectionCommands.cs` (`GetTypedSubConnection` method)

---

## Migrated Typed Surfaces

These were incorrectly left as `dynamic` during earlier PIA migration work due to false negatives from binary string search on the assembly. They are now typed or should remain typed.

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
- **Migration effort**: Low — one helper method

### ComUtilities Helpers — Return Types

- `FindQuery(Excel.Workbook workbook, ...)` returns `Excel.WorkbookQuery?`
- `FindConnection(dynamic workbook, ...)` — return type can be `Excel.WorkbookConnection?`
- `FindSheet(...)` — return type can be `Excel.Worksheet?`
- `FindName(...)` — return type can be `Excel.Name?`
- `ResolveRange()` in `RangeHelpers.cs` — return type can be `Excel.Range?`

---

## Verified Methodology for Checking PIA Coverage

### THE CORRECT METHOD: Compile test

Create a probe project **outside the repo** (to avoid Central Package Management interference):

```powershell
mkdir D:\temp\pia-probe
cd D:\temp\pia-probe
dotnet new console
dotnet add package Microsoft.Office.Interop.Excel --version 15.0.4795.1000
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

If it compiles → the type is in the PIA. If CS0246 → it's not.

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

## Pre-Commit Enforcement

The `scripts/check-dynamic-casts.ps1` script enforces that every `((dynamic))` cast has a justification comment on the preceding line.

Valid comment prefixes:
- `// PIA gap: ...` — Type is a true gap (not in the referenced Excel PIA)
- `// TODO: ...` — Type IS in PIA but migration is pending
- `// Reason: ...` — Other documented reason

The check is run automatically by `scripts/pre-commit.ps1`.

---

## How This Happened (Root Cause)

During the original PIA migration, coverage was checked using binary string search with `[System.Text.Encoding]::Unicode` on the PIA DLL. This method reads the binary assembly as UTF-16LE text, which produces garbage. All string searches return false negatives — every searched type name was reported as "NOT FOUND" even when present.

This caused 8 DataModel and 3 Connection sub-type APIs to be incorrectly classified as "PIA gaps" when they are in fact available in the referenced Excel PIA.

**Prevention**: Always use a compile test (see above). This file and the `check-dynamic-casts.ps1` pre-commit check prevent silent regression.
