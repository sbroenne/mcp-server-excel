# PIA Coverage Guide

This document tracks the status of `Microsoft.Office.Interop.Excel` (v16) type coverage across ExcelMcp, explains why some casts remain `dynamic`, and documents the verified methodology for checking PIA coverage.

---

## TL;DR — Current Status

| Area | Status |
|------|--------|
| Core Excel types (`Workbook`, `Worksheet`, `Range`, etc.) | ✅ Fully typed |
| Collections (`Sheets`, `Names`, `ListObjects`, etc.) | ✅ Fully typed |
| Power Query (`Workbook.Queries`, `WorkbookQuery`) | ✅ Fully typed |
| DataModel (`Model`, `ModelMeasures`, `ModelTables`, etc.) | ✅ Migrated — all model types use PIA |
| Connection sub-types (`OLEDBConnection`, `ODBCConnection`, `TextConnection`) | ✅ Migrated — callers use typed WorkbookConnection |
| `ModelTableColumn.IsCalculatedColumn` | ❌ True PIA gap — property missing from v16 PIA |
| `ModelMeasure.FormatInformation` | ⚠️ Returns `object` in PIA; cast to `dynamic` for property probing |
| VBA (`VBProject`, `VBComponents`) | ❌ True PIA gap — in `Microsoft.Vbe.Interop` only |
| `AutomationSecurity` | ❌ True PIA gap — in `Microsoft.Office.Core` (office.dll); use `((dynamic)(object))` cast |
| WebConnection | ❌ True PIA gap — not in Excel PIA |
| ADO types (ADODB.Connection, Recordset, Fields) | ❌ True PIA gap — in ADODB, not Excel PIA |

---

## True PIA Gaps

These APIs are **not** in `Microsoft.Office.Interop.Excel` and will remain `dynamic` permanently.

### VBProject / VBComponents / VBComponent

- **Location in COM**: `Microsoft.Vbe.Interop` (separate DLL — `vbe7.dll` / `VBA7.dll`)
- **Why not available**: No .NET 5+ compatible NuGet package exists. `ThammimTech.Microsoft.Vbe.Interop` targets .NET Framework only. The official `Microsoft.Vbe.Interop` NuGet package is unsigned and unmaintained.
- **Workaround**: `((dynamic)ctx.Book).VBProject` — COM late binding to the VBE object model
- **Affected files**: `VbaCommands.Lifecycle.cs`, `VbaCommands.Operations.cs`

### MsoAutomationSecurity (= 3)

- **Location in COM**: `Microsoft.Office.Core` (office.dll / `Microsoft.Office.Interop.Word` / `Microsoft.Office.Interop.PowerPoint` host DLLs)
- **Why not available**: The `office.dll` shared types are injected via `[assembly: PrimaryInteropAssembly]` into host Office PIAs. They are not directly available as a standalone typed constant in the Excel PIA.
- **Workaround**: Use the literal integer value `3` (= `msoAutomationSecurityForceDisable`) via `((dynamic)(object)tempExcel).AutomationSecurity = 3;`
- **CRITICAL — cast to `(object)` first**: Casting a typed `Excel.Application` directly to `dynamic` retains COM type metadata; the DLR then tries to load `office.dll` to resolve `MsoAutomationSecurity`, causing a `FileNotFoundException` crash (`office, Version=16.0.0.0`). Casting to `(object)` first erases the static type and forces pure IDispatch binding, which never loads `office.dll`.
- **Do NOT add a `<Reference>` to office.dll**: A GAC hint path is version-specific and machine-specific (15.0 vs 16.0 mismatch causes the same crash). `EmbedInteropTypes=true` + `(object)` cast is sufficient.
- **Affected files**: `ExcelBatch.cs`

### WebConnection

- **Simply not in the Excel PIA**: `WorkbookConnection.WebConnection` is not exposed in `Microsoft.Office.Interop.Excel` v16. Only `OLEDBConnection`, `ODBCConnection`, and `TextConnection` have typed sub-connection properties.
- **Affected files**: `ConnectionCommands.cs` (`GetTypedSubConnection` method)

---

## TODO — Types IN the PIA, Migration Pending

These were incorrectly left as `dynamic` during the initial PIA migration due to a false negative from binary string search on the assembly. All have been confirmed by compile test.

### Power Query — `Excel.Queries` / `Excel.WorkbookQuery`

- **Compile test result**: `Excel.Queries q = book.Queries;` compiles cleanly with the v16 PIA
- **WorkbookQuery properties available**: `.Name`, `.Formula`
- **Why left as dynamic**: Binary inspection using `Encoding.Unicode` (incorrect) reported these types as absent. The true method (reflection / compile test) was not used.
- **Affected files**: `PowerQueryCommands.Create.cs`, `Evaluate.cs`, `Lifecycle.cs` (2x), `Refresh.cs`, `Rename.cs`, `Update.cs`, `LoadTo.cs`, `View.cs`, `ComUtilities.cs`
- **Migration effort**: Medium — replace 9 occurrences of `((dynamic)ctx.Book).Queries` + update `ComUtilities.FindQuery()` signature

### DataModel — `Excel.Model`, `Excel.ModelMeasures`, `Excel.ModelTables`, etc.

- **Types confirmed in PIA**:
  - `Excel.Model` (`workbook.Model`)
  - `Excel.ModelMeasures`, `Excel.ModelMeasure`
  - `Excel.ModelTables`, `Excel.ModelTable`
  - `Excel.ModelRelationships`, `Excel.ModelRelationship`
  - `Excel.ModelTableColumns`, `Excel.ModelTableColumn`
  - Format types: `ModelFormatGeneral`, `ModelFormatCurrency`, `ModelFormatDecimalNumber`, `ModelFormatDate`, `ModelFormatBoolean`, `ModelFormatPercentageNumber`, `ModelFormatWholeNumber`, `ModelFormatScientificNumber`
- **Affected files**: `DataModelCommands.Helpers.cs` and all DataModel command files
- **Migration effort**: High — `DataModelCommands.Helpers.cs` is entirely `dynamic` internally

### Connection Sub-Types — `Excel.OLEDBConnection`, `Excel.ODBCConnection`, `Excel.TextConnection`

- **Compile test result**: All three sub-connection types compile cleanly
- **Note**: The collection is `Excel.Connections` (NOT `WorkbookConnections`)
- **Affected files**: `ConnectionCommands.cs` (`GetTypedSubConnection` helper)
- **Migration effort**: Low — one helper method

### ComUtilities Helpers — Return Types

- `FindQuery(dynamic workbook, ...)` — both parameter and return type can be typed once Queries migration is done
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
- `// PIA gap: ...` — Type is a true gap (not in v16 PIA)
- `// TODO: ...` — Type IS in PIA but migration is pending
- `// Reason: ...` — Other documented reason

The check is run automatically by `scripts/pre-commit.ps1`.

---

## How This Happened (Root Cause)

During the original PIA migration, coverage was checked using binary string search with `[System.Text.Encoding]::Unicode` on the PIA DLL. This method reads the binary assembly as UTF-16LE text, which produces garbage. All string searches return false negatives — every searched type name was reported as "NOT FOUND" even when present.

This caused 9 Power Query, 8 DataModel, and 3 Connection sub-type APIs to be incorrectly classified as "PIA gaps" when they are in fact available in the v16 PIA.

**Prevention**: Always use a compile test (see above). This file and the `check-dynamic-casts.ps1` pre-commit check prevent silent regression.
