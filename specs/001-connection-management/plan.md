# Implementation Plan: Connection Management

**Branch**: `001-connection-management` | **Date**: 2024-01-15 | **Status**: ✅ IMPLEMENTED  
**Spec**: [spec.md](spec.md)

## Summary

Connection Management provides full CRUD operations for Excel data connections (OLEDB, ODBC, Text, Web, XML, etc.) through COM interop. Enables version control of connection definitions, automated testing, and programmatic connection management. **Fully implemented across Core, CLI, and MCP Server layers.**

## Technical Context

### Tech Stack

- **.NET 9.0** - Target framework
- **Excel COM Interop** - Late binding via `dynamic`
- **System.Text.Json** - Connection definition serialization
- **Spectre.Console** - CLI table formatting
- **xUnit + FluentAssertions** - Testing framework

### Key Libraries

- `ExcelMcp.ComInterop` - Session management, COM utilities
- `ExcelMcp.Core.Connections` - Connection helpers (type detection, sanitization)
- `System.Text.Json` - JSON import/export

### Excel COM API

- `Workbook.Connections` collection
- `WorkbookConnection` object
- Type-specific: `OLEDBConnection`, `ODBCConnection`, `TextConnection`, `WebConnection`

## Project Structure

### Source Code (repository root)

```text
src/
├── ExcelMcp.Core/
│   ├── Commands/
│   │   └── Connection/
│   │       ├── IConnectionCommands.cs           # Interface (11 methods)
│   │       ├── ConnectionCommands.cs            # Partial: List, View
│   │       ├── ConnectionCommands.Lifecycle.cs   # Import, Export, Update, Delete
│   │       ├── ConnectionCommands.Operations.cs  # LoadTo, Test
│   │       └── ConnectionCommands.Refresh.cs     # Refresh with timeout
│   └── Connections/
│       └── ConnectionHelpers.cs                  # Type detection, sanitization
├── ExcelMcp.CLI/
│   └── Commands/
│       └── ConnectionCommands.cs                 # CLI wrapper with Spectre formatting
└── ExcelMcp.McpServer/
    └── Tools/
        └── ConnectionTool.cs                     # MCP action routing

tests/
└── ExcelMcp.Core.Tests/
    └── Commands/
        └── ConnectionCommandsTests.cs            # Integration tests
```

### Documentation (this feature)

```text
specs/001-connection-management/
├── spec.md              # This feature specification (what to build)
└── plan.md              # This implementation plan (how it's built)
```

## Architecture Decisions

### 1. Partial Class Organization

**Decision**: Split ConnectionCommands into 4 partial files by responsibility

**Files**:
- `ConnectionCommands.cs` - Base (List, View, GetProperties)
- `ConnectionCommands.Lifecycle.cs` - CRUD (Import, Export, Update, Delete)
- `ConnectionCommands.Operations.cs` - Operations (LoadTo, Test)
- `ConnectionCommands.Refresh.cs` - Refresh with timeout

**Rationale**:
- ✅ Each file 100-300 lines (readable, git-friendly)
- ✅ Clear separation of concerns
- ✅ Mirrors existing PowerQueryCommands pattern
- ✅ Parallel development without merge conflicts

### 2. Type 3/4 Handling Pattern

**Decision**: Handle TEXT (3) and WEB (4) interchangeably

**Implementation**:
```csharp
if (connType == 3 || connType == 4) {
    dynamic? textOrWeb = null;
    try { textOrWeb = conn.TextConnection; }
    catch { try { textOrWeb = conn.WebConnection; } catch { } }
    // Use textOrWeb...
}
```

**Rationale**:
- ✅ TEXT connections often report as WEB (Excel COM quirk)
- ✅ Handles both cases gracefully
- ✅ Documented in `excel-connection-types-guide.instructions.md`

### 3. Timeout Configuration

**Decision**: 2-minute default, 5-minute max for heavy operations

**Implementation**:
```csharp
public async Task<OperationResult> RefreshAsync(
    IExcelBatch batch, 
    string connectionName, 
    TimeSpan? timeout = null)
{
    return await batch.ExecuteAsync((ctx, ct) => {
        // Refresh logic
    }, timeout: timeout ?? TimeSpan.FromMinutes(2));
}
```

**Rationale**:
- ✅ Prevents indefinite hangs on slow data sources
- ✅ Configurable per operation
- ✅ Documented in `TIMEOUT-IMPLEMENTATION-GUIDE.md`

### 5. Test Strategy: TEXT Connections Only

**Decision**: Use TEXT file connections for automated tests

**Rationale**:
- ✅ No database dependencies
- ✅ Reliable creation via COM API (unlike OLEDB/ODBC)
- ✅ Fast test execution
- ❌ **Limitation**: OLEDB/ODBC creation unreliable, manage existing connections only

## Implementation Details

### Core Commands (11 Methods)

| Method | Purpose | Timeout | Batch API |
|--------|---------|---------|-----------|
| `ListAsync` | Enumerate all connections | No | Yes |
| `ViewAsync` | Get connection details | No | Yes |
| `ImportAsync` | Create from JSON | No | Yes |
| `ExportAsync` | Save to JSON | No | Yes |
| `UpdatePropertiesAsync` | Modify settings | No | Yes |
| `RefreshAsync` | Update data | 2-5 min | Yes |
| `DeleteAsync` | Remove connection | No | Yes |
| `LoadToAsync` | Create QueryTable | No | Yes |
| `GetPropertiesAsync` | Get refresh settings | No | Yes |
| `SetPropertiesAsync` | Update settings | No | Yes |
| `TestAsync` | Validate connectivity | 1-2 min | Yes |

### JSON Schema

```json
{
  "type": "OLEDB|ODBC|TEXT|WEB|...",
  "description": "string",
  "connectionString": "string (passwords excluded)",
  "commandText": "string",
  "commandType": "SQL|Table|...",
  "backgroundQuery": "bool",
  "refreshOnFileOpen": "bool",
  "savePassword": "bool (always false on export)"
}
```

### CLI Commands (11 Commands)

- `conn-list <file>`
- `conn-view <file> <name>`
- `conn-import <file> <name> <json>`
- `conn-export <file> <name> <json>`
- `conn-update <file> <name> <json>`
- `conn-refresh <file> <name>`
- `conn-delete <file> <name>`
- `conn-loadto <file> <name> <sheet>`
- `conn-properties <file> <name>`
- `conn-set-properties <file> <name> <json>`
- `conn-test <file> <name>`

### MCP Server Actions (11 Actions)

Excel Connection Tool (`excel_connection`) with actions:
- `List`, `View`, `Import`, `Export`, `UpdateProperties`
- `Refresh`, `Delete`, `LoadTo`
- `GetProperties`, `SetProperties`, `Test`

## Testing Approach

### Integration Tests (11 Test Methods)

**File**: `tests/ExcelMcp.Core.Tests/Commands/ConnectionCommandsTests.cs`

1. `List_EmptyWorkbook_ReturnsEmptyList`
2. `List_WithConnections_ReturnsAll`
3. `View_ValidConnection_ReturnsDetails`
4. `Export_ValidConnection_CreatesJsonFile`
5. `Import_ValidJson_CreatesConnection`
6. `UpdateProperties_ValidChanges_Persists`
7. `Refresh_ValidConnection_UpdatesData`
8. `Delete_ValidConnection_Removes`
9. `LoadTo_ValidConnection_CreatesQueryTable`
10. `GetProperties_ValidConnection_ReturnsSettings`
11. `Test_ValidConnection_ReturnsSuccess`

### Manual Test Scenarios

1. **OLEDB Connection Workflow**:
   - Create OLEDB connection in Excel UI to SQL Server
   - Export to JSON → Verify passwords excluded
   - Import to new workbook → Test connectivity
   - Refresh data → Verify updates

2. **TEXT Connection Workflow**:
   - Create TEXT connection to CSV file
   - List connections → Verify appears (may show as WEB type)
   - Update command text → Verify file path changes
   - Delete → Verify removed

## Known Limitations

### 1. Creation Limitations

**Issue**: `connections.Add()` for OLEDB/ODBC fails with "Value does not fall within expected range"

**Workaround**:
- ✅ Manage existing connections created via Excel UI
- ✅ Import from .odc files (user creates in Excel)
- ✅ TEXT connections work reliably for testing

**Documented**: `excel-connection-types-guide.instructions.md`

### 2. Type 3/4 Ambiguity

**Issue**: TEXT connections (Type 3) often report as WEB (Type 4)

**Workaround**: Handle both types interchangeably in code

**Documented**: `excel-connection-types-guide.instructions.md`

### 3. Password Recovery

**Issue**: Cannot retrieve existing passwords from Excel (by design)

**Workaround**: Not a bug - security feature, passwords never exported

## Quality Gates

### Pre-Commit Checks

- ✅ Build passes (0 warnings)
- ✅ Integration tests pass (11 tests)
- ✅ COM leak detection: `scripts\check-com-leaks.ps1`
- ✅ Success flag validation: `scripts\check-success-flag.ps1`

### Code Review Checklist

- ✅ All Core methods have XML doc comments
- ✅ Error messages include connection name
- ✅ Passwords excluded from logs and exports
- ✅ COM objects released in finally blocks
- ✅ Batch API pattern used consistently
- ✅ Integration tests for each method

## Related Features

- **Power Query** (`001-power-query`) - Similar CRUD pattern for queries
- **QueryTable** (`007-querytable`) - LoadTo creates QueryTables from connections
- **Data Model** (`002-data-model`) - Connections can load to Data Model

## Migration Notes

**No breaking changes** - This is a new feature, no migration required.

**Version**: Added in v1.2.0 (2024-01-15)
