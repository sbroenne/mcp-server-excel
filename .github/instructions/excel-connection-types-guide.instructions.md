---
applyTo: "src/ExcelMcp.Core/Commands/ConnectionCommands.cs,src/ExcelMcp.Core/Connections/**/*.cs,tests/**/ConnectionCommands*.cs,tests/**/ConnectionTestHelper.cs"
---

# Excel Connection Types - LLM Quick Reference

> **What works, what doesn't, and what to do instead**

## CRITICAL: LoadTo Operation Limitations

**LoadTo action only works with OLEDB/ODBC connections:**

| Connection Type | LoadTo Support | What to Use Instead |
|----------------|----------------|---------------------|
| OLEDB | Works | Primary use case |
| ODBC | Works | Primary use case |
| TEXT | FAILS | Use powerquery create + refresh |
| WEB | FAILS | Use powerquery create + refresh |
| Power Query | Works | Use powerquery refresh |

**Error pattern:** If LoadTo returns "Value does not fall within the expected range" then connection type doesn't support QueryTable pattern - use Power Query instead.

## Connection Action Compatibility

| Action | OLEDB/ODBC | TEXT | WEB | Power Query |
|--------|-----------|------|-----|-------------|
| List | Works | Works | Works | Works |
| View | Works | Works | Works | Works |
| Create | Works | Works | Works | Use powerquery |
| Delete | Works | Works | Works | Use powerquery |
| LoadTo | Works | FAILS | FAILS | Use powerquery refresh |
| Refresh | Works | Works* | Works* | Use powerquery refresh |
| Test | Works | Works | Works | Works |

*TEXT/WEB Refresh succeeds but doesn't validate data source existence until actual data access

## Decision Tree: Connection vs Power Query

```
Need to import data from file/URL?
├─ OLEDB/ODBC data source?
│  └─ Use connection (LoadTo, Refresh)
│
├─ TEXT file (CSV, TXT)?
│  └─ Use powerquery (create with M code, refresh)
│
├─ Web API/URL?
│  └─ Use powerquery (create with M code, refresh)
│
└─ Already has Power Query?
   └─ Use powerquery (refresh)
```

## Recommended Workflows

**OLEDB/ODBC Data Loading:**
```
1. connection create → Creates connection object
2. connection loadto → Loads data to worksheet
3. connection refresh → Updates data from source
```

**TEXT/CSV File Import:**
```
1. powerquery create → Import CSV with M code
2. powerquery refresh → Reload data
   (Don't use connection loadto - will fail!)
```

**Web Data Import:**
```
1. powerquery create → Import from URL with M code
2. powerquery refresh → Update data
   (Don't use connection loadto - will fail!)
```

## Common Mistakes to Avoid

1. **Using LoadTo with TEXT connections** - Will fail with E_INVALIDARG - Use Power Query instead
2. **Using LoadTo with WEB connections** - Will fail - Use Power Query instead
3. **Assuming Refresh validates TEXT file existence** - Excel doesn't check until data access
4. **Mixing connection and Power Query operations** - Power Query connections need powerquery tool

## Connection String Examples

```
OLEDB:  "Provider=SQLOLEDB;Data Source=server;Initial Catalog=db;..."
ODBC:   "DSN=MyDataSource;UID=username;PWD=password;..."
TEXT:   "TEXT;C:\\path\\to\\file.csv"
WEB:    "URL;https://example.com/data.xml"
```

## Security

**Always sanitize connection strings before displaying** - Never expose passwords or sensitive credentials in error messages or logs.

---

## Developer Reference (Implementation Details)

<details>
<summary>Click to expand developer implementation notes</summary>

### Implementation Notes

**Connections.Add2() method required for OLEDB/ODBC:**

Use the COM Add2 method with parameters: Name, Description, ConnectionString, CommandText (empty), lCmdtype (auto-detect), CreateModelConnection (false), ImportRelationships (false).

### Type 3/4 Ambiguity

TEXT connections created with "TEXT;path" may return type 4 (WEB) instead of 3 (TEXT) - handle both types interchangeably in type detection logic.

### Test Strategy

- **OLEDB** - Use for LoadTo, Refresh, and QueryTable operation tests
- **TEXT** - Use for connection lifecycle tests (List, View, Delete) without LoadTo
- **ODBC** - Use for validation of multiple connection types

### Connection String Internal Formats

```
OLEDB:        "Provider=SQLOLEDB;Data Source=server;Initial Catalog=db;..."
ODBC:         "DSN=MyDataSource;UID=username;PWD=password;..."
TEXT:         "TEXT;C:\\path\\to\\file.csv"
WEB:          "URL;https://example.com/data.xml"
Power Query:  "OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=QueryName"
```

</details>
