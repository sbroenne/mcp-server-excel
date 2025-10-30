---
applyTo: "src/ExcelMcp.Core/Commands/ConnectionCommands.cs,src/ExcelMcp.Core/Connections/**/*.cs,tests/**/ConnectionCommands*.cs,tests/**/ConnectionTestHelper.cs"
---

# Excel Connection Types - Essential Patterns

> **Critical patterns for working with Excel connections**

## Connection Type Enum

[Official docs](https://learn.microsoft.com/en-us/office/vba/api/excel.xlconnectiontype): Types 1-9 (OLEDB, ODBC, TEXT, WEB, XMLMAP, DATAFEED, MODEL, WORKSHEET, NOSOURCE)

## COM API Limitations

### ❌ Don't Create Connections via COM

```csharp
// OLEDB/ODBC creation fails with "Value does not fall within expected range"
connections.Add(Name: "Test", ConnectionString: oledbString);  // ❌ FAILS
```

**Workaround:**
- Users create via Excel UI (Data → Get Data)
- Users import from .odc files
- ConnectionCommands manages existing connections only

### ✅ Use TEXT Connections for Testing

```csharp
// TEXT connections work reliably
string connectionString = $"TEXT;{csvFilePath}";
dynamic conn = connections.Add(
    Name: "TestText",
    Description: "Test",
    ConnectionString: connectionString,
    CommandText: ""
);  // ✅ WORKS
```

## Type 3/4 Handling Pattern

**Issue:** TEXT connections created with `"TEXT;path"` return type 4 (WEB) instead of 3 (TEXT).

**Solution:** Handle both types interchangeably:

```csharp
// Try TEXT first, fall back to WEB
if (connType == 3 || connType == 4)  // Handle both
{
    dynamic? textOrWeb = null!;
    try { textOrWeb = conn.TextConnection; }
    catch { try { textOrWeb = conn.WebConnection; } catch { return defaultValue; } }
    
    if (textOrWeb != null)
        return textOrWeb.SomeProperty ?? defaultValue;
}
```

## Connection String Formats

```csharp
// OLEDB
"Provider=SQLOLEDB;Data Source=server;Initial Catalog=db;"

// ODBC
"DSN=MyDataSource;UID=username;PWD=password;"

// TEXT
"TEXT;C:\\path\\to\\file.csv"

// WEB
"URL;https://example.com/data.xml"

// Power Query
"OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=QueryName"
```

## Testing Strategy

**Use TEXT connections for all automated tests:**

```csharp
public static async Task CreateTextFileConnectionAsync(
    string filePath, string connectionName, string csvFilePath)
{
    await using var batch = await ExcelSession.BeginBatchAsync(filePath);
    await batch.ExecuteAsync<int>((ctx, ct) =>
    {
        dynamic connections = ctx.Book.Connections;
        
        if (!File.Exists(csvFilePath))
            File.WriteAllText(csvFilePath, "Col1,Col2\nVal1,Val2\n");
        
        string connectionString = $"TEXT;{csvFilePath}";
        dynamic conn = connections.Add(
            Name: connectionName,
            Description: "Test text connection",
            ConnectionString: connectionString,
            CommandText: ""
        );
        
        return ValueTask.FromResult(0);
    });
    await batch.SaveAsync();
}
```

## Security

```csharp
// ALWAYS sanitize before displaying
string safe = ConnectionHelpers.SanitizeConnectionString(rawConnectionString);
Console.WriteLine(safe);  // Passwords masked

// NEVER save passwords by default
oledb.SavePassword = false;
```

## Type-Specific APIs

```csharp
// Each connection type has specific COM object
if (conn.Type == 1)  // OLEDB
{
    dynamic oledb = conn.OLEDBConnection;
    oledb.BackgroundQuery = true;
    oledb.RefreshOnFileOpen = false;
}
else if (conn.Type == 3 || conn.Type == 4)  // TEXT or WEB
{
    // Handle both due to Excel quirk
    dynamic? textOrWeb = null!;
    try { textOrWeb = conn.TextConnection; }
    catch { textOrWeb = conn.WebConnection; }
    // Use textOrWeb...
}
```

## Key Takeaways

1. **OLEDB/ODBC creation unreliable** - manage existing connections only
2. **TEXT connections work** - use for testing
3. **Type 3/4 ambiguity** - handle both interchangeably
4. **Always sanitize** - never expose passwords
5. **Test with TEXT** - no DB dependencies
