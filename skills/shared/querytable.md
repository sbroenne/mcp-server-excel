# querytable - Local Text and Web Imports

Use `querytable` for worksheet QueryTables backed by the desktop Excel COM object model.

## Choose the Right Import Surface

| Need | Tool |
|------|------|
| Direct text/CSV import with delimiter and encoding control | `querytable create-text` |
| Legacy HTML page/table import through Excel's web-query engine | `querytable create-web` |
| Modern connectors, transformations, APIs, JSON, or reusable M | `powerquery` |
| Existing OLEDB/ODBC workbook connection | `connection` |

## Text Import

```mcp
querytable(action: 'create-text', session_id: sessionId,
    query_table_name: 'OrdersCsv',
    source_path: 'C:\Data\orders.csv',
    sheet_name: 'Orders',
    destination_address: 'A1',
    delimiter: ',',
    text_qualifier: 'double-quote',
    encoding: 65001,
    has_headers: true)
```

```cli
excelcli -q querytable create-text --session $sessionId --query-table-name OrdersCsv --source-path $sourcePath --sheet Orders --destination-address A1 --delimiter ',' --text-qualifier double-quote --encoding 65001 --has-headers true
```

Use a known readable source path and an existing destination sheet. The example
names are illustrative; use the actual target and inspect occupied cells first.

- `delimiter` is exactly one character.
- `text_qualifier` is `double-quote`, `single-quote`, or `none`.
- `encoding` is a Windows code page; use `65001` for UTF-8.
- Creation refreshes synchronously so imported data is ready when the call returns.

## Legacy Web Import

```mcp
querytable(action: 'create-web', session_id: sessionId,
    query_table_name: 'RatesHtml',
    url: 'https://example.com/rates.html',
    sheet_name: 'Rates',
    destination_address: 'A1',
    selection_type: 'specified-tables',
    web_tables: '1',
    formatting: 'none')
```

```cli
excelcli -q querytable create-web --session $sessionId --query-table-name RatesHtml --url $sourceUrl --sheet Rates --destination-address A1 --selection-type specified-tables --web-tables '1' --formatting none
```

The source URL must identify the user's intended HTML page, not a guessed site.

- `selection_type` is `entire-page`, `all-tables`, or `specified-tables`.
- `web_tables` is required with `specified-tables`.
- `formatting` is `none`, `rich-text`, or `all`.
- This is Excel's legacy HTML web-query engine, not a general HTTP or browser automation API.

## Lifecycle and Refresh

Use `list`, `view`, `set-properties`, `refresh`, `get-refresh-status`, `cancel-refresh`, and `delete` for existing QueryTables.

## Hard Exclusions

Local QueryTable COM automation cannot access Microsoft 365 cloud service state or APIs:

- No workbook sharing or permissions
- No coauthor presence, cursors, conflicts, or live collaboration state
- No comment @mentions, assignments, reactions, or notification delivery
- No authenticated Graph, SharePoint, Teams, or OneDrive service operations
- No Power Query M definition or modern connector configuration

Use the relevant Microsoft 365 service API for cloud workflows and `powerquery` for modern data transformation.
